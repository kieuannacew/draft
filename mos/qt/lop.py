"""Lớp học trực tuyến: trang Lớp học (giáo viên / học sinh), Tài khoản máy chủ (quản trị),
hộp thoại Đăng ký bằng mã lớp và Cài đặt máy chủ."""
from __future__ import annotations

import re
from datetime import datetime

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (QApplication, QComboBox, QFileDialog, QGridLayout, QHBoxLayout, QLineEdit,
                               QVBoxLayout)

from .. import lop_hoc
from ..accounts import ROLE_NAMES, USERNAME_RE, random_password
from ..i18n import pick, role_name, tr
from . import theme as T
from .admin import _Dialog, _Page, _toolbar
from .app import _banner, esc, mode_name
from .theme import PRIMARY_SOFT, SUCCESS, WARN, Card, button, label

ERR_PREFIX = "Máy chủ báo lỗi: "


def err_text(exc: Exception) -> str:
    msg = str(exc)
    if msg.startswith(ERR_PREFIX):
        return tr("Máy chủ báo lỗi: {err}").format(err=msg[len(ERR_PREFIX):])
    return tr(msg)


def busy(parent, fn, *args, **kw):
    """Chạy thao tác mạng với con trỏ chờ; lỗi → hộp thoại, trả None."""
    QApplication.setOverrideCursor(Qt.WaitCursor)
    try:
        return fn(*args, **kw)
    except lop_hoc.CloudError as exc:
        QApplication.restoreOverrideCursor()
        T.error(parent, tr("Lỗi"), err_text(exc))
        return None
    finally:
        if QApplication.overrideCursor() is not None:
            QApplication.restoreOverrideCursor()


def act(parent, fn, *args) -> bool:
    """Thao tác mạng không cần kết quả: True nếu thành công (lỗi đã hiện hộp thoại)."""
    QApplication.setOverrideCursor(Qt.WaitCursor)
    try:
        fn(*args)
        return True
    except lop_hoc.CloudError as exc:
        QApplication.restoreOverrideCursor()
        T.error(parent, tr("Lỗi"), err_text(exc))
        return False
    finally:
        if QApplication.overrideCursor() is not None:
            QApplication.restoreOverrideCursor()


def local_time(value: str) -> str:
    if not value:
        return "—"
    try:
        return datetime.fromisoformat(value).astimezone().strftime("%Y-%m-%d %H:%M")
    except ValueError:
        return value[:16].replace("T", " ")


def exam_title(r: dict) -> str:
    return pick(r.get("exam", ""), r.get("exam_en"))


# ====================================================================== cài đặt máy chủ


class ServerDialog(_Dialog):
    """Nhập địa chỉ + khóa Supabase của lớp học (quản trị làm một lần, rồi gửi kèm cau_hinh.json cho học sinh)."""

    def __init__(self, parent):
        super().__init__(parent, tr("Máy chủ lớp học"), 560)
        cur = lop_hoc.load_server() or {"url": "", "key": ""}
        self.lay.addWidget(label(tr("Dùng Supabase (miễn phí) để giáo viên xem điểm học sinh ở mọi nơi. Cách tạo máy "
                                    "chủ xem file HUONG_DAN_MAY_CHU.md. Để trống và bấm “Tắt máy chủ” để dùng "
                                    "ngoại tuyến như cũ."), "muted", wrap=True))
        self.url = self.field(tr("Địa chỉ (Project URL)"), QLineEdit(cur["url"]))
        self.url.setPlaceholderText("https://abcdxyz.supabase.co")
        self.key = self.field(tr("Khóa công khai (anon hoặc publishable key)"), QLineEdit(cur["key"]),
                              tr("Supabase → Project Settings → API. Khóa này được phép để lộ; dữ liệu được bảo vệ "
                                 "bằng quyền truy cập trên máy chủ. KHÔNG dùng khóa service_role / secret."))
        self.status = label("", wrap=True)
        self.lay.addWidget(self.status)
        row = QHBoxLayout()
        row.addWidget(button(tr("Kiểm tra kết nối"), self.check))
        if lop_hoc.load_server():
            row.addWidget(button(tr("Tắt máy chủ"), self.turn_off, "danger"))
        row.addStretch()
        row.addWidget(button(tr("Hủy"), self.reject))
        row.addWidget(button(tr("Lưu"), self.save, "primary"))
        self.lay.addSpacing(6)
        self.lay.addLayout(row)

    def _cloud(self):
        url, key = self.url.text().strip().rstrip("/"), self.key.text().strip()
        if not url or not key:
            self._say(tr("Hãy nhập đủ địa chỉ và khóa."), T.DANGER)
            return None
        if not url.startswith(("http://", "https://")):
            url = "https://" + url
        return lop_hoc.Cloud(url, key, timeout=10)

    def _say(self, text, color):
        self.status.setText(text)
        self.status.setStyleSheet(f"color:{color};")

    def check(self) -> bool:
        cloud = self._cloud()
        if cloud is None:
            return False
        QApplication.setOverrideCursor(Qt.WaitCursor)
        try:
            info = cloud.check()
        except lop_hoc.CloudError as exc:
            self._say("✗ " + err_text(exc), T.DANGER)
            return False
        finally:
            QApplication.restoreOverrideCursor()
        if not info["schema"]:
            self._say("✗ " + tr("Kết nối được nhưng chưa cài đặt: hãy chạy file may_chu/supabase_lop_hoc.sql trong "
                                "Supabase → SQL Editor."), T.DANGER)
            return False
        notes = []
        if not info["autoconfirm"]:
            notes.append(tr("Hãy tắt “Confirm email” (Authentication → Sign In / Providers → Email), nếu không học "
                            "sinh không đăng ký được."))
        if not info["signup"]:
            notes.append(tr("Máy chủ đang tắt đăng ký tài khoản mới."))
        if notes:
            self._say("⚠ " + " ".join(notes), WARN)
        else:
            self._say("✓ " + tr("Kết nối tốt, máy chủ đã sẵn sàng."), SUCCESS)
        return True

    def save(self):
        if not self.check():
            if not T.confirm(self, tr("Máy chủ lớp học"), tr("Chưa kiểm tra được máy chủ. Vẫn lưu?")):
                return
        path = lop_hoc.save_server(self.url.text(), self.key.text())
        T.info(self, tr("Đã lưu"), tr("Đã lưu vào:\n{path}\n\nGửi kèm file này (đặt cạnh run.bat) cho học sinh để app "
                                      "tự kết nối máy chủ lớp học.").format(path=path))
        self.accept()

    def turn_off(self):
        if T.confirm(self, tr("Tắt máy chủ"), tr("Tắt máy chủ lớp học và dùng tài khoản trên máy này như cũ?")):
            lop_hoc.clear_server()
            self.accept()


# ====================================================================== đăng ký


class RegisterDialog(_Dialog):
    """Học sinh tự tạo tài khoản trên máy chủ và vào lớp bằng mã lớp."""

    def __init__(self, parent, server: dict):
        super().__init__(parent, tr("Đăng ký tài khoản"), 480)
        self.server = server
        self.result_data = None
        self.lay.addWidget(label(tr("Nhập mã lớp giáo viên gửi để điểm của bạn tự gửi cho giáo viên."), "muted",
                                 wrap=True))
        self.ho_ten = self.field(tr("Họ tên"), QLineEdit())
        self.ho_ten.setPlaceholderText(tr("vd: Nguyễn Văn An"))
        self.user = self.field(tr("Tên đăng nhập"), QLineEdit(),
                               tr("Chữ thường không dấu, số, dấu _ hoặc . (vd: nguyenvana)"))
        self.pw = self.field(tr("Mật khẩu (ít nhất 6 ký tự)"), QLineEdit())
        self.pw.setEchoMode(QLineEdit.Password)
        self.again = self.field(tr("Nhập lại mật khẩu"), QLineEdit())
        self.again.setEchoMode(QLineEdit.Password)
        self.code = self.field(tr("Mã lớp"), QLineEdit(), tr("6 ký tự, vd K7M2QX. Giáo viên có thể để trống."))
        self.code.setMaxLength(6)
        self.code.textEdited.connect(lambda t: self.code.setText(t.upper()))
        self.err = label("", wrap=True)
        self.err.setStyleSheet(f"color:{T.DANGER};")
        self.lay.addWidget(self.err)
        self.ho_ten.textEdited.connect(self._suggest)
        self.buttons(tr("Đăng ký"), self.register)

    def _suggest(self, name):
        if self.user.isModified():
            return
        import unicodedata
        text = unicodedata.normalize("NFKD", name.replace("đ", "d").replace("Đ", "D"))
        text = "".join(c for c in text if not unicodedata.combining(c)).lower()
        self.user.setText(re.sub(r"[^a-z0-9]", "", text)[:32])
        self.user.setModified(False)

    def register(self):
        username = self.user.text().strip().lower()
        if not self.ho_ten.text().strip():
            return self.err.setText(tr("Hãy nhập họ tên."))
        if not USERNAME_RE.fullmatch(username):
            return self.err.setText(tr("Tên đăng nhập 3–32 ký tự, chỉ gồm chữ thường không dấu, số, dấu _ hoặc ."))
        if len(self.pw.text()) < lop_hoc.MIN_PASSWORD:
            return self.err.setText(tr("Mật khẩu phải có ít nhất 6 ký tự."))
        if self.pw.text() != self.again.text():
            return self.err.setText(tr("Hai lần nhập mật khẩu không giống nhau."))
        code = self.code.text().strip().upper()
        if code and len(code) != 6:
            return self.err.setText(tr("Mã lớp gồm 6 ký tự."))
        cloud = lop_hoc.Cloud(**self.server)
        QApplication.setOverrideCursor(Qt.WaitCursor)
        joined, join_err = None, None
        try:
            profile = cloud.signup(username, self.pw.text(), self.ho_ten.text())
            if code:
                try:
                    joined = cloud.join_class(code)
                except lop_hoc.CloudError as exc:
                    join_err = exc
        except lop_hoc.CloudError as exc:
            QApplication.restoreOverrideCursor()
            self.err.setText(err_text(exc))
            return
        QApplication.restoreOverrideCursor()
        if join_err is not None:
            T.warn(self, tr("Chưa vào được lớp"),
                   tr("Đã tạo tài khoản nhưng chưa vào được lớp: {err}\nVào mục “Lớp học” để nhập lại mã lớp.")
                   .format(err=err_text(join_err)))
        elif joined:
            T.info(self, tr("Chào mừng!"), tr("Đã tạo tài khoản và vào lớp “{name}”.").format(name=joined["ten"]))
        self.result_data = (cloud, profile, username, self.pw.text())
        self.accept()


# ====================================================================== trang Lớp học


class ClassPage(_Page):
    def __init__(self, shell, lop_id: str | None = None):
        super().__init__()
        self.shell, self.cloud = shell, shell.main.cloud
        self.user = shell.account.username
        if self.cloud is None:
            self._offline()
        elif shell.account.is_staff:
            self._teacher(lop_id)
        else:
            self._student()
        self.lay.addStretch()

    def reload(self, **kw):
        self.shell.go("lop", **kw)

    # ------------------------------------------------------------ ngoại tuyến
    def _offline(self):
        self.lay.addWidget(T.page_header(tr("Lớp học"), tr("Theo dõi lớp và gửi điểm cho giáo viên.")))
        n = lop_hoc.pending(self.user)
        text = tr("Bạn đang ngoại tuyến (không kết nối được máy chủ lớp học). Bài làm vẫn được lưu trên máy và tự gửi "
                  "khi bạn đăng nhập lại lúc có mạng.")
        if n:
            text += "\n\n" + tr("Đang chờ gửi: {n} mục.").format(n=n)
        self.lay.addWidget(T.empty_state(text, "📡"))

    # ------------------------------------------------------------ học sinh
    def _student(self):
        self.lay.addWidget(T.page_header(tr("Lớp học"), tr("Vào lớp bằng mã giáo viên gửi. Điểm bài làm của bạn tự "
                                                          "gửi cho giáo viên.")))
        join = Card(padding=20, spacing=10)
        join.lay.addWidget(label(tr("Tham gia lớp"), "h3"))
        row = QHBoxLayout()
        self.code = QLineEdit()
        self.code.setProperty("size", "lg")
        self.code.setMaxLength(6)
        self.code.setPlaceholderText(tr("Mã lớp, vd K7M2QX"))
        self.code.setMaximumWidth(260)
        self.code.textEdited.connect(lambda t: self.code.setText(t.upper()))
        self.code.returnPressed.connect(self.join)
        row.addWidget(self.code)
        row.addWidget(button(tr("Vào lớp"), self.join, "primary"))
        row.addStretch()
        join.lay.addLayout(row)
        self.lay.addWidget(join)

        classes = busy(self, self.cloud.my_classes) or []
        self.lay.addWidget(label(tr("Lớp của tôi"), "h2"))
        if not classes:
            self.lay.addWidget(T.empty_state(tr("Bạn chưa vào lớp nào."), "🏫"))
        grid = QGridLayout()
        grid.setSpacing(14)
        for i, c in enumerate(classes):
            card = Card(padding=18, spacing=6)
            card.lay.addWidget(label("🏫  " + c["ten"], "h3"))
            card.lay.addWidget(label(tr("Giáo viên: {name}").format(name=c.get("ten_giao_vien", "")), "muted"))
            foot = QHBoxLayout()
            foot.addWidget(label(tr("{n} học sinh").format(n=c.get("si_so", 0)), "caption"))
            foot.addStretch()
            foot.addWidget(button(tr("Rời lớp"), lambda c=c: self.leave(c), "ghost", "sm"))
            card.lay.addLayout(foot)
            grid.addWidget(card, i // 3, i % 3)
        for col in range(3):
            grid.setColumnStretch(col, 1)
        self.lay.addLayout(grid)

        sync = Card(padding=18, spacing=8)
        sync.lay.addWidget(label(tr("Gửi điểm lên máy chủ"), "h3"))
        sent = busy(self, self.cloud.results) or []
        waiting = lop_hoc.pending(self.user)
        sync.lay.addWidget(label(tr("Đã gửi {n} bài làm. Đang chờ gửi: {m} mục.").format(n=len(sent), m=waiting),
                                 "muted"))
        if waiting:
            sync.lay.addLayout(_toolbar(button(tr("Gửi ngay"), self.flush, "primary", "sm")))
        self.lay.addWidget(sync)

    def join(self):
        code = self.code.text().strip().upper()
        if len(code) != 6:
            T.warn(self, tr("Mã lớp"), tr("Mã lớp gồm 6 ký tự."))
            return
        joined = busy(self, self.cloud.join_class, code)
        if joined:
            T.info(self, tr("Chào mừng!"), tr("Bạn đã vào lớp “{name}”.").format(name=joined["ten"]))
            self.reload()

    def leave(self, c):
        if T.confirm(self, tr("Rời lớp"), tr("Rời lớp “{name}”? Giáo viên sẽ không xem được điểm của bạn nữa.")
                     .format(name=c["ten"])):
            act(self, self.cloud.leave_class, c["id"])
            self.reload()

    def flush(self):
        n = busy(self, lop_hoc.flush, self.cloud, self.user)
        if n is not None:
            T.info(self, tr("Đã gửi"), tr("Đã gửi {n} mục lên máy chủ.").format(n=n))
            self.reload()

    # ------------------------------------------------------------ giáo viên
    def _teacher(self, lop_id):
        self.lay.addWidget(T.page_header(
            tr("Lớp học"), tr("Tạo lớp, gửi mã lớp cho học sinh, theo dõi điểm và tiến độ học của cả lớp."),
            [button(tr("Làm mới"), lambda: self.reload(lop_id=lop_id)),
             button(tr("+ Tạo lớp"), self.create, "primary")]))
        self.classes = busy(self, self.cloud.my_classes)
        if self.classes is None:
            return
        if not self.classes:
            self.lay.addWidget(T.empty_state(tr("Chưa có lớp nào. Bấm “+ Tạo lớp” để bắt đầu."), "🏫"))
            return
        cur = next((c for c in self.classes if c["id"] == lop_id), self.classes[0])
        tabs = QGridLayout()
        tabs.setSpacing(8)
        for i, c in enumerate(self.classes):
            text = f"{c['ten']}  ·  {c.get('si_so', 0)}"
            if c.get("giao_vien") != self.cloud.user_id:
                text += f"  ({c.get('ten_giao_vien', '')})"
            tabs.addWidget(button(text, lambda c=c: self.reload(lop_id=c["id"]),
                                  "primary" if c is cur else None, "sm"), i // 5, i % 5)
        tabs.setColumnStretch(5, 1)
        self.lay.addLayout(tabs)
        self._class_detail(cur)

    def _class_detail(self, c):
        self.cur = c
        head = Card(padding=20, spacing=10)
        top = QHBoxLayout()
        names = QVBoxLayout()
        names.setSpacing(2)
        names.addWidget(label(c["ten"], "h2"))
        names.addWidget(label(tr("Giáo viên: {name} · {n} học sinh").format(name=c.get("ten_giao_vien", ""),
                                                                             n=c.get("si_so", 0)), "muted"))
        top.addLayout(names, 1)
        code = label(c.get("ma", ""))
        code.setStyleSheet(f"font-size: 22pt; font-weight: 800; letter-spacing: 4px; color: {T.FOREST};"
                           f"background: {PRIMARY_SOFT}; border-radius: 12px; padding: 6px 16px;")
        code.setTextInteractionFlags(Qt.TextSelectableByMouse)
        top.addWidget(code)
        top.addWidget(button(tr("Sao chép mã"), self.copy_code, size="sm"))
        head.lay.addLayout(top)
        head.lay.addWidget(_banner(tr("Học sinh mở app → Đăng ký → nhập mã lớp {code}. Học sinh đã có tài khoản: vào "
                                      "mục Lớp học → Tham gia lớp.").format(code=c.get("ma", "")), "info"))
        head.lay.addLayout(_toolbar(right=(button(tr("Đổi tên"), self.rename, size="sm"),
                                           button(tr("Xuất Excel"), self.export, size="sm"),
                                           button(tr("Xóa lớp"), self.delete, "danger", "sm"))))
        self.lay.addWidget(head)

        self.members = busy(self, self.cloud.members, c["id"]) or []
        ids = [m["id"] for m in self.members]
        self.results = (busy(self, self.cloud.results, ids) or []) if ids else []
        progress = (busy(self, self.cloud.progress, ids) or []) if ids else []
        self.progress_rows = progress
        self.summary = lop_hoc.summarize(self.members, self.results, progress)

        stats = QHBoxLayout()
        stats.setSpacing(16)
        passed = sum(bool(r.get("passed")) for r in self.results)
        avg = round(sum(r.get("score", 0) for r in self.results) / len(self.results)) if self.results else 0
        stats.addWidget(T.stat_card(tr("Sĩ số"), str(len(self.members)), icon="👥"))
        stats.addWidget(T.stat_card(tr("Lượt làm bài"), str(len(self.results)), icon="📝"))
        stats.addWidget(T.stat_card(tr("Tỉ lệ đạt"), f"{round(100 * passed / len(self.results)) if self.results else 0}%",
                                    accent=SUCCESS, icon="✓"))
        stats.addWidget(T.stat_card(tr("Điểm trung bình"), str(avg), accent=WARN, icon="★"))
        self.lay.addLayout(stats)

        self.lay.addWidget(label(tr("Học sinh"), "h2"))
        if not self.members:
            self.lay.addWidget(T.empty_state(tr("Chưa có học sinh nào vào lớp. Gửi mã lớp {code} cho học sinh.")
                                             .format(code=c.get("ma", "")), "👋"))
            return
        self.table = T.table([(tr("Họ tên"), None), (tr("Tên đăng nhập"), 120), (tr("Lần làm"), 72),
                              (tr("Lần đạt"), 72), ("Word", 64), ("Excel", 64), ("PowerPoint", 96),
                              (tr("Gần nhất"), 150), (tr("Bài giảng xong"), 118)])
        rows, tones = [], []
        for s in self.summary:
            b = s["best"]
            rows.append([s["ho_ten"], s["username"], s["lan_lam"], s["lan_dat"],
                         b["WORD"] if b["WORD"] is not None else "—", b["EXCEL"] if b["EXCEL"] is not None else "—",
                         b["POWERPOINT"] if b["POWERPOINT"] is not None else "—", local_time(s["gan_nhat"]),
                         s["bai_xong"]])
            tones.append("bad" if s["khoa"] else ("ok" if s["lan_dat"] else None))
        T.set_rows(self.table, rows, tones, tone_cols={0}, keys=[s["id"] for s in self.summary])
        self.table.setMinimumHeight(min(80 + 42 * len(rows), 460))
        self.table.doubleClicked.connect(lambda *_: self.detail())
        self.lay.addWidget(self.table)
        self.lay.addLayout(_toolbar(button(tr("Xem chi tiết"), self.detail, size="sm"),
                                    button(tr("Đặt lại mật khẩu"), self.reset_password, size="sm"),
                                    button(tr("Xóa khỏi lớp"), self.remove, "danger", "sm")))

        self.lay.addWidget(label(tr("Bài làm gần đây"), "h2"))
        names = {m["id"]: m.get("ho_ten", "") for m in self.members}
        recent = T.table([(tr("Thời gian"), 150), (tr("Học sinh"), 180), (tr("Bài thi"), None), (tr("Chế độ"), 150),
                          (tr("Điểm"), 70), (tr("Kết quả"), 100)])
        rows, tones = [], []
        for r in self.results[:30]:
            rows.append([local_time(r.get("lam_luc", "")), names.get(r.get("hoc_vien"), ""), exam_title(r),
                         mode_name(r.get("mode", "")), r.get("score"), tr("Đạt") if r.get("passed") else tr("Chưa đạt")])
            tones.append("ok" if r.get("passed") else "bad")
        T.set_rows(recent, rows, tones)
        recent.setMinimumHeight(min(80 + 42 * max(len(rows), 1), 520))
        self.lay.addWidget(recent)

    def _selected_student(self):
        key = T.selected_key(self.table) if hasattr(self, "table") else None
        if key is None:
            T.info(self, tr("Chọn học sinh"), tr("Hãy chọn một học sinh trong bảng."))
            return None
        return next(m for m in self.members if m["id"] == key)

    def copy_code(self):
        QApplication.clipboard().setText(self.cur.get("ma", ""))
        T.info(self, tr("Đã sao chép"), tr("Đã chép mã lớp {code} vào clipboard.").format(code=self.cur.get("ma", "")))

    def create(self):
        dlg = _NameDialog(self.window(), tr("Tạo lớp"), "", tr("vd: 10A1 – Tin học"))
        if dlg.exec():
            c = busy(self, self.cloud.create_class, dlg.value)
            if c:
                self.reload(lop_id=c["id"])

    def rename(self):
        dlg = _NameDialog(self.window(), tr("Đổi tên lớp"), self.cur["ten"])
        if dlg.exec():
            act(self, self.cloud.rename_class, self.cur["id"], dlg.value)
            self.reload(lop_id=self.cur["id"])

    def delete(self):
        if T.confirm(self, tr("Xóa lớp"), tr("Xóa lớp “{name}”? Học sinh vẫn giữ tài khoản và điểm, chỉ không còn trong "
                                            "lớp này.").format(name=self.cur["ten"])):
            act(self, self.cloud.delete_class, self.cur["id"])
            self.reload()

    def export(self):
        safe = re.sub(r"[^\w\- ]", "_", self.cur["ten"]).strip() or "lop"
        path, _ = QFileDialog.getSaveFileName(self, tr("Lưu file"), f"Diem_{safe}.xlsx", "Excel (*.xlsx)")
        if not path:
            return
        try:
            lop_hoc.export_excel(path, self.cur["ten"], getattr(self, "summary", []), getattr(self, "results", []),
                                 getattr(self, "members", []))
        except OSError as exc:
            T.error(self, tr("Lỗi"), str(exc))
            return
        T.info(self, tr("Đã lưu"), tr("Đã lưu file:\n{path}").format(path=path))

    def detail(self):
        m = self._selected_student()
        if m:
            StudentDialog(self.window(), m, [r for r in self.results if r.get("hoc_vien") == m["id"]],
                          [p for p in self.progress_rows if p.get("hoc_vien") == m["id"]]).exec()

    def reset_password(self):
        m = self._selected_student()
        if not m:
            return
        pw = random_password()
        if T.confirm(self, tr("Đặt lại mật khẩu"), tr("Đặt mật khẩu mới cho “{user}” là:\n\n      {pw}")
                     .format(user=m["username"], pw=pw)):
            if act(self, self.cloud.reset_password, m["id"], pw):
                QApplication.clipboard().setText(pw)
                T.info(self, tr("Đã đặt lại"), tr("Mật khẩu mới: {pw}\n(đã chép vào clipboard)").format(pw=pw))

    def remove(self):
        m = self._selected_student()
        if m and T.confirm(self, tr("Xóa khỏi lớp"), tr("Xóa “{name}” khỏi lớp? Tài khoản và điểm vẫn giữ.")
                           .format(name=m.get("ho_ten") or m["username"])):
            act(self, self.cloud.leave_class, self.cur["id"], m["id"])
            self.reload(lop_id=self.cur["id"])


class _NameDialog(_Dialog):
    def __init__(self, parent, title, value="", placeholder=""):
        super().__init__(parent, title, 420)
        self.edit = self.field(tr("Tên lớp"), QLineEdit(value))
        self.edit.setPlaceholderText(placeholder)
        self.edit.returnPressed.connect(self.ok)
        self.value = ""
        self.buttons(tr("Lưu"), self.ok)

    def ok(self):
        self.value = self.edit.text().strip()
        if self.value:
            self.accept()


class StudentDialog(_Dialog):
    """Chi tiết một học sinh: mọi lần làm bài + tiến độ bài giảng."""

    def __init__(self, parent, member, results, progress):
        super().__init__(parent, member.get("ho_ten") or member["username"], 820)
        self.lay.addWidget(label(tr("Tên đăng nhập: {user} · {n} lần làm bài")
                                 .format(user=member["username"], n=len(results)), "muted"))
        t = T.table([(tr("Thời gian"), 150), (tr("Bài thi"), None), (tr("Chế độ"), 150), (tr("Điểm"), 70),
                     (tr("Đúng"), 70), (tr("Kết quả"), 100)])
        T.set_rows(t, [[local_time(r.get("lam_luc", "")), exam_title(r), mode_name(r.get("mode", "")), r.get("score"),
                        f"{r.get('correct')}/{r.get('total')}", tr("Đạt") if r.get("passed") else tr("Chưa đạt")]
                       for r in results], ["ok" if r.get("passed") else "bad" for r in results])
        t.setMinimumHeight(260)
        self.lay.addWidget(t)
        wrong = [w for r in results[:1] for w in (r.get("wrong") or [])]
        if wrong:
            self.lay.addWidget(label(tr("Câu sai ở lần gần nhất:"), "h3"))
            self.lay.addWidget(label("• " + "<br>• ".join(esc(w) for w in wrong[:12]), "muted", wrap=True))
        self.lay.addWidget(label(tr("Tài liệu học"), "h3"))
        p = T.table([(tr("Bài giảng"), None), (tr("Đã xem"), 90), (tr("Xong"), 70), (tr("Điểm"), 90)])
        T.set_rows(p, [[x.get("ten_bai") or x.get("bai"), f"{x.get('xem')}/{x.get('tong')}",
                        "✓" if x.get("xong") else "", x.get("diem") or ""] for x in progress])
        p.setMinimumHeight(160)
        self.lay.addWidget(p)
        self.buttons(tr("Đóng"), cancel=False)


# ====================================================================== tài khoản máy chủ (quản trị)


class OnlineAccountsPage(_Page):
    def __init__(self, shell):
        super().__init__()
        self.shell, self.cloud = shell, shell.main.cloud
        self.lay.addWidget(T.page_header(
            tr("Tài khoản"), tr("Tài khoản trên máy chủ lớp học. Học sinh tự đăng ký trong app; cấp quyền giáo viên "
                                "tại đây."), [button(tr("Máy chủ lớp học…"), self.server)]))
        if self.cloud is None:
            self.lay.addWidget(T.empty_state(tr("Đang ngoại tuyến – không tải được danh sách tài khoản."), "📡"))
            self.lay.addStretch()
            return
        self.stats = QHBoxLayout()
        self.stats.setSpacing(16)
        self.lay.addLayout(self.stats)
        self.search = QLineEdit()
        self.search.setPlaceholderText(tr("Tìm theo tên đăng nhập hoặc họ tên…"))
        self.search.setMaximumWidth(340)
        self.search.textChanged.connect(self.fill)
        self.lay.addLayout(_toolbar(self.search, right=(
            button(tr("Đổi vai trò…"), self.change_role, size="sm"),
            button(tr("Khóa / Mở khóa"), self.toggle_lock, size="sm"),
            button(tr("Đặt lại mật khẩu"), self.reset_password, size="sm"),
            button(tr("Xóa"), self.delete, "danger", "sm"))))
        self.table = T.table([(tr("Tên đăng nhập"), 170), (tr("Họ tên"), None), (tr("Vai trò"), 130),
                              (tr("Trạng thái"), 120), (tr("Tạo lúc"), 150)])
        self.table.setMinimumHeight(420)
        self.lay.addWidget(self.table, 1)
        self.load()

    def load(self):
        self.profiles = busy(self, self.cloud.profiles) or []
        while self.stats.count():
            item = self.stats.takeAt(0)
            if item.widget():
                item.widget().deleteLater()
        self.stats.addWidget(T.stat_card(tr("Tổng số tài khoản"), str(len(self.profiles))))
        self.stats.addWidget(T.stat_card(tr("Giáo viên"), str(sum(p["vai_tro"] == "giao_vien" for p in self.profiles)),
                                         accent=WARN))
        self.stats.addWidget(T.stat_card(tr("Học viên"), str(sum(p["vai_tro"] == "hoc_vien" for p in self.profiles)),
                                         accent=SUCCESS))
        self.fill()

    def fill(self):
        q = self.search.text().strip().casefold()
        order = {"quan_tri": 0, "giao_vien": 1, "hoc_vien": 2}
        shown = sorted((p for p in self.profiles if not q or q in p["username"] or q in p.get("ho_ten", "").casefold()),
                       key=lambda p: (order.get(p["vai_tro"], 3), p["username"]))
        T.set_rows(self.table, [[p["username"], p.get("ho_ten", ""), role_name(ROLE_NAMES[p["vai_tro"]]),
                                 tr("Bị khóa") if p.get("khoa") else tr("Hoạt động"), local_time(p.get("tao_luc", ""))]
                                for p in shown], ["bad" if p.get("khoa") else "ok" for p in shown], tone_cols={3},
                   keys=[p["id"] for p in shown])

    def _selected(self):
        key = T.selected_key(self.table)
        if key is None:
            T.info(self, tr("Chọn tài khoản"), tr("Hãy chọn một tài khoản trong bảng."))
            return None
        return next(p for p in self.profiles if p["id"] == key)

    def change_role(self):
        p = self._selected()
        if not p:
            return
        dlg = _Dialog(self.window(), tr("Đổi vai trò"), 400)
        dlg.lay.addWidget(label(f"{p['username']} – {esc(p.get('ho_ten', ''))}", "muted"))
        box = QComboBox()
        for key, name in ROLE_NAMES.items():
            box.addItem(role_name(name), key)
        box.setCurrentIndex(box.findData(p["vai_tro"]))
        dlg.field(tr("Vai trò"), box)
        dlg.buttons()
        if dlg.exec() and box.currentData() != p["vai_tro"]:
            act(self, self.cloud.set_role, p["id"], box.currentData())
            self.load()

    def toggle_lock(self):
        p = self._selected()
        if p:
            act(self, self.cloud.set_locked, p["id"], not p.get("khoa"))
            self.load()

    def reset_password(self):
        p = self._selected()
        if not p:
            return
        pw = random_password()
        if T.confirm(self, tr("Đặt lại mật khẩu"), tr("Đặt mật khẩu mới cho “{user}” là:\n\n      {pw}")
                     .format(user=p["username"], pw=pw)):
            if not act(self, self.cloud.reset_password, p["id"], pw):
                return
            QApplication.clipboard().setText(pw)
            T.info(self, tr("Đã đặt lại"), tr("Mật khẩu mới: {pw}\n(đã chép vào clipboard)").format(pw=pw))

    def delete(self):
        p = self._selected()
        if p and T.confirm(self, tr("Xóa tài khoản"), tr("Xóa hẳn tài khoản “{user}” cùng toàn bộ điểm trên máy chủ?")
                           .format(user=p["username"])):
            act(self, self.cloud.delete_account, p["id"])
            self.load()

    def server(self):
        ServerDialog(self.window()).exec()

