"""Các trang Quản trị: Tài khoản, Đề thi (kèm Soạn đề), Kết quả học viên."""
from __future__ import annotations

import csv
import shutil
from pathlib import Path

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (QApplication, QCheckBox, QComboBox, QDialog, QFileDialog, QFormLayout, QFrame,
                               QHBoxLayout, QLabel, QLineEdit, QListWidget, QPlainTextEdit, QScrollArea,
                               QSplitter, QVBoxLayout, QWidget)

from .. import core, custom, rules
from .. import importer
from ..accounts import ROLE_NAMES, STUDENT, AccountError, random_password
from . import theme as T
from .app import _banner, page_body
from .theme import (DANGER, DANGER_SOFT, MUTED, SUCCESS, SUCCESS_SOFT, WARN, WARN_SOFT, Card, button, chip,
                    label)

MON_NAMES = {"WORD": "Word", "EXCEL": "Excel", "POWERPOINT": "PowerPoint"}
FILE_FILTERS = {"WORD": "Word (*.docx)", "EXCEL": "Excel (*.xlsx)", "POWERPOINT": "PowerPoint (*.pptx)"}


class _Page(QScrollArea):
    def __init__(self):
        super().__init__()
        self.body, self.lay = page_body()
        self.setObjectName("Page")
        self.setWidgetResizable(True)
        self.setFrameShape(QFrame.NoFrame)
        self.setWidget(self.body)


def _toolbar(*widgets, right=()) -> QHBoxLayout:
    row = QHBoxLayout()
    row.setSpacing(8)
    for w in widgets:
        row.addWidget(w)
    row.addStretch()
    for w in right:
        row.addWidget(w)
    return row


def _save_csv(parent, default_name, header, rows):
    path, _ = QFileDialog.getSaveFileName(parent, "Lưu file", default_name, "CSV (*.csv)")
    if not path:
        return
    with open(path, "w", newline="", encoding="utf-8-sig") as f:   # utf-8-sig để Excel đọc đúng tiếng Việt
        w = csv.writer(f)
        w.writerow(header)
        w.writerows(rows)
    T.info(parent, "Đã lưu", f"Đã lưu file:\n{path}")


# ====================================================================== Tài khoản


class AccountsPage(_Page):
    """Quản trị: mọi tài khoản. Giáo viên: chỉ học viên của mình."""

    def __init__(self, shell):
        super().__init__()
        self.shell, self.store, self.me = shell, shell.main.store, shell.account
        title = "Tài khoản" if self.me.is_admin else "Học viên của tôi"
        sub = ("Tạo / nhập tài khoản, phân quyền giáo viên, đặt lại mật khẩu, khóa hoặc đặt hạn dùng."
               if self.me.is_admin else "Tạo hoặc nhập danh sách học viên, đặt lại mật khẩu, khóa hoặc đặt hạn dùng.")
        self.lay.addWidget(T.page_header(title, sub, [
            button("Nhập từ file (CSV / Excel)", self.import_file),
            button("Tạo cho cả lớp", self.bulk), button("+ Thêm tài khoản", self.add, "primary")]))
        self.stats = QHBoxLayout()
        self.stats.setSpacing(16)
        self.lay.addLayout(self.stats)

        self.search = QLineEdit()
        self.search.setPlaceholderText("Tìm theo tên đăng nhập, họ tên hoặc lớp…")
        self.search.setMinimumWidth(300)
        self.search.textChanged.connect(self.refresh)
        self.lop = QComboBox()
        self.lop.setMinimumWidth(150)
        self.lop.currentIndexChanged.connect(self.refresh)
        self.lay.addLayout(_toolbar(self.search, self.lop, right=(
            button("Sửa", self.edit), button("Đặt lại mật khẩu", self.reset_password),
            button("Khóa / Mở khóa", self.toggle_lock), button("Xóa", self.delete, "danger"))))
        cols = [("Tên đăng nhập", 150), ("Họ tên", None), ("Lớp", 90), ("Vai trò", 100)]
        if self.me.is_admin:
            cols.append(("Giáo viên phụ trách", 150))
        cols += [("Trạng thái", 110), ("Hạn dùng", 110), ("Số lần thi", 90), ("Điểm cao nhất", 110)]
        self.table = T.table(cols)
        self.status_col = [c[0] for c in cols].index("Trạng thái")
        self.table.setMinimumHeight(380)
        self.table.doubleClicked.connect(lambda *_: self.edit())
        self.lay.addWidget(self.table, 1)
        self._fill_classes()
        self.refresh()

    def _fill_classes(self):
        current = self.lop.currentData()
        self.lop.blockSignals(True)
        self.lop.clear()
        self.lop.addItem("Tất cả lớp", None)
        for c in sorted({a.lop for a in self.store.visible_to(self.me) if a.lop}):
            self.lop.addItem(f"Lớp {c}", c)
        self.lop.setCurrentIndex(max(0, self.lop.findData(current)))
        self.lop.blockSignals(False)

    def refresh(self):
        while self.stats.count():
            item = self.stats.takeAt(0)
            if item.widget():
                item.widget().deleteLater()
        accounts = self.store.visible_to(self.me)
        students = [a for a in accounts if a.vai_tro == STUDENT]
        active = [a for a in students if not a.khoa and not a.expired()]
        if self.me.is_admin:
            self.stats.addWidget(T.stat_card("Tổng số tài khoản", str(len(accounts))))
            self.stats.addWidget(T.stat_card("Giáo viên", str(sum(a.is_teacher for a in accounts)), accent="#7A5AF8"))
        self.stats.addWidget(T.stat_card("Học viên", str(len(students)), accent=SUCCESS))
        self.stats.addWidget(T.stat_card("Học viên đang hoạt động", str(len(active)), accent=WARN))
        self.stats.addWidget(T.stat_card("Số lớp", str(len({a.lop for a in students if a.lop}))))

        q, lop = self.search.text().strip().casefold(), self.lop.currentData()
        history = core.load_history()
        rows, tones, keys = [], [], []
        for a in accounts:
            if q and q not in a.username and q not in a.ho_ten.casefold() and q not in a.lop.casefold():
                continue
            if lop and a.lop != lop:
                continue
            mine = [h for h in history if h.get("user") == a.username]
            status = "Bị khóa" if a.khoa else "Hết hạn" if a.expired() else "Hoạt động"
            row = [a.username, a.ho_ten, a.lop or "—", ROLE_NAMES[a.vai_tro]]
            if self.me.is_admin:
                gv = self.store.get(a.giao_vien) if a.giao_vien else None
                row.append(gv.ho_ten if gv else "—")
            row += [status, a.han_dung or "—", len(mine), max((h["score"] for h in mine), default="—")]
            rows.append(row)
            tones.append("ok" if status == "Hoạt động" else "bad")
            keys.append(a.username)
        T.set_rows(self.table, rows, tones, tone_cols={self.status_col}, keys=keys)

    def _reload(self):
        self._fill_classes()
        self.refresh()

    def _selected(self, manage=True):
        key = T.selected_key(self.table)
        if key is None:
            T.info(self, "Chọn tài khoản", "Hãy chọn một tài khoản trong bảng.")
            return None
        acc = self.store.get(key)
        if manage and not self.store.can_manage(self.me, acc):
            T.warn(self, "Không có quyền", "Giáo viên chỉ quản lý được tài khoản học viên của mình.")
            return None
        return acc

    def add(self):
        if UserDialog(self, self.store, self.me, None).exec():
            self._reload()

    def edit(self):
        acc = self._selected()
        if acc and UserDialog(self, self.store, self.me, acc).exec():
            self._reload()

    def bulk(self):
        dlg = BulkUsersDialog(self, self.store, self.me)
        if dlg.exec():
            self._reload()
            CredentialsDialog(self, dlg.created, dlg.errors).exec()

    def import_file(self):
        dlg = ImportDialog(self, self.store, self.me)
        if dlg.exec():
            self._reload()
            CredentialsDialog(self, dlg.created, dlg.errors).exec()

    def reset_password(self):
        acc = self._selected()
        if not acc:
            return
        pw = random_password()
        if T.confirm(self, "Đặt lại mật khẩu", f"Đặt mật khẩu mới cho “{acc.username}” là:\n\n      {pw}\n\n"
                                               "Người dùng sẽ phải đổi mật khẩu khi đăng nhập."):
            self.store.set_password(acc.username, pw, must_change=True)
            QApplication.clipboard().setText(pw)
            T.info(self, "Đã đặt lại", f"Mật khẩu mới: {pw}\n(đã chép vào clipboard)")

    def toggle_lock(self):
        acc = self._selected()
        if not acc:
            return
        if acc.username == self.me.username:
            T.warn(self, "Không thể", "Không thể khóa tài khoản đang đăng nhập.")
            return
        try:
            self.store.update(acc.username, khoa=not acc.khoa)
        except AccountError as exc:
            T.error(self, "Lỗi", str(exc))
        self.refresh()

    def delete(self):
        acc = self._selected()
        if not acc:
            return
        if acc.username == self.me.username:
            T.warn(self, "Không thể", "Không thể xóa tài khoản đang đăng nhập.")
            return
        if T.confirm(self, "Xóa tài khoản", f"Xóa tài khoản “{acc.username}”?\n"
                                            "(Kết quả làm bài cũ vẫn được giữ lại.)"):
            try:
                self.store.delete(acc.username)
            except AccountError as exc:
                T.error(self, "Lỗi", str(exc))
            self._reload()


class _Dialog(QDialog):
    def __init__(self, parent, title, width=460):
        super().__init__(parent)
        self.setWindowTitle(title)
        self.setMinimumWidth(width)
        self.lay = QVBoxLayout(self)
        self.lay.setContentsMargins(28, 24, 28, 22)
        self.lay.setSpacing(10)
        self.lay.addWidget(label(title, "h2"))

    def field(self, text, widget, hint=None):
        self.lay.addSpacing(4)
        self.lay.addWidget(label(text, "h3"))
        self.lay.addWidget(widget)
        if hint:
            self.lay.addWidget(label(hint, "caption", wrap=True))
        return widget

    def buttons(self, ok_text="Lưu", ok=None, cancel=True):
        row = QHBoxLayout()
        row.addStretch()
        if cancel:
            row.addWidget(button("Hủy", self.reject))
        row.addWidget(button(ok_text, ok or self.accept, "primary"))
        self.lay.addSpacing(6)
        self.lay.addLayout(row)


def _teacher_combo(store, me, current=None) -> QComboBox | None:
    """Ô chọn giáo viên phụ trách (chỉ quản trị mới chọn; giáo viên luôn là chính mình)."""
    if not me.is_admin:
        return None
    cb = QComboBox()
    cb.addItem("(Không có – quản trị quản lý)", None)
    for t in store.teachers():
        cb.addItem(f"{t.ho_ten} ({t.username})", t.username)
    cb.setCurrentIndex(max(0, cb.findData(current)))
    return cb


def _owner(me, combo) -> str | None:
    return me.username if me.is_teacher else (combo.currentData() if combo else None)


class UserDialog(_Dialog):
    def __init__(self, parent, store, me, acc):
        super().__init__(parent, "Sửa tài khoản" if acc else "Thêm tài khoản")
        self.store, self.me, self.acc = store, me, acc
        self.username = self.field("Tên đăng nhập", QLineEdit(acc.username if acc else ""),
                                   "Chữ thường không dấu, số, dấu _ hoặc . (để trống = tạo từ họ tên)")
        self.username.setEnabled(acc is None)
        self.ho_ten = self.field("Họ tên", QLineEdit(acc.ho_ten if acc else ""))
        if not acc:
            self.password = self.field("Mật khẩu", QLineEdit(random_password()), "Đã tạo ngẫu nhiên, có thể sửa.")
        self.role = QComboBox()
        roles = store.assignable_roles(me)
        if acc and acc.vai_tro not in roles:
            roles = [acc.vai_tro] + roles
        for r in roles:
            self.role.addItem(ROLE_NAMES[r], r)
        self.role.setCurrentIndex(max(0, self.role.findData(acc.vai_tro if acc else STUDENT)))
        self.role.setEnabled(len(roles) > 1)
        self.field("Vai trò", self.role, "Giáo viên: quản lý học viên của mình, soạn đề, xem kết quả."
                   if me.is_admin else None)
        self.lop = self.field("Lớp (tuỳ chọn)", QLineEdit(acc.lop if acc else ""))
        self.lop.setPlaceholderText("vd: 10A1")
        self.teacher = _teacher_combo(store, me, acc.giao_vien if acc else None)
        if self.teacher:
            self.field("Giáo viên phụ trách (với học viên)", self.teacher)
        self.han = self.field("Hạn dùng (tuỳ chọn)", QLineEdit((acc.han_dung or "") if acc else ""),
                              "Dạng YYYY-MM-DD, vd 2026-12-31. Để trống = không giới hạn.")
        self.han.setPlaceholderText("YYYY-MM-DD")
        if not acc:
            self.must = QCheckBox("Bắt đổi mật khẩu ở lần đăng nhập đầu")
            self.must.setChecked(True)
            self.lay.addWidget(self.must)
        self.buttons(ok=self.save)

    def save(self):
        role = self.role.currentData()
        owner = _owner(self.me, self.teacher) if role == STUDENT else None
        try:
            if self.acc:
                kw = {"giao_vien": owner} if self.me.is_admin else {}
                self.store.update(self.acc.username, ho_ten=self.ho_ten.text(), vai_tro=role,
                                  han_dung=self.han.text(), lop=self.lop.text(), **kw)
            else:
                username = self.username.text().strip() or self.store.unique_username(self.ho_ten.text())
                acc = self.store.create(username, self.ho_ten.text(), self.password.text(), role,
                                        han_dung=self.han.text(), must_change=self.must.isChecked(),
                                        lop=self.lop.text(), giao_vien=owner)
                T.info(self, "Đã tạo tài khoản", f"Tên đăng nhập:  {acc.username}\nMật khẩu:  {self.password.text()}"
                                                 "\n\nHãy gửi thông tin này cho người dùng.")
        except AccountError as exc:
            T.error(self, "Lỗi", str(exc))
            return
        self.accept()


class BulkUsersDialog(_Dialog):
    def __init__(self, parent, store, me):
        super().__init__(parent, "Tạo tài khoản cho cả lớp", 560)
        self.store, self.me = store, me
        self.created, self.errors = [], []
        self.lay.addWidget(label("Dán danh sách, mỗi dòng một học viên: chỉ <b>Họ tên</b>, hoặc "
                                 "<b>tên_đăng_nhập, Họ tên</b>. Tên đăng nhập và mật khẩu được tạo tự động nếu thiếu.",
                                 "muted", wrap=True))
        self.text = QPlainTextEdit()
        self.text.setPlaceholderText("Nguyễn Văn An\nTrần Thị Bình\nlevanc, Lê Văn Cường")
        self.text.setMinimumHeight(200)
        self.lay.addWidget(self.text)
        self.lop = self.field("Lớp", QLineEdit())
        self.lop.setPlaceholderText("vd: 10A1")
        self.teacher = _teacher_combo(store, me)
        if self.teacher:
            self.field("Giáo viên phụ trách", self.teacher)
        self.han = self.field("Hạn dùng chung (tuỳ chọn)", QLineEdit(), "Dạng YYYY-MM-DD. Để trống = không giới hạn.")
        self.buttons("Tạo tài khoản", self.create)

    def create(self):
        owner = _owner(self.me, self.teacher)
        for n, line in enumerate(self.text.toPlainText().splitlines(), start=1):
            if not line.strip():
                continue
            first, sep, rest = line.partition(",")
            username, name = (first.strip(), rest.strip()) if sep else ("", first.strip())
            username = username or self.store.unique_username(name)
            pw = random_password()
            try:
                acc = self.store.create(username, name, pw, STUDENT, han_dung=self.han.text(),
                                        lop=self.lop.text(), giao_vien=owner)
                self.created.append((acc.username, acc.ho_ten, acc.lop, pw))
            except AccountError as exc:
                self.errors.append(f"Dòng {n}: {exc}")
        if not self.created:
            T.error(self, "Chưa tạo được", "\n".join(self.errors[:12]) or "Danh sách đang trống.")
            self.errors.clear()
            return
        self.accept()


class ImportDialog(_Dialog):
    """Nhập danh sách học viên từ file CSV / Excel: chọn file → xem trước → tạo."""

    def __init__(self, parent, store, me):
        super().__init__(parent, "Nhập học viên từ file", 860)
        self.store, self.me = store, me
        self.entries, self.created, self.errors = [], [], []
        self.lay.addWidget(label(
            "File <b>Excel (.xlsx)</b> hoặc <b>CSV</b> có dòng tiêu đề. Chỉ cột <b>Họ tên</b> là bắt buộc; "
            "các cột tuỳ chọn: <b>Tên đăng nhập</b>, <b>Mật khẩu</b>, <b>Lớp</b>, <b>Hạn dùng</b>. "
            "Thiếu tên đăng nhập / mật khẩu thì app tự tạo.", "muted", wrap=True))
        row = QHBoxLayout()
        row.addWidget(button("Chọn file…", self.pick, "primary"))
        row.addWidget(button("Tải file mẫu (Excel)", lambda: self.template(".xlsx")))
        row.addWidget(button("Tải file mẫu (CSV)", lambda: self.template(".csv")))
        row.addStretch()
        self.file_lbl = label("Chưa chọn file", "muted")
        row.addWidget(self.file_lbl)
        self.lay.addLayout(row)
        self.teacher = _teacher_combo(store, me)
        if self.teacher:
            self.field("Giáo viên phụ trách các học viên này", self.teacher)
        self.summary = label("", wrap=True)
        self.lay.addWidget(self.summary)
        self.table = T.table([("Dòng", 60), ("Tên đăng nhập", 150), ("Họ tên", None), ("Lớp", 80),
                              ("Mật khẩu", 100), ("Trạng thái", 260)])
        self.table.setMinimumHeight(300)
        self.lay.addWidget(self.table, 1)
        row = QHBoxLayout()
        row.addStretch()
        row.addWidget(button("Hủy", self.reject))
        self.ok_btn = button("Tạo tài khoản", self.create, "primary")
        self.ok_btn.setEnabled(False)
        row.addWidget(self.ok_btn)
        self.lay.addLayout(row)

    def template(self, ext):
        path, _ = QFileDialog.getSaveFileName(self, "Lưu file mẫu", f"mau_danh_sach_hoc_vien{ext}",
                                              "Excel (*.xlsx)" if ext == ".xlsx" else "CSV (*.csv)")
        if path:
            importer.write_template(Path(path))
            T.info(self, "Đã lưu", f"Đã lưu file mẫu:\n{path}\n\nMở bằng Excel, điền danh sách rồi chọn lại ở đây.")

    def pick(self):
        path, _ = QFileDialog.getOpenFileName(self, "Chọn danh sách học viên", "",
                                              "Excel / CSV (*.xlsx *.csv);;Tất cả (*.*)")
        if path:
            self.load(Path(path))

    def load(self, path: Path):
        try:
            self.entries = importer.plan(path, self.store)
        except (importer.ImportFileError, OSError, ValueError) as exc:
            T.error(self, "Không đọc được file", str(exc))
            return
        self.file_lbl.setText(path.name)
        ok = [e for e in self.entries if not e.error]
        T.set_rows(self.table, [[e.row, e.username, e.ho_ten, e.lop or "—", e.password,
                                 e.error or "✓ Sẽ tạo"] for e in self.entries],
                   ["bad" if e.error else "ok" for e in self.entries])
        bad = len(self.entries) - len(ok)
        self.summary.setText(f"<b>{len(ok)}</b> học viên sẽ được tạo" +
                             (f" · <span style='color:{DANGER}'><b>{bad}</b> dòng bị bỏ qua (xem cột Trạng thái)"
                              "</span>" if bad else ""))
        self.ok_btn.setText(f"Tạo {len(ok)} tài khoản")
        self.ok_btn.setEnabled(bool(ok))

    def create(self):
        created, self.errors = importer.apply(self.entries, self.store, _owner(self.me, self.teacher))
        self.created = [(e.username, e.ho_ten, e.lop, e.password) for e in created]
        if not self.created:
            T.error(self, "Chưa tạo được", "\n".join(self.errors[:12]))
            return
        self.accept()


class CredentialsDialog(_Dialog):
    """Danh sách tài khoản vừa tạo – lưu ra CSV để phát cho học viên."""

    HEADER = ["Tên đăng nhập", "Họ tên", "Lớp", "Mật khẩu"]

    def __init__(self, parent, rows, errors):
        super().__init__(parent, f"Đã tạo {len(rows)} tài khoản", 680)
        self.rows = rows
        self.lay.addWidget(label("Hãy lưu danh sách này để phát cho học viên. Mật khẩu sẽ không hiển thị lại.",
                                 "muted", wrap=True))
        t = T.table([("Tên đăng nhập", 160), ("Họ tên", None), ("Lớp", 80), ("Mật khẩu", 110)])
        T.set_rows(t, [list(r) for r in rows])
        t.setMinimumHeight(min(80 + 42 * len(rows), 380))
        self.lay.addWidget(t)
        if errors:
            self.lay.addWidget(_banner("Bỏ qua: " + "; ".join(errors[:6]) +
                                       (f" … (+{len(errors) - 6})" if len(errors) > 6 else ""), "bad"))
        row = QHBoxLayout()
        row.addWidget(button("Chép vào clipboard", self.copy))
        row.addStretch()
        row.addWidget(button("Đóng", self.accept))
        row.addWidget(button("Lưu danh sách (CSV)", self.save, "primary"))
        self.lay.addLayout(row)

    def copy(self):
        QApplication.clipboard().setText("\n".join("\t".join(r) for r in [self.HEADER, *self.rows]))
        T.info(self, "Đã chép", "Đã chép danh sách vào clipboard – dán được vào Excel.")

    def save(self):
        _save_csv(self, "tai_khoan_hoc_vien.csv", self.HEADER, self.rows)


# ====================================================================== Đề thi


class ExamsPage(_Page):
    def __init__(self, shell):
        super().__init__()
        self.shell = shell
        self.lay.addWidget(T.page_header(
            "Đề thi", "Soạn đề của riêng bạn. Đề được gộp vào bài thi của môn tương ứng trên Trang chủ.",
            [button("+ Soạn đề mới", self.new, "primary")]))
        self.lay.addLayout(_toolbar(right=(button("Sửa", self.edit), button("Kiểm tra đề", self.check),
                                           button("Mở thư mục", self.open_folder),
                                           button("Xóa", self.delete, "danger"))))
        self.table = T.table([("Tên đề", None), ("Môn", 120), ("Dự án", 90), ("Nhiệm vụ", 100),
                              ("Trạng thái", 300)])
        self.table.setMinimumHeight(300)
        self.table.doubleClicked.connect(lambda *_: self.edit())
        self.lay.addWidget(self.table, 1)
        how = Card(padding=20, spacing=8)
        how.lay.addWidget(label("Cách soạn một đề", "h3"))
        for i, text in enumerate(("Soạn <b>file gốc</b> (file chưa làm) bằng Word / Excel / PowerPoint.",
                                  "Bấm <b>Soạn đề mới</b>, thêm dự án và chọn file gốc.",
                                  "Thêm <b>nhiệm vụ</b>: yêu cầu, gợi ý và <b>luật chấm</b> (chọn từ danh sách).",
                                  "Chọn <b>file đáp án</b> (đã làm đúng) rồi bấm <b>Lưu & kiểm tra</b>."), 1):
            how.lay.addWidget(label(f"<span style='color:{T.PRIMARY}; font-weight:800'>{i}.</span>  {text}",
                                    wrap=True))
        self.lay.addWidget(how)
        self.refresh()

    def refresh(self):
        rows, tones, keys = [], [], []
        for folder in custom.list_exam_folders():
            try:
                data = custom.read_exam_json(folder)
            except (OSError, ValueError):
                data = {}
            _, errors = custom.load_exam(folder)
            projects = data.get("du_an", [])
            rows.append([data.get("ten", folder.name), MON_NAMES.get(str(data.get("mon", "")).upper(), "?"),
                         len(projects), sum(len(p.get("nhiem_vu", [])) for p in projects),
                         f"Lỗi: {errors[0]}" if errors else "✓ Sẵn sàng"])
            tones.append("bad" if errors else "ok")
            keys.append(str(folder))
        T.set_rows(self.table, rows, tones, keys=keys)

    def _selected(self):
        key = T.selected_key(self.table)
        if key is None:
            T.info(self, "Chọn đề", "Hãy chọn một đề trong bảng.")
        return Path(key) if key else None

    def new(self):
        ExamEditor(self, None).exec()
        self.refresh()

    def edit(self):
        folder = self._selected()
        if folder:
            ExamEditor(self, folder).exec()
            self.refresh()

    def check(self):
        folder = self._selected()
        if folder:
            CheckReportDialog(self, folder).exec()

    def open_folder(self):
        folder = Path(T.selected_key(self.table) or custom.editable_dir())
        folder.mkdir(parents=True, exist_ok=True)
        core.open_in_office(folder)

    def delete(self):
        folder = self._selected()
        if folder and T.confirm(self, "Xóa đề", f"Xóa toàn bộ thư mục đề:\n{folder}\n\nKhông thể hoàn tác."):
            shutil.rmtree(folder, ignore_errors=True)
            self.refresh()


# ====================================================================== Kết quả học viên


class ResultsPage(_Page):
    """Quản trị: kết quả của mọi người. Giáo viên: chỉ học viên của mình."""

    def __init__(self, shell):
        super().__init__()
        self.store, self.me = shell.main.store, shell.account
        self.people = {a.username: a for a in self.store.visible_to(self.me)}
        self.export_btn = button("Xuất CSV (mở bằng Excel)", self.export)
        sub = "Mọi lượt nộp bài của tất cả tài khoản." if self.me.is_admin else "Các lượt nộp bài của học viên bạn phụ trách."
        self.lay.addWidget(T.page_header("Kết quả học viên", sub, [self.export_btn]))
        self.stats = QHBoxLayout()
        self.stats.setSpacing(16)
        self.lay.addLayout(self.stats)
        self.lop = QComboBox()
        self.lop.addItem("Tất cả lớp", None)
        for c in sorted({a.lop for a in self.people.values() if a.lop}):
            self.lop.addItem(f"Lớp {c}", c)
        self.lop.setMinimumWidth(150)
        self.who = QComboBox()
        self.who.addItem("Tất cả học viên", None)
        for a in self.people.values():
            self.who.addItem(f"{a.ho_ten} ({a.username})", a.username)
        self.who.setMinimumWidth(280)
        for cb in (self.lop, self.who):
            cb.currentIndexChanged.connect(self.refresh)
        self.lay.addLayout(_toolbar(label("Lọc:", "h3"), self.lop, self.who))
        self.table = T.table([("Thời gian", 150), ("Tài khoản", 120), ("Họ tên", 180), ("Lớp", 80),
                              ("Bài thi", None), ("Chế độ", 100), ("Làm trong", 100), ("Điểm", 80), ("Kết quả", 100)])
        self.table.setMinimumHeight(380)
        self.lay.addWidget(self.table, 1)
        self.refresh()

    def _rows(self):
        rows, who, lop = [], self.who.currentData(), self.lop.currentData()
        for h in reversed(core.load_history(who)):
            user = h.get("user")
            acc = self.people.get(user) if user else None
            if not self.me.is_admin and acc is None:          # giáo viên: chỉ học viên của mình
                continue
            if lop and (acc is None or acc.lop != lop):
                continue
            rows.append([h["time"], user or "—", acc.ho_ten if acc else "—", (acc.lop if acc else "") or "—",
                         h["exam"], "Luyện tập" if h["mode"] == "training" else "Thi thử",
                         T.fmt_time(h.get("duration")), h["score"], "Đạt" if h["passed"] else "Chưa đạt"])
        return rows

    def refresh(self):
        rows = self._rows()
        T.set_rows(self.table, rows, ["ok" if r[-1] == "Đạt" else "bad" for r in rows])
        while self.stats.count():
            item = self.stats.takeAt(0)
            if item.widget():
                item.widget().deleteLater()
        passed = sum(r[-1] == "Đạt" for r in rows)
        self.stats.addWidget(T.stat_card("Lượt làm bài", str(len(rows))))
        self.stats.addWidget(T.stat_card("Lượt đạt", str(passed), accent=SUCCESS))
        self.stats.addWidget(T.stat_card("Điểm trung bình",
                                         str(round(sum(r[7] for r in rows) / len(rows))) if rows else "—",
                                         accent="#7A5AF8"))
        self.stats.addWidget(T.stat_card("Tỉ lệ đạt", f"{round(100 * passed / len(rows))}%" if rows else "—",
                                         accent=WARN))

    def export(self):
        _save_csv(self, "ket_qua_MOS.csv", ["Thời gian", "Tài khoản", "Họ tên", "Lớp", "Bài thi", "Chế độ",
                                            "Làm trong", "Điểm", "Kết quả"], self._rows())


# ====================================================================== Soạn đề


def _list() -> QListWidget:
    lw = QListWidget()
    lw.setTextElideMode(Qt.ElideRight)
    lw.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
    return lw


def describe_spec(spec: dict) -> str:
    name = spec.get("luat", "?")
    if name not in rules.RULES:
        return f"⚠ luật không tồn tại: {name}"
    params = {p["name"]: p for p in rules.rule_params(name)}
    parts = [f"{params[k]['label'].split(' (')[0]}: {rules.format_value(params[k]['kind'], v)}"
             for k, v in spec.items() if k != "luat" and k in params]
    return rules.rule_title(name) + (f"   —   {';  '.join(parts)}" if parts else "")


class ExamEditor(QDialog):
    """Form soạn đề: Đề → Dự án (file gốc) → Nhiệm vụ → Luật chấm."""

    def __init__(self, parent, folder: Path | None):
        super().__init__(parent)
        self.folder = folder
        self.data = custom.read_exam_json(folder) if folder else {"mon": "EXCEL", "ten": "", "du_an": []}
        self.p_idx = self.t_idx = None
        self.dirty = False
        self._loading = False
        self.setWindowTitle("Soạn đề")
        self.resize(1280, 800)
        self.setMinimumSize(1100, 680)
        self.setWindowFlag(Qt.WindowMaximizeButtonHint, True)

        root = QVBoxLayout(self)
        root.setContentsMargins(24, 20, 24, 20)
        root.setSpacing(14)

        # --- đầu trang
        head = QHBoxLayout()
        head.setSpacing(10)
        col = QVBoxLayout()
        col.setSpacing(2)
        col.addWidget(label("Soạn đề" if not folder else "Sửa đề", "h1"))
        col.addWidget(label("① Dự án & file gốc  →  ② Nhiệm vụ  →  ③ Luật chấm  →  Lưu & kiểm tra với file đáp án",
                            "muted"))
        head.addLayout(col, 1)
        head.addWidget(button("Lưu && kiểm tra", self.save_and_check))
        head.addWidget(button("Lưu đề", self.save, "primary"))
        root.addLayout(head)

        info = Card(padding=16, spacing=12, horizontal=True)
        info.lay.addWidget(label("Tên đề", "h3"))
        self.ten = QLineEdit(self.data.get("ten", ""))
        self.ten.setPlaceholderText("vd: Excel – Đề luyện số 2")
        self.ten.textChanged.connect(self._mark)
        info.lay.addWidget(self.ten, 2)
        info.lay.addSpacing(12)
        info.lay.addWidget(label("Môn", "h3"))
        self.mon = QComboBox()
        for code, name in MON_NAMES.items():
            self.mon.addItem(name, code)
        self.mon.setCurrentIndex(self.mon.findData(str(self.data.get("mon", "EXCEL")).upper()))
        self.mon.currentIndexChanged.connect(self._mark)
        self.mon.setMinimumWidth(160)
        info.lay.addWidget(self.mon)
        root.addWidget(info)

        split = QSplitter(Qt.Horizontal)
        split.setChildrenCollapsible(False)
        root.addWidget(split, 1)

        # --- cột 1: dự án
        c1 = Card(padding=16, spacing=10)
        c1.lay.addWidget(self._step("1", "Dự án"))
        self.p_list = _list()
        self.p_list.currentRowChanged.connect(self.on_project_select)
        c1.lay.addWidget(self.p_list, 1)
        c1.lay.addLayout(_toolbar(button("+ Thêm dự án", self.add_project, "ghost", "sm"),
                                  right=(button("↑", lambda: self.move_project(-1), size="sm"),
                                         button("↓", lambda: self.move_project(1), size="sm"),
                                         button("Xóa", self.delete_project, "danger", "sm"))))
        split.addWidget(c1)

        # --- cột 2: dự án + nhiệm vụ
        c2 = self.c2 = Card(padding=16, spacing=8)
        c2.lay.addWidget(label("Thông tin dự án", "h3"))
        form = QFormLayout()
        form.setLabelAlignment(Qt.AlignLeft)
        form.setVerticalSpacing(8)
        self.p_ten = QLineEdit()
        self.p_ten.textChanged.connect(self._commit_project)
        self.p_mota = QLineEdit()
        self.p_mota.setPlaceholderText("Tình huống, vd: Bạn là kế toán…")
        self.p_mota.textChanged.connect(self._commit_project)
        form.addRow(label("Tên", "caption"), self.p_ten)
        form.addRow(label("Mô tả", "caption"), self.p_mota)
        c2.lay.addLayout(form)
        self.src_row = self._file_row("File gốc", self.pick_source)
        self.ans_row = self._file_row("File đáp án", self.pick_answer)
        c2.lay.addWidget(self.src_row)
        c2.lay.addWidget(self.ans_row)
        c2.lay.addSpacing(6)
        c2.lay.addWidget(self._step("2", "Nhiệm vụ"))
        self.t_list = _list()
        self.t_list.currentRowChanged.connect(self.on_task_select)
        c2.lay.addWidget(self.t_list, 1)
        c2.lay.addLayout(_toolbar(button("+ Thêm nhiệm vụ", self.add_task, "ghost", "sm"),
                                  right=(button("↑", lambda: self.move_task(-1), size="sm"),
                                         button("↓", lambda: self.move_task(1), size="sm"),
                                         button("Xóa", self.delete_task, "danger", "sm"))))
        split.addWidget(c2)

        # --- cột 3: chi tiết nhiệm vụ + luật chấm
        c3 = self.c3 = Card(padding=16, spacing=8)
        c3.lay.addWidget(label("Chi tiết nhiệm vụ", "h3"))
        c3.lay.addWidget(label("Yêu cầu (học viên nhìn thấy)", "caption"))
        self.t_yc = QPlainTextEdit()
        self.t_yc.setPlaceholderText("vd: Cố định hàng tiêu đề để luôn hiển thị khi cuộn.")
        self.t_yc.setFixedHeight(84)
        self.t_yc.textChanged.connect(self._commit_task)
        c3.lay.addWidget(self.t_yc)
        c3.lay.addWidget(label("Gợi ý cách làm", "caption"))
        self.t_gy = QPlainTextEdit()
        self.t_gy.setPlaceholderText("vd: View > Freeze Panes > Freeze Top Row")
        self.t_gy.setFixedHeight(84)
        self.t_gy.textChanged.connect(self._commit_task)
        c3.lay.addWidget(self.t_gy)
        c3.lay.addSpacing(6)
        c3.lay.addWidget(self._step("3", "Luật chấm  (phải đúng tất cả)"))
        self.r_list = QListWidget()
        self.r_list.setWordWrap(True)
        self.r_list.doubleClicked.connect(lambda *_: self.edit_rule())
        c3.lay.addWidget(self.r_list, 1)
        c3.lay.addLayout(_toolbar(button("+ Thêm luật chấm", self.add_rule, "ghost", "sm"),
                                  right=(button("Sửa", self.edit_rule, size="sm"),
                                         button("Xóa", self.delete_rule, "danger", "sm"))))
        split.addWidget(c3)
        split.setSizes([260, 420, 560])

        self.refresh_projects()
        self._show_project(0 if self.data["du_an"] else None)

    # ------------------------------------------------------------ tiện ích
    @staticmethod
    def _step(num, text) -> QWidget:
        w = QWidget()
        lay = QHBoxLayout(w)
        lay.setContentsMargins(0, 0, 0, 0)
        n = QLabel(num)
        n.setFixedSize(22, 22)
        n.setAlignment(Qt.AlignCenter)
        n.setStyleSheet(f"background:{T.PRIMARY}; color:white; border-radius:11px; font-weight:800; font-size:9pt;")
        lay.addWidget(n)
        lay.addWidget(label(text, "h3"), 1)
        return w

    def _file_row(self, title, on_pick) -> QFrame:
        f = QFrame()
        f.setObjectName("Soft")
        lay = QHBoxLayout(f)
        lay.setContentsMargins(12, 8, 8, 8)
        col = QVBoxLayout()
        col.setSpacing(0)
        col.addWidget(label(title, "caption"))
        f.value = label("")
        f.value.setStyleSheet("font-weight:600; background: transparent;")
        col.addWidget(f.value)
        lay.addLayout(col, 1)
        lay.addWidget(button("Chọn file…", on_pick, size="sm"))
        return f

    def _mark(self, *_):
        self.dirty = True

    def _mon(self) -> str:
        return self.mon.currentData()

    @property
    def project(self):
        return self.data["du_an"][self.p_idx] if self.p_idx is not None else None

    @property
    def task(self):
        p = self.project
        return p["nhiem_vu"][self.t_idx] if p is not None and self.t_idx is not None else None

    # ------------------------------------------------------------ dự án
    def refresh_projects(self):
        self._loading = True
        self.p_list.clear()
        for i, p in enumerate(self.data["du_an"], start=1):
            self.p_list.addItem(f"{i}.  {p.get('ten') or '(chưa đặt tên)'}")
        if self.p_idx is not None:
            self.p_list.setCurrentRow(self.p_idx)
        self._loading = False

    def _show_project(self, idx):
        self.p_idx = idx
        p = self.project
        self._loading = True
        self.p_ten.setText(p.get("ten", "") if p else "")
        self.p_mota.setText(p.get("mo_ta", "") if p else "")
        self._loading = False
        self.c2.setEnabled(p is not None)
        self._refresh_files()
        self.refresh_projects()
        self.t_idx = None
        self.refresh_tasks()
        self._show_task(0 if p and p["nhiem_vu"] else None)

    def _refresh_files(self):
        p = self.project
        if not p:
            self.src_row.value.setText("—")
            self.ans_row.value.setText("—")
            return
        src = p.get("_src") or (self.folder / p["file"] if self.folder and p.get("file") else None)
        self.src_row.value.setText(Path(src).name if src else "Chưa chọn")
        self.src_row.value.setStyleSheet(f"font-weight:600; background:transparent; color:{T.TEXT if src else DANGER};")
        ans = p.get("_ans") or (self.folder / "dap_an" / p["file"] if self.folder and p.get("file") else None)
        has = bool(ans) and Path(ans).is_file()
        self.ans_row.value.setText(Path(ans).name if has else "Chưa có — nên chọn để kiểm tra đề")
        self.ans_row.value.setStyleSheet(f"font-weight:600; background:transparent; color:{SUCCESS if has else MUTED};")

    def _commit_project(self):
        if self._loading or not self.project:
            return
        self.project["ten"] = self.p_ten.text().strip()
        self.project["mo_ta"] = self.p_mota.text().strip()
        self._mark()
        item = self.p_list.item(self.p_idx)
        if item:
            item.setText(f"{self.p_idx + 1}.  {self.project['ten'] or '(chưa đặt tên)'}")

    def on_project_select(self, row):
        if not self._loading and row >= 0 and row != self.p_idx:
            self._show_project(row)

    def add_project(self):
        n = len(self.data["du_an"]) + 1
        self.data["du_an"].append({"ten": f"Dự án {n}", "file": "", "mo_ta": "", "nhiem_vu": []})
        self._mark()
        self._show_project(n - 1)
        self.pick_source()

    def delete_project(self):
        if self.project and T.confirm(self, "Xóa dự án", f"Xóa dự án “{self.project['ten']}”?"):
            del self.data["du_an"][self.p_idx]
            self._mark()
            self._show_project(min(self.p_idx, len(self.data["du_an"]) - 1) if self.data["du_an"] else None)

    def move_project(self, step):
        lst, i = self.data["du_an"], self.p_idx
        if i is not None and 0 <= i + step < len(lst):
            lst[i], lst[i + step] = lst[i + step], lst[i]
            self._mark()
            self._show_project(i + step)

    def _pick(self, title):
        path, _ = QFileDialog.getOpenFileName(self, title, "", FILE_FILTERS[self._mon()])
        return path

    def pick_source(self):
        if self.project:
            path = self._pick("Chọn file gốc (file chưa làm)")
            if path:
                self.project["_src"] = path
                self.project["file"] = Path(path).name
                self._mark()
                self._refresh_files()

    def pick_answer(self):
        if self.project:
            path = self._pick("Chọn file đáp án (đã làm đúng hết)")
            if path:
                self.project["_ans"] = path
                self._mark()
                self._refresh_files()

    # ------------------------------------------------------------ nhiệm vụ
    def refresh_tasks(self):
        self._loading = True
        self.t_list.clear()
        for i, t in enumerate(self.project["nhiem_vu"] if self.project else [], start=1):
            self.t_list.addItem(f"{i}.  {t.get('yeu_cau') or '(chưa có yêu cầu)'}")
        if self.t_idx is not None:
            self.t_list.setCurrentRow(self.t_idx)
        self._loading = False

    def _show_task(self, idx):
        self.t_idx = idx
        t = self.task
        self._loading = True
        self.t_yc.setPlainText(t.get("yeu_cau", "") if t else "")
        self.t_gy.setPlainText(t.get("goi_y", "") if t else "")
        self._loading = False
        self.c3.setEnabled(t is not None)
        self.refresh_tasks()
        self.refresh_rules()

    def _commit_task(self):
        t = self.task
        if self._loading or not t:
            return
        t["yeu_cau"] = self.t_yc.toPlainText().strip()
        t["goi_y"] = self.t_gy.toPlainText().strip()
        self._mark()
        item = self.t_list.item(self.t_idx)
        if item:
            item.setText(f"{self.t_idx + 1}.  {t['yeu_cau'] or '(chưa có yêu cầu)'}")

    def on_task_select(self, row):
        if not self._loading and row >= 0 and row != self.t_idx:
            self._show_task(row)

    def add_task(self):
        if self.project:
            self.project["nhiem_vu"].append({"yeu_cau": "", "goi_y": "", "cham": []})
            self._mark()
            self._show_task(len(self.project["nhiem_vu"]) - 1)
            self.t_yc.setFocus()

    def delete_task(self):
        if self.task and T.confirm(self, "Xóa nhiệm vụ", "Xóa nhiệm vụ đang chọn?"):
            tasks = self.project["nhiem_vu"]
            del tasks[self.t_idx]
            self._mark()
            self._show_task(min(self.t_idx, len(tasks) - 1) if tasks else None)

    def move_task(self, step):
        if self.project and self.t_idx is not None:
            lst, i = self.project["nhiem_vu"], self.t_idx
            if 0 <= i + step < len(lst):
                lst[i], lst[i + step] = lst[i + step], lst[i]
                self._mark()
                self._show_task(i + step)

    # ------------------------------------------------------------ luật chấm
    def _rules(self):
        t = self.task
        if t is None:
            return []
        t["cham"] = rules._as_list(t.get("cham") or [])
        return t["cham"]

    def refresh_rules(self):
        self.r_list.clear()
        for spec in self._rules():
            self.r_list.addItem(describe_spec(spec))

    def add_rule(self):
        if self.task is not None:
            dlg = RuleDialog(self, self._mon(), None)
            if dlg.exec():
                self._rules().append(dlg.spec)
                self._mark()
                self.refresh_rules()

    def edit_rule(self):
        row = self.r_list.currentRow()
        if row >= 0:
            dlg = RuleDialog(self, self._mon(), self._rules()[row])
            if dlg.exec():
                self._rules()[row] = dlg.spec
                self._mark()
                self.refresh_rules()

    def delete_rule(self):
        row = self.r_list.currentRow()
        if row >= 0:
            del self._rules()[row]
            self._mark()
            self.refresh_rules()

    # ------------------------------------------------------------ lưu
    def _validate(self) -> list[str]:
        problems = []
        if not self.ten.text().strip():
            problems.append("Chưa đặt tên đề.")
        if not self.data["du_an"]:
            problems.append("Đề chưa có dự án nào.")
        names = [p.get("file", "").casefold() for p in self.data["du_an"]]
        for i, p in enumerate(self.data["du_an"], start=1):
            where = f"Dự án {i}"
            if not p.get("file"):
                problems.append(f"{where}: chưa chọn file gốc.")
            elif names.count(p["file"].casefold()) > 1:
                problems.append(f"{where}: trùng tên file gốc “{p['file']}” với dự án khác.")
            if not p["nhiem_vu"]:
                problems.append(f"{where}: chưa có nhiệm vụ.")
            for j, t in enumerate(p["nhiem_vu"], start=1):
                if not t.get("yeu_cau"):
                    problems.append(f"{where} › nhiệm vụ {j}: chưa ghi yêu cầu.")
                if not t.get("cham"):
                    problems.append(f"{where} › nhiệm vụ {j}: chưa có luật chấm.")
        return problems

    def save(self, quiet=False) -> bool:
        problems = self._validate()
        if problems:
            T.warn(self, "Chưa lưu được", "\n".join("• " + p for p in problems[:12]))
            return False
        self.data["ten"] = self.ten.text().strip()
        self.data["mon"] = self._mon()
        folder = self.folder or custom.new_exam_folder(self.data["ten"])
        try:
            folder.mkdir(parents=True, exist_ok=True)
            for p in self.data["du_an"]:
                if p.get("_src") and Path(p["_src"]).resolve() != (folder / p["file"]).resolve():
                    shutil.copyfile(p["_src"], folder / p["file"])
                if p.get("_ans"):
                    (folder / "dap_an").mkdir(exist_ok=True)
                    shutil.copyfile(p["_ans"], folder / "dap_an" / p["file"])
            clean = {**self.data, "du_an": [{k: v for k, v in p.items() if not k.startswith("_")}
                                            for p in self.data["du_an"]]}
            custom.save_exam_json(folder, clean)
        except OSError as exc:
            T.error(self, "Lỗi khi lưu", str(exc))
            return False
        for p in self.data["du_an"]:
            p.pop("_src", None)
            p.pop("_ans", None)
        self.folder = folder
        self.dirty = False
        self._refresh_files()
        if not quiet:
            T.info(self, "Đã lưu", f"Đã lưu đề vào:\n{folder}")
        return True

    def save_and_check(self):
        if self.save(quiet=True):
            CheckReportDialog(self, self.folder).exec()

    def reject(self):
        if self.dirty and not T.confirm(self, "Chưa lưu", "Đề có thay đổi chưa lưu. Đóng mà không lưu?"):
            return
        super().reject()


class RuleDialog(_Dialog):
    """Chọn luật chấm và điền tham số – form tự sinh theo luật."""

    def __init__(self, parent, mon: str, spec: dict | None):
        super().__init__(parent, "Luật chấm", 620)
        self.names = rules.rules_for(mon)
        self.spec = None
        self.lay.addWidget(label("Kiểm tra điều gì?", "h3"))
        self.choice = QComboBox()
        for n in self.names:
            self.choice.addItem(rules.rule_title(n), n)
        self.choice.setMaxVisibleItems(18)
        self.lay.addWidget(self.choice)
        self.doc = label("", "caption", wrap=True)
        self.lay.addWidget(self.doc)
        self.form_box = Card(padding=18, spacing=6)
        self.form = QVBoxLayout()
        self.form.setSpacing(6)
        self.form_box.lay.addLayout(self.form)
        self.lay.addWidget(self.form_box)
        self.lay.addStretch()
        self.buttons("Xong", self.ok)
        start = spec.get("luat") if spec and spec.get("luat") in self.names else self.names[0]
        self.choice.setCurrentIndex(self.names.index(start))
        self.build_form(spec if spec and spec.get("luat") == start else None)
        self.choice.currentIndexChanged.connect(lambda *_: self.build_form())

    @property
    def name(self):
        return self.choice.currentData()

    def build_form(self, spec=None):
        while self.form.count():
            item = self.form.takeAt(0)
            if item.widget():
                item.widget().deleteLater()
        self.widgets = {}
        params = rules.rule_params(self.name)
        doc = " ".join((rules.RULES[self.name].__doc__ or "").split())
        for p in params:
            doc = doc.replace(f"`{p['name']}`", "«" + p["label"].split(" (")[0].lower() + "»")
        self.doc.setText(doc.replace("`", ""))
        for p in params:
            value = spec.get(p["name"]) if spec else None
            if value is None and not p["required"]:
                value = p["default"]
            text = rules.format_value(p["kind"], value)
            if p["kind"] in ("choice", "bool"):
                w = QComboBox()
                options = p["choices"] if p["kind"] == "choice" else list(rules.BOOL_TEXT)
                if not p["required"] and p["default"] is None:
                    w.addItem("(bỏ qua)", "")
                for o in options:
                    w.addItem(o, o)
                w.setCurrentIndex(max(0, w.findData(text)))
            else:
                w = QLineEdit(text)
            cap = label(p["label"] + ("" if p["required"] else "   <span style='color:#98A2B3'>(tuỳ chọn)</span>"),
                        wrap=True)
            cap.setStyleSheet("font-weight: 600; margin-top: 6px;")
            self.form.addWidget(cap)
            self.form.addWidget(w)
            self.widgets[p["name"]] = (w, p)

    def ok(self):
        spec = {"luat": self.name}
        try:
            for name, (w, p) in self.widgets.items():
                raw = w.currentData() if isinstance(w, QComboBox) else w.text()
                value = rules.parse_value(p["kind"], raw)
                if value is None:
                    if p["required"]:
                        raise ValueError(f"Chưa điền “{p['label']}”.")
                    continue
                if value != p["default"]:
                    spec[name] = value
        except ValueError as exc:
            T.warn(self, "Thiếu thông tin", str(exc))
            return
        err = rules.validate(spec)
        if err:
            T.error(self, "Lỗi", err)
            return
        self.spec = spec
        self.accept()


class CheckReportDialog(_Dialog):
    """Kiểm tra đề: file gốc phải chấm SAI, file đáp án phải chấm ĐÚNG."""

    def __init__(self, parent, folder: Path):
        super().__init__(parent, "Kiểm tra đề", 900)
        errors, rows = custom.check_exam(folder)
        if errors:
            self.lay.addWidget(_banner("Đề còn lỗi khai báo:<br>" + "<br>".join("• " + e for e in errors), "bad"))
            self.buttons("Đóng", cancel=False)
            return
        problems = sum(r["start_ok"] for r in rows) + sum(r["answer_ok"] is False for r in rows)
        missing = sum(r["answer_ok"] is None for r in rows)
        head = QHBoxLayout()
        if problems:
            head.addWidget(chip(f"CẦN XEM LẠI {problems} CHỖ", DANGER, DANGER_SOFT))
        elif missing:
            head.addWidget(chip("CHƯA ĐỦ ĐÁP ÁN", WARN, WARN_SOFT))
        else:
            head.addWidget(chip("ĐỀ ỔN", SUCCESS, SUCCESS_SOFT))
        head.addWidget(label("File gốc (chưa làm) phải chấm SAI · file đáp án (đã làm đúng) phải chấm ĐÚNG.",
                             "muted"), 1)
        self.lay.addLayout(head)
        t = T.table([("Dự án", 200), ("Nhiệm vụ", None), ("File gốc", 130), ("File đáp án", 130)])
        data, tones = [], []
        for r in rows:
            data.append([r["project"], f"{r['index']}. {r['task']}",
                         "ĐÚNG sẵn ⚠" if r["start_ok"] else "sai ✓",
                         "chưa có" if r["answer_ok"] is None else ("đúng ✓" if r["answer_ok"] else "SAI ✗")])
            tones.append("bad" if r["start_ok"] or r["answer_ok"] is False
                         else "muted" if r["answer_ok"] is None else "ok")
        T.set_rows(t, data, tones, tone_cols={2, 3})
        t.setMinimumHeight(min(80 + 42 * len(rows), 420))
        self.lay.addWidget(t)
        tips = []
        if any(r["start_ok"] for r in rows):
            tips.append("<b>“ĐÚNG sẵn”</b>: file chưa làm đã đạt — luật chấm quá dễ, hãy siết điều kiện.")
        if any(r["answer_ok"] is False for r in rows):
            tips.append("<b>“SAI” ở đáp án</b>: làm đúng vẫn bị chấm sai — kiểm tra lại tham số của luật.")
        if missing:
            tips.append("Chọn <b>file đáp án</b> cho từng dự án để kiểm tra chiều đúng.")
        if tips:
            self.lay.addWidget(_banner("<br>".join(tips), "warn" if not problems else "bad"))
        self.buttons("Đóng", cancel=False)
