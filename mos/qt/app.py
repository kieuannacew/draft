"""Ứng dụng Luyện thi MOS – giao diện Qt (PySide6).

Cấu trúc màn hình:
  MainWindow
   ├─ LoginPage                 đăng nhập (có nút chuyển VI | EN)
   └─ Shell                     thanh bên (menu) + trang nội dung
        ├─ HomePage             cấp độ / XP / chuỗi ngày, huy hiệu, chọn bài thi + chế độ → Bắt đầu
        ├─ HistoryPage          lịch sử của tôi
        ├─ ResultPage           kết quả sau khi nộp bài (pháo giấy khi đạt, XP nhận được, huy hiệu mới)
        └─ (quản trị) AccountsPage / ExamsPage / ResultsPage   – xem admin.py
  ExamBar                       cửa sổ làm bài nằm ở cạnh dưới màn hình, luôn nổi trên Office

Mọi chữ trên giao diện bọc bằng tr(...) (mos/i18n.py); nội dung đề chọn bằng pick(vi, en).
"""
from __future__ import annotations

import html
import sys
import time

from PySide6.QtCore import Qt, QTimer
from PySide6.QtGui import QColor, QFont, QGuiApplication, QIcon, QPainter, QPixmap
from PySide6.QtWidgets import (QApplication, QButtonGroup, QDialog, QFrame, QGridLayout, QHBoxLayout, QLabel,
                               QLineEdit, QMainWindow, QPushButton, QScrollArea, QSizePolicy, QStackedWidget,
                               QVBoxLayout, QWidget)

from .. import chuong, core, gamify, i18n, lop_hoc, nhap_de
from ..accounts import ROLE_NAMES, AccountError, AccountStore
from ..custom import load_custom_exams, merge_exams
from ..exams import ALL_EXAMS
from ..i18n import pick, tr
from . import theme as T
from .theme import (DANGER, DANGER_SOFT, GOLD_SOFT, PRIMARY, PRIMARY_SOFT, SUCCESS, SUCCESS_SOFT, WARN,
                    WARN_SOFT, Card, ClickableCard, button, chip, label)

def mode_name(mode: str) -> str:
    return {"training": tr("Luyện tập"), "chapter": tr("Luyện theo chương")}.get(mode, tr("Thi thử"))


def exam_name(exam) -> str:
    return pick(exam.name, exam.name_en)


def history_exam_name(entry: dict) -> str:
    return pick(entry.get("exam", ""), entry.get("exam_en"))


def esc(text: str) -> str:
    return html.escape(text or "", quote=False)


# ====================================================================== khung trang


def scroll_page(content: QWidget) -> QScrollArea:
    area = QScrollArea()
    area.setObjectName("Page")
    area.setWidgetResizable(True)
    area.setFrameShape(QFrame.NoFrame)
    area.setWidget(content)
    return area


def page_body(margins=36, spacing=22) -> tuple[QWidget, QVBoxLayout]:
    w = QWidget()
    w.setObjectName("Page")
    lay = QVBoxLayout(w)
    lay.setContentsMargins(margins, margins - 6, margins, margins)
    lay.setSpacing(spacing)
    return w, lay


# ====================================================================== đăng nhập


class LoginPage(QWidget):
    def __init__(self, main: "MainWindow"):
        super().__init__()
        self.main = main
        self.setObjectName("Login")
        self.setAttribute(Qt.WA_StyledBackground, True)
        outer = QVBoxLayout(self)
        outer.setContentsMargins(24, 18, 24, 24)

        top = QHBoxLayout()
        top.addStretch()
        top.addWidget(T.lang_switch(i18n.get_lang(), main.change_language, dark=True))
        outer.addLayout(top)
        outer.addStretch(1)

        card = Card(padding=38, spacing=0, raised=True)
        card.setFixedWidth(440)
        outer.addWidget(card, 0, Qt.AlignCenter)
        lay = card.lay

        badges = QHBoxLayout()
        badges.setSpacing(8)
        for code in ("WORD", "EXCEL", "POWERPOINT"):
            badges.addWidget(T.badge(code, 36))
        badges.addStretch()
        leaf = QLabel("🌿")
        leaf.setStyleSheet("font-size: 20pt; background: transparent;")
        badges.addWidget(leaf)
        lay.addLayout(badges)
        lay.addSpacing(22)
        lay.addWidget(label(tr("Luyện thi MOS"), "h1"))
        lay.addSpacing(4)
        lay.addWidget(label(tr("Mỗi ngày một chút – vững vàng bước vào phòng thi."), "muted", wrap=True))
        lay.addSpacing(24)

        lay.addWidget(label(tr("Tên đăng nhập"), "h3"))
        lay.addSpacing(6)
        self.user = QLineEdit()
        self.user.setProperty("size", "lg")
        self.user.setPlaceholderText(tr("vd: nguyenvana"))
        lay.addWidget(self.user)
        lay.addSpacing(14)
        lay.addWidget(label(tr("Mật khẩu"), "h3"))
        lay.addSpacing(6)
        self.pw = QLineEdit()
        self.pw.setProperty("size", "lg")
        self.pw.setEchoMode(QLineEdit.Password)
        self.pw.setPlaceholderText("••••••")
        lay.addWidget(self.pw)
        lay.addSpacing(8)
        self.err = label("", wrap=True)
        self.err.setStyleSheet(f"color:{DANGER};")
        self.err.hide()
        lay.addWidget(self.err)
        lay.addSpacing(14)
        lay.addWidget(button(tr("Đăng nhập"), self.login, "primary", "lg"))

        self.server = lop_hoc.load_server()
        store = main.store
        admin = store.get("admin")
        if self.server:
            lay.addSpacing(10)
            lay.addWidget(button(tr("Chưa có tài khoản? Đăng ký bằng mã lớp"), self.register, "ghost"))
        elif admin and admin.doi_mat_khau and len(store.accounts) == 1:
            lay.addSpacing(16)
            tip = QFrame()
            tip.setObjectName("Banner")
            tl = QVBoxLayout(tip)
            tl.setContentsMargins(14, 10, 14, 10)
            tl.addWidget(label(tr("<b>Lần đầu sử dụng?</b><br>Đăng nhập <b>admin</b> / <b>admin</b> "
                                  "rồi đặt mật khẩu mới."), wrap=True))
            lay.addWidget(tip)

        outer.addSpacing(18)
        foot = label(tr("Chấm điểm tự động • Word • Excel • PowerPoint"), "caption")
        foot.setStyleSheet("color: rgba(255,255,255,0.7);")
        outer.addWidget(foot, 0, Qt.AlignCenter)
        outer.addStretch(1)
        bottom = QHBoxLayout()
        state = label(("🌐  " + tr("Lớp học trực tuyến")) if self.server else ("💻  " + tr("Dùng trên máy này")),
                      "caption")
        state.setStyleSheet("color: rgba(255,255,255,0.7);")
        bottom.addWidget(state)
        bottom.addStretch()
        bottom.addWidget(button("⚙  " + tr("Máy chủ lớp học"), self.server_settings, "side", "sm"))
        outer.addLayout(bottom)

        self.user.returnPressed.connect(self.pw.setFocus)
        self.pw.returnPressed.connect(self.login)

    def showEvent(self, e):
        super().showEvent(e)
        self.user.setFocus()

    def _error(self, text: str) -> None:
        self.err.setText(text)
        self.err.show()

    def login(self):
        if self.server:
            return self._login_online(self.user.text().strip().lower(), self.pw.text())
        try:
            acc = self.main.store.authenticate(self.user.text(), self.pw.text())
        except AccountError as exc:
            self._error(tr(str(exc)))
            return
        self.err.hide()
        self.pw.clear()
        self.main.enter(acc)

    def _login_online(self, user: str, pw: str):
        """Đăng nhập máy chủ lớp học; mất mạng thì dùng bản sao tài khoản đã lưu trên máy."""
        from .lop import err_text
        cloud = lop_hoc.Cloud(**self.server)
        QApplication.setOverrideCursor(Qt.WaitCursor)
        try:
            profile = cloud.login(user, pw)
        except lop_hoc.OfflineError:
            QApplication.restoreOverrideCursor()
            acc = self.main.store.get(user)
            if acc is None or acc.nguon != "may_chu":
                return self._error(tr("Không kết nối được máy chủ lớp học. Kiểm tra mạng Internet."))
            try:
                acc = self.main.store.authenticate(user, pw)
            except AccountError as exc:
                return self._error(tr(str(exc)))
            self.pw.clear()
            T.info(self, tr("Ngoại tuyến"), tr("Không kết nối được máy chủ – bạn đang học ngoại tuyến. Bài làm được lưu "
                                               "trên máy và tự gửi cho giáo viên khi đăng nhập lại lúc có mạng."))
            return self.main.enter(acc)
        except lop_hoc.CloudError as exc:
            QApplication.restoreOverrideCursor()
            return self._error(err_text(exc))
        QApplication.restoreOverrideCursor()
        self.err.hide()
        self.pw.clear()
        self.main.enter_online(cloud, profile, pw)

    def register(self):
        from .lop import RegisterDialog
        dlg = RegisterDialog(self, self.server)
        if dlg.exec() and dlg.result_data:
            cloud, profile, _, pw = dlg.result_data
            self.main.enter_online(cloud, profile, pw)

    def server_settings(self):
        from .lop import ServerDialog
        if ServerDialog(self).exec():
            self.main.logout()


class ChangePasswordDialog(QDialog):
    def __init__(self, parent, store: AccountStore, acc, forced=False, cloud=None):
        super().__init__(parent)
        self.store, self.acc, self.forced, self.cloud = store, acc, forced, cloud
        self.setWindowTitle(tr("Đổi mật khẩu"))
        self.setMinimumWidth(420)
        lay = QVBoxLayout(self)
        lay.setContentsMargins(28, 24, 28, 24)
        lay.setSpacing(8)
        lay.addWidget(label(tr("Đổi mật khẩu"), "h2"))
        if forced:
            lay.addWidget(label(tr("Đây là lần đăng nhập đầu tiên — hãy đặt mật khẩu mới cho tài khoản "
                                   "<b>{user}</b>.").format(user=esc(acc.username)), "muted", wrap=True))
        lay.addSpacing(8)
        self.old = self._field(lay, tr("Mật khẩu hiện tại")) if not forced else None
        self.new = self._field(lay, tr("Mật khẩu mới"))
        self.again = self._field(lay, tr("Nhập lại mật khẩu mới"))
        self.err = label("", wrap=True)
        self.err.setStyleSheet(f"color:{DANGER};")
        lay.addWidget(self.err)
        row = QHBoxLayout()
        row.addStretch()
        if not forced:
            row.addWidget(button(tr("Hủy"), self.reject))
        row.addWidget(button(tr("Lưu mật khẩu"), self.save, "primary"))
        lay.addLayout(row)
        self.again.returnPressed.connect(self.save)

    @staticmethod
    def _field(lay, text):
        lay.addWidget(label(text, "h3"))
        e = QLineEdit()
        e.setEchoMode(QLineEdit.Password)
        lay.addWidget(e)
        lay.addSpacing(4)
        return e

    def reject(self):
        if not self.forced:
            super().reject()

    def save(self):
        try:
            if self.old is not None:
                self.store.authenticate(self.acc.username, self.old.text())
            if self.new.text() != self.again.text():
                raise AccountError("Hai lần nhập mật khẩu mới không giống nhau.")
            if self.forced and self.new.text() == "admin":
                raise AccountError("Hãy chọn mật khẩu khác mật khẩu mặc định.")
            if self.acc.nguon == "may_chu" and lop_hoc.load_server():   # tài khoản lớp học: đổi trên máy chủ trước
                if self.cloud is None:
                    raise AccountError("Cần kết nối máy chủ lớp học để đổi mật khẩu.")
                if len(self.new.text()) < lop_hoc.MIN_PASSWORD:
                    raise AccountError("Mật khẩu phải có ít nhất 6 ký tự.")
                self.cloud.change_password(self.new.text())
            self.store.set_password(self.acc.username, self.new.text())
        except AccountError as exc:
            self.err.setText(tr(str(exc)))
            return
        except lop_hoc.CloudError as exc:
            from .lop import err_text
            self.err.setText(err_text(exc))
            return
        self.accept()


# ====================================================================== khung ứng dụng


class Shell(QWidget):
    """Thanh bên + vùng nội dung."""

    def __init__(self, main: "MainWindow", account, page: str = "home"):
        super().__init__()
        self.main, self.account = main, account
        lay = QHBoxLayout(self)
        lay.setContentsMargins(0, 0, 0, 0)
        lay.setSpacing(0)
        lay.addWidget(self._sidebar())
        self.content = QStackedWidget()
        lay.addWidget(self.content, 1)
        self.current = None
        self.go(page if page in ("home", "history", "lessons", "search", "lop", "accounts", "exams", "results")
                else "home")

    # ------------------------------------------------------------ thanh bên
    def _sidebar(self) -> QWidget:
        side = QWidget()
        side.setObjectName("Sidebar")
        side.setAttribute(Qt.WA_StyledBackground, True)
        side.setFixedWidth(252)
        lay = QVBoxLayout(side)
        lay.setContentsMargins(16, 22, 16, 18)
        lay.setSpacing(4)

        brand_row = QHBoxLayout()
        brand_row.setSpacing(10)
        logo = QLabel()
        logo.setPixmap(app_icon().pixmap(36, 36))
        brand_row.addWidget(logo)
        names = QVBoxLayout()
        names.setSpacing(0)
        title = label(tr("Luyện thi MOS"))
        title.setStyleSheet(f"font-family: '{T.pick_heading_family()}'; font-size: 14pt; font-weight: 800;")
        names.addWidget(title)
        names.addWidget(label("Word · Excel · PowerPoint", "caption"))
        brand_row.addLayout(names, 1)
        lay.addLayout(brand_row)
        lay.addSpacing(26)

        self.nav = QButtonGroup(self)
        self.nav.setExclusive(True)
        self.nav_buttons = {}

        def nav(key, icon, text):
            b = QPushButton(f"{icon}   {text}")
            b.setProperty("nav", True)
            b.setCheckable(True)
            b.setCursor(Qt.PointingHandCursor)
            b.clicked.connect(lambda: self.go(key))
            self.nav.addButton(b)
            self.nav_buttons[key] = b
            lay.addWidget(b)

        lay.addWidget(label(tr("HỌC TẬP"), "overline"))
        lay.addSpacing(4)
        nav("home", "🏠", tr("Trang chủ"))
        nav("lessons", "📖", tr("Tài liệu học"))
        nav("search", "🔎", tr("Tra từ khóa"))
        nav("history", "📈", tr("Lịch sử của tôi"))
        if self.main.online:
            nav("lop", "🏫", tr("Lớp học"))
        if self.account.is_admin:
            lay.addSpacing(18)
            lay.addWidget(label(tr("QUẢN TRỊ"), "overline"))
            lay.addSpacing(4)
            nav("accounts", "👥", tr("Tài khoản"))
            nav("exams", "📝", tr("Đề thi"))
            if not self.main.online:
                nav("results", "📊", tr("Kết quả học viên"))
        lay.addStretch()

        lang_row = QHBoxLayout()
        lang_row.addWidget(label(tr("Ngôn ngữ"), "caption"))
        lang_row.addStretch()
        lang_row.addWidget(T.lang_switch(i18n.get_lang(), self.main.change_language, dark=True))
        lay.addLayout(lang_row)
        lay.addSpacing(10)

        me = QFrame()
        me.setObjectName("Me")
        me.setStyleSheet("QFrame#Me { background: rgba(255,255,255,0.07); border-radius: 14px; }")
        ml = QVBoxLayout(me)
        ml.setContentsMargins(12, 12, 12, 12)
        top = QHBoxLayout()
        top.addWidget(T.avatar(self.account.ho_ten, 36, T.OLIVE if self.account.is_admin else T.SAGE))
        names = QVBoxLayout()
        names.setSpacing(0)
        n = label(self.account.ho_ten)
        n.setStyleSheet("font-weight: 700;")
        names.addWidget(n)
        names.addWidget(label(f"{self.account.username} · {i18n.role_name(ROLE_NAMES[self.account.vai_tro])}",
                              "caption"))
        top.addLayout(names, 1)
        ml.addLayout(top)
        if self.main.online:
            net = label(("🌐  " + tr("Đã kết nối lớp học")) if self.main.cloud else ("📡  " + tr("Ngoại tuyến")),
                        "caption")
            ml.addWidget(net)
        row = QHBoxLayout()
        row.addWidget(button(tr("Mật khẩu"), self.change_password, "side"))
        row.addWidget(button(tr("Đăng xuất"), self.main.logout, "side"))
        ml.addLayout(row)
        lay.addWidget(me)
        return side

    # ------------------------------------------------------------ điều hướng
    def go(self, key: str, **kw) -> None:
        from . import admin, hoc, lop, tra_cuu
        factories = {
            "lop": lambda: lop.ClassPage(self, **kw),
            "lessons": lambda: hoc.LessonsPage(self, **kw),
            "lesson": lambda: hoc.open_viewer(self, **kw),
            "search": lambda: tra_cuu.SearchPage(self, **kw),
            "home": lambda: HomePage(self),
            "history": lambda: HistoryPage(self),
            "result": lambda: ResultPage(self, **kw),
            "accounts": lambda: (lop.OnlineAccountsPage(self) if self.main.online else admin.AccountsPage(self)),
            "exams": lambda: admin.ExamsPage(self),
            "results": lambda: admin.ResultsPage(self),
        }
        page = factories[key]()
        old = self.content.currentWidget()
        self.content.addWidget(page)
        self.content.setCurrentWidget(page)
        if old is not None:
            self.content.removeWidget(old)
            old.deleteLater()
        self.current = key
        btn = self.nav_buttons.get("lessons" if key == "lesson" else key)
        if btn:
            btn.setChecked(True)
        else:
            self.nav.setExclusive(False)
            for b in self.nav_buttons.values():
                b.setChecked(False)
            self.nav.setExclusive(True)

    def change_password(self):
        ChangePasswordDialog(self, self.main.store, self.account, cloud=self.main.cloud).exec()

    # ------------------------------------------------------------ gửi dữ liệu lên lớp học
    def sync_lesson(self, lesson, xem: int, xong: bool, diem: str = "") -> None:
        if self.main.online:
            lop_hoc.queue_progress(self.account.username, lesson.id, lesson.ten, xem, lesson.count, xong, diem)
            lop_hoc.flush_in_background(self.main.cloud, self.account.username)

    # ------------------------------------------------------------ làm bài
    def start_exam(self, exam, mode) -> None:
        try:
            session = core.new_session(exam, mode, user=self.account.username)
        except Exception:
            T.error(self, tr("Lỗi"), tr("Không tạo được file bài thi:") + "\n" + core.format_exception())
            return
        self.main.hide()
        self.bar = ExamBar(session, on_finish=self.show_result, on_quit=self.main.show)
        self.bar.show()

    def show_result(self, session, report) -> None:
        if self.main.online:
            try:
                lop_hoc.queue_result(self.account.username, core.history_entry(report))
            except OSError:
                pass
            lop_hoc.flush_in_background(self.main.cloud, self.account.username)
        self.main.show()
        self.main.raise_()
        self.main.activateWindow()
        self.go("result", session=session, report=report)


class MainWindow(QMainWindow):
    def __init__(self, store: AccountStore):
        super().__init__()
        self.store = store
        self.account = None
        self.cloud: lop_hoc.Cloud | None = None      # phiên máy chủ lớp học (None = ngoại tuyến)
        self.setWindowTitle(tr("Luyện thi MOS"))
        self.setWindowIcon(app_icon())
        self.resize(1280, 820)
        self.setMinimumSize(1080, 700)
        self.stack = QStackedWidget()
        self.setCentralWidget(self.stack)
        self.logout()

    def _set(self, widget):
        old = self.stack.currentWidget()
        self.stack.addWidget(widget)
        self.stack.setCurrentWidget(widget)
        if old is not None:
            self.stack.removeWidget(old)
            old.deleteLater()

    @property
    def online(self) -> bool:
        """Đang dùng chế độ Lớp học trực tuyến (đã cấu hình máy chủ)."""
        return lop_hoc.load_server() is not None

    def logout(self):
        if self.cloud:
            self.cloud.logout()
        self.account = None
        self.cloud = None
        self._set(LoginPage(self))

    def enter_online(self, cloud, profile: dict, password: str):
        """Đăng nhập máy chủ thành công: lưu bản sao tài khoản trên máy rồi vào app, gửi dữ liệu đang chờ."""
        acc = self.store.mirror(profile["username"], profile.get("ho_ten", ""), profile.get("vai_tro", ""), password)
        self.cloud = cloud
        self.enter(acc)
        lop_hoc.flush_in_background(cloud, acc.username)

    def enter(self, acc, page: str = "home"):
        if acc.doi_mat_khau:
            if ChangePasswordDialog(self, self.store, acc, forced=True).exec() != QDialog.Accepted:
                return
        self.account = self.store.get(acc.username)
        self._set(Shell(self, self.account, page))

    def change_language(self, code: str) -> None:
        """Đổi ngôn ngữ và dựng lại màn hình đang xem."""
        i18n.set_lang(code)
        self.setWindowTitle(tr("Luyện thi MOS"))
        current = self.stack.currentWidget()
        if self.account is None:
            self._set(LoginPage(self))
        else:
            page = current.current if isinstance(current, Shell) else "home"
            self._set(Shell(self, self.account, page))


def app_icon() -> QIcon:
    pix = QPixmap(64, 64)
    pix.fill(Qt.transparent)
    p = QPainter(pix)
    p.setRenderHint(QPainter.Antialiasing)
    p.setBrush(QColor(T.OLIVE))
    p.setPen(Qt.NoPen)
    p.drawRoundedRect(0, 0, 64, 64, 16, 16)
    p.setBrush(QColor(T.FOREST))
    p.drawRoundedRect(4, 4, 56, 56, 13, 13)
    p.setPen(QColor("#F4F6EE"))
    f = QFont(T.pick_heading_family())
    f.setBold(True)
    f.setPixelSize(20)
    p.setFont(f)
    p.drawText(pix.rect(), Qt.AlignCenter, "MOS")
    p.end()
    return QIcon(pix)


def _banner(text, tone="info") -> QFrame:
    f = QFrame()
    f.setObjectName("Note")
    colors = {"info": (T.FOREST, PRIMARY_SOFT, T.PRIMARY_LINE), "warn": (WARN, WARN_SOFT, "#E8C98A"),
              "bad": (DANGER, DANGER_SOFT, "#E2B5AA"), "ok": (SUCCESS, SUCCESS_SOFT, "#A9CFA0"),
              "gold": ("#7A5A12", GOLD_SOFT, "#EAD9A6")}[tone]
    f.setStyleSheet(f"QFrame#Note {{ background:{colors[1]}; border:1px solid {colors[2]}; border-radius:12px; }}"
                    f"QFrame#Note QLabel {{ color:{colors[0]}; background: transparent; border: none; }}")
    lay = QHBoxLayout(f)
    lay.setContentsMargins(16, 12, 16, 12)
    lay.addWidget(label(text, wrap=True), 1)
    return f


def _clear_page(area: QScrollArea, body: QWidget) -> None:
    area.setObjectName("Page")
    area.setWidgetResizable(True)
    area.setFrameShape(QFrame.NoFrame)
    area.setWidget(body)


# ====================================================================== trang chủ


class HomePage(QScrollArea):
    def __init__(self, shell: Shell):
        super().__init__()
        self.shell = shell
        acc = shell.account
        if acc.is_admin:
            try:
                nhap_de.auto_import()           # bộ đề đặt trong thư mục nhap_de/ cạnh app
            except Exception:  # không để lỗi nhập đề chặn Trang chủ
                pass
        custom, errors = load_custom_exams()
        self.exams = merge_exams(ALL_EXAMS, custom)
        history = core.load_history(acc.username)

        body, lay = page_body()
        lay.addWidget(self._hero(acc, history))
        if errors and acc.is_admin:
            lay.addWidget(_banner(tr("Có đề tự soạn bị lỗi nên chưa được nạp: {err}{more}. Vào Đề thi để sửa.")
                                  .format(err=errors[0], more=tr(" (và {n} lỗi khác)").format(n=len(errors) - 1)
                                          if len(errors) > 1 else ""), tone="warn"))

        # --- số liệu
        stats = QHBoxLayout()
        stats.setSpacing(16)
        passed = sum(h["passed"] for h in history)
        best = max((h["score"] for h in history), default=None)
        stats.addWidget(T.stat_card(tr("Số lần làm bài"), str(len(history)), tr("tính cả luyện tập và thi thử"),
                                    icon="📝"))
        stats.addWidget(T.stat_card(tr("Điểm cao nhất"), f"{best}" if best is not None else "—",
                                    tr("thang điểm 1000"), icon="🏔"))
        stats.addWidget(T.stat_card(tr("Tỉ lệ đạt"), f"{round(100 * passed / len(history))}%" if history else "—",
                                    tr("điểm đạt từ 700"), icon="🎯"))
        lay.addLayout(stats)

        # --- huy hiệu
        head = QHBoxLayout()
        head.addWidget(label(tr("Huy hiệu"), "h2"))
        all_badges = gamify.badges(history)
        head.addWidget(label(tr("{n}/{total} đã mở khóa").format(n=sum(b.earned for b in all_badges),
                                                                 total=len(all_badges)), "muted"), 0, Qt.AlignBottom)
        head.addStretch()
        lay.addLayout(head)
        grid = QGridLayout()
        grid.setSpacing(10)
        for i, b in enumerate(all_badges):
            grid.addWidget(T.badge_tile(b.icon, pick(b.name_vi, b.name_en), pick(b.desc_vi, b.desc_en), b.earned),
                           i // 5, i % 5)
        for c in range(5):
            grid.setColumnStretch(c, 1)
        lay.addLayout(grid)

        # --- chọn môn
        self.history = history
        self.by_code = {code: [e for e in self.exams if e.code == code] for code in T.BRAND}
        self.code = next((c for c in T.BRAND if self.by_code[c]), "WORD")
        self.exam_idx = {code: 0 for code in T.BRAND}
        self.chapter_no = {code: None for code in T.BRAND}
        lay.addWidget(label(tr("Chọn môn"), "h2"))
        subjects = QHBoxLayout()
        subjects.setSpacing(16)
        self.subject_cards = {}
        for code in T.BRAND:
            card = self._subject_card(code)
            card.clicked.connect(lambda c=code: self.select_subject(c))
            subjects.addWidget(card)
            self.subject_cards[code] = card
        lay.addLayout(subjects)

        # --- chế độ
        lay.addWidget(label(tr("Chế độ"), "h2"))
        modes = QHBoxLayout()
        modes.setSpacing(16)
        self.mode_cards = {}
        for key, icon, title, desc in (
                ("training", "🌱", tr("Luyện tập"), tr("Không giới hạn giờ · có gợi ý · kiểm tra từng dự án")),
                ("chapter", "📚", tr("Luyện theo chương"), tr("Chọn một nhóm kỹ năng · gom câu từ mọi đề · có gợi ý")),
                ("testing", "⏳", tr("Thi thử"), tr("Tính giờ như thi thật · không gợi ý · chấm khi nộp bài · "
                                                   "XP ×1.2"))):
            card = ClickableCard(padding=18, spacing=14, horizontal=True)
            dot = QLabel()
            dot.setFixedSize(20, 20)
            card.lay.addWidget(dot)
            ic = QLabel(icon)
            ic.setStyleSheet("font-size: 18pt; background: transparent;")
            card.lay.addWidget(ic)
            col = QVBoxLayout()
            col.setSpacing(2)
            col.addWidget(label(title, "h3"))
            col.addWidget(label(desc, "caption", wrap=True))
            card.lay.addLayout(col, 1)
            card.clicked.connect(lambda k=key: self.select_mode(k))
            modes.addWidget(card)
            self.mode_cards[key] = (card, dot)
        lay.addLayout(modes)

        # --- chọn đề / chọn chương (đổi theo môn và chế độ)
        self.picker_title = label("", "h2")
        lay.addWidget(self.picker_title)
        self.picker = QWidget()
        self.picker_lay = QVBoxLayout(self.picker)
        self.picker_lay.setContentsMargins(0, 0, 0, 0)
        lay.addWidget(self.picker)

        foot = QHBoxLayout()
        if history:
            h = history[-1]
            foot.addWidget(label(tr("Lần gần nhất: {exam} — {score}/1000 ({verdict}) · {time}").format(
                exam=history_exam_name(h), score=h["score"], verdict=tr("Đạt") if h["passed"] else tr("Chưa đạt"),
                time=h["time"]), "muted"))
        foot.addStretch()
        self.start_btn = button(tr("Bắt đầu làm bài  →"), self.start, "primary", "lg")
        foot.addWidget(self.start_btn)
        lay.addLayout(foot)
        lay.addStretch()

        _clear_page(self, body)
        self.mode = "training"
        self.select_mode("training")
        self.select_subject(self.code)

    # ------------------------------------------------------------ khối chào
    def _hero(self, acc, history) -> QFrame:
        hero = Card(padding=28, spacing=26, horizontal=True, name="Hero")
        T.shadow(hero, blur=36, y=10, alpha=40)
        xp = gamify.total_xp(history)
        lv = gamify.level(xp)
        hero.lay.addWidget(T.LevelBadge(lv.number, lv.ratio, 104), 0, Qt.AlignVCenter)

        col = QVBoxLayout()
        col.setSpacing(6)
        first = acc.ho_ten.split()[-1] if acc.ho_ten.split() else acc.username
        col.addWidget(label(tr("Xin chào, {name}!").format(name=first), "display"))
        col.addWidget(label(tr("Cấp {n} · {title} — còn {xp} XP nữa để lên cấp").format(
            n=lv.number, title=pick(lv.name_vi, lv.name_en), xp=lv.xp_for_next - lv.xp_in_level), "caption"))
        col.addWidget(T.progress(lv.xp_in_level, lv.xp_for_next, "xp"))
        col.addSpacing(6)
        chips = QHBoxLayout()
        chips.setSpacing(10)
        days = gamify.streak(history)
        for icon, text in (("✨", tr("{xp} XP").format(xp=xp)),
                           ("🔥", tr("Chuỗi {n} ngày").format(n=days) if days else tr("Bắt đầu chuỗi hôm nay")),
                           ("🏅", tr("{n} lần đạt").format(n=sum(h["passed"] for h in history)))):
            f = QFrame()
            f.setObjectName("HeroChip")
            fl = QHBoxLayout(f)
            fl.setContentsMargins(12, 6, 12, 6)
            fl.addWidget(label(f"{icon}  {text}"))
            chips.addWidget(f)
        chips.addStretch()
        col.addLayout(chips)
        hero.lay.addLayout(col, 1)

        tip_vi, tip_en = gamify.tip_of_day()
        tip = QFrame()
        tip.setObjectName("HeroChip")
        tip.setFixedWidth(300)
        tl = QVBoxLayout(tip)
        tl.setContentsMargins(16, 14, 16, 14)
        tl.setSpacing(6)
        tl.addWidget(label("💡  " + tr("MẸO HÔM NAY"), "caption"))
        tl.addWidget(label(pick(tip_vi, tip_en), wrap=True))
        tl.addStretch()
        hero.lay.addWidget(tip)
        return hero

    def _best(self, exam) -> int | None:
        return max((h["score"] for h in self.history if h["exam"] == exam.name), default=None)

    def _subject_card(self, code) -> ClickableCard:
        _, color, tint, short = T.brand(code)
        exams = self.by_code[code]
        card = ClickableCard(padding=20, spacing=6)
        top = QHBoxLayout()
        top.addWidget(T.badge(code, 50))
        top.addStretch()
        card.check = chip("✓ " + tr("Đã chọn"), T.FOREST, PRIMARY_SOFT)
        top.addWidget(card.check, 0, Qt.AlignTop)
        card.lay.addLayout(top)
        card.lay.addSpacing(6)
        card.lay.addWidget(label(short, "h2"))
        exam_code = {"WORD": "MO-100 / MO-110", "EXCEL": "MO-200", "POWERPOINT": "MO-300"}[code]
        card.lay.addWidget(label(exam_code, "caption"))
        n_tasks = sum(len(p.tasks) for e in exams for p in e.projects)
        card.lay.addWidget(label(tr("{n} đề · {t} nhiệm vụ").format(n=len(exams), t=n_tasks), "muted"))
        bests = [b for b in (self._best(e) for e in exams) if b is not None]
        done = sum(1 for b in bests if b >= core.PASS_SCORE)
        card.lay.addSpacing(6)
        card.lay.addWidget(T.progress(done, max(len(exams), 1), None))
        card.lay.addWidget(label(tr("Đã đạt {n}/{total} đề").format(n=done, total=len(exams)), "caption"))
        return card

    def _exam_row(self, exam, on) -> ClickableCard:
        card = ClickableCard(padding=14, spacing=12, horizontal=True)
        dot = QLabel()
        dot.setFixedSize(18, 18)
        dot.setStyleSheet(f"border-radius: 9px; border: {'5px' if on else '2px'} solid "
                          f"{PRIMARY if on else '#D5D2C4'}; background: white;")
        card.lay.addWidget(dot)
        col = QVBoxLayout()
        col.setSpacing(1)
        name, _, sub = exam_name(exam).replace("Microsoft ", "").partition(" (")
        col.addWidget(label(name, "h3"))
        n_tasks = sum(len(p.tasks) for p in exam.projects)
        col.addWidget(label(f"{sub.rstrip(')') or tr('Đề tự soạn')} · " + tr("{p} dự án · {t} nhiệm vụ · {m} phút")
                            .format(p=len(exam.projects), t=n_tasks, m=exam.minutes), "caption"))
        card.lay.addLayout(col, 1)
        best = self._best(exam)
        right = QVBoxLayout()
        right.setSpacing(0)
        right.addWidget(T.stars(gamify.stars(best), 12), 0, Qt.AlignRight)
        right.addWidget(label(f"{best}/1000" if best is not None else tr("Chưa làm"), "caption"), 0, Qt.AlignRight)
        card.lay.addLayout(right)
        if on:
            card.setStyleSheet(f"QFrame#Card {{ background: {PRIMARY_SOFT}; border: 2px solid {PRIMARY}; "
                               "border-radius: 16px; }")
        return card

    def _chapter_card(self, ch, count, on) -> ClickableCard:
        card = ClickableCard(padding=16, spacing=4)
        top = QHBoxLayout()
        ic = QLabel(ch.icon)
        ic.setStyleSheet(f"font-size: 16pt; background: transparent; color: {T.FOREST};")
        top.addWidget(ic)
        top.addStretch()
        top.addWidget(chip(tr("Chương {n}").format(n=ch.number), T.FOREST, PRIMARY_SOFT))
        card.lay.addLayout(top)
        card.lay.addWidget(label(pick(ch.name_vi, ch.name_en), "h3", wrap=True))
        card.lay.addWidget(label(tr("{n} nhiệm vụ để luyện").format(n=count), "caption"))
        if on:
            card.setStyleSheet(f"QFrame#Card {{ background: {PRIMARY_SOFT}; border: 2px solid {PRIMARY}; "
                               "border-radius: 16px; }")
        elif not count:
            card.setEnabled(False)
        return card

    def _refresh_picker(self):
        while self.picker_lay.count():
            item = self.picker_lay.takeAt(0)
            if item.widget():
                item.widget().deleteLater()
            elif item.layout():
                QWidget().setLayout(item.layout())
        grid = QGridLayout()
        grid.setSpacing(12)
        short = T.brand(self.code)[3]
        if self.mode == "chapter":
            self.picker_title.setText(tr("Chọn chương – {subject}").format(subject=short))
            counts = chuong.count_tasks(self.exams, self.code)
            chs = chuong.chapters(self.code)
            if self.chapter_no[self.code] is None:
                self.chapter_no[self.code] = next((c.number for c in chs if counts.get(c.number)), None)
            for i, ch in enumerate(chs):
                card = self._chapter_card(ch, counts.get(ch.number, 0), ch.number == self.chapter_no[self.code])
                card.clicked.connect(lambda n=ch.number: self.select_chapter(n))
                grid.addWidget(card, i // 3, i % 3)
            cols = 3
            ready = bool(counts.get(self.chapter_no[self.code] or 0))
        else:
            self.picker_title.setText(tr("Chọn đề – {subject}").format(subject=short))
            exams = self.by_code[self.code]
            for i, exam in enumerate(exams):
                card = self._exam_row(exam, i == self.exam_idx[self.code])
                card.clicked.connect(lambda i=i: self.select_exam(i))
                grid.addWidget(card, i // 2, i % 2)
            cols = 2
            ready = bool(exams)
            if not exams:
                self.picker_lay.addWidget(T.empty_state(tr("Môn này chưa có đề nào.")))
        for c in range(cols):
            grid.setColumnStretch(c, 1)
        self.picker_lay.addLayout(grid)
        self.start_btn.setEnabled(ready)

    def select_subject(self, code):
        self.code = code
        for c, card in self.subject_cards.items():
            on = c == code
            card.setStyleSheet(f"QFrame#Card {{ background: white; border: 2px solid {T.FOREST}; "
                               "border-radius: 16px; }" if on else "")
            card.check.setVisible(on)
        self._refresh_picker()

    def select_exam(self, idx):
        self.exam_idx[self.code] = idx
        self._refresh_picker()

    def select_chapter(self, number):
        self.chapter_no[self.code] = number
        self._refresh_picker()

    def select_mode(self, key):
        self.mode = key
        for k, (card, dot) in self.mode_cards.items():
            on = k == key
            card.setStyleSheet(f"QFrame#Card {{ background: {PRIMARY_SOFT}; border: 2px solid {PRIMARY}; "
                               "border-radius: 16px; }" if on else "")
            dot.setStyleSheet(f"border-radius: 10px; border: {'6px' if on else '2px'} solid "
                              f"{PRIMARY if on else '#D5D2C4'}; background: white;")
        if hasattr(self, "subject_cards"):
            self._refresh_picker()

    def start(self):
        if self.mode == "chapter":
            exam = chuong.practice_exam(self.exams, self.code, self.chapter_no[self.code] or 0)
            if exam is None:
                T.info(self, tr("Chưa có nhiệm vụ"), tr("Chương này chưa có nhiệm vụ nào để luyện."))
                return
            self.shell.start_exam(exam, "chapter")
        else:
            exams = self.by_code[self.code]
            if exams:
                self.shell.start_exam(exams[self.exam_idx[self.code]], self.mode)


# ====================================================================== lịch sử


class HistoryPage(QScrollArea):
    def __init__(self, shell: Shell):
        super().__init__()
        history = core.load_history(shell.account.username)
        body, lay = page_body()
        lay.addWidget(T.page_header(tr("Lịch sử của tôi"), tr("Tất cả các lần bạn đã nộp bài.")))
        stats = QHBoxLayout()
        stats.setSpacing(16)
        passed = sum(h["passed"] for h in history)
        stats.addWidget(T.stat_card(tr("Số lần làm bài"), str(len(history)), icon="📝"))
        stats.addWidget(T.stat_card(tr("Điểm cao nhất"), str(max((h["score"] for h in history), default="—")),
                                    icon="🏔"))
        stats.addWidget(T.stat_card(tr("Điểm trung bình"),
                                    str(round(sum(h["score"] for h in history) / len(history))) if history else "—",
                                    icon="📊"))
        stats.addWidget(T.stat_card(tr("Tổng XP"), str(gamify.total_xp(history)), icon="✨"))
        stats.addWidget(T.stat_card(tr("Tỉ lệ đạt"), f"{round(100 * passed / len(history))}%" if history else "—",
                                    icon="🎯"))
        lay.addLayout(stats)
        if history:
            t = T.table([(tr("Thời gian"), 150), (tr("Bài thi"), None), (tr("Chế độ"), 110), (tr("Làm trong"), 100),
                         (tr("Điểm"), 80), ("XP", 70), (tr("Kết quả"), 110)])
            rows = [[h["time"], history_exam_name(h), mode_name(h["mode"]), T.fmt_time(h.get("duration")),
                     f'{h["score"]}', f"+{gamify.attempt_xp(h)}", tr("Đạt") if h["passed"] else tr("Chưa đạt")]
                    for h in reversed(history)]
            T.set_rows(t, rows, ["ok" if h["passed"] else "bad" for h in reversed(history)])
            t.setMinimumHeight(360)
            lay.addWidget(t, 1)
        else:
            lay.addWidget(T.empty_state(tr("Bạn chưa nộp bài lần nào. Vào Trang chủ để bắt đầu."), "🌱"))
            lay.addStretch()
        _clear_page(self, body)


# ====================================================================== kết quả


class ResultPage(QScrollArea):
    def __init__(self, shell: Shell, session, report):
        super().__init__()
        self.shell, self.session, self.results = shell, session, report["results"]
        passed = report["passed"]
        body, lay = page_body()

        # --- thẻ điểm
        head = Card(padding=28, spacing=28, raised=True, horizontal=True)
        head.lay.addWidget(T.ScoreRing(report["score"], passed))
        info = QVBoxLayout()
        info.setSpacing(6)
        chips = QHBoxLayout()
        chips.addWidget(chip(tr("ĐẠT") if passed else tr("CHƯA ĐẠT"), SUCCESS if passed else DANGER,
                             SUCCESS_SOFT if passed else DANGER_SOFT))
        if report.get("xp") is not None:
            chips.addWidget(chip(f"✨ +{report['xp']} XP", "#7A5A12", GOLD_SOFT))
        chips.addStretch()
        info.addLayout(chips)
        info.addWidget(label(exam_name(session.exam), "h1"))
        if passed:
            msg = tr("Tuyệt vời! Bạn đã vượt qua điểm đạt. 🎉") if report["score"] < 1000 else \
                tr("Hoàn hảo! 1000/1000 – không sai câu nào. 🏆")
        else:
            msg = tr("Cần thêm {n} điểm để đạt. Xem lại các nhiệm vụ sai bên dưới – lần sau chắc chắn tốt hơn! 💪")\
                .format(n=core.PASS_SCORE - report["score"])
        info.addWidget(label(msg, "muted", wrap=True))
        info.addSpacing(8)
        boxes = QHBoxLayout()
        boxes.setSpacing(10)
        for title, value in ((tr("Đúng"), f'{report["correct"]}/{report["total"]}'),
                             (tr("Thời gian"), T.fmt_time(report.get("duration", 0))),
                             (tr("Điểm đạt"), str(core.PASS_SCORE)),
                             (tr("Chế độ"), mode_name(session.mode))):
            b = QFrame()
            b.setObjectName("Soft")
            bl = QVBoxLayout(b)
            bl.setContentsMargins(14, 10, 14, 10)
            bl.setSpacing(0)
            bl.addWidget(label(title, "caption"))
            v = label(value)
            v.setStyleSheet(f"font-size: 13pt; font-weight: 700; color: {T.FOREST};")
            bl.addWidget(v)
            boxes.addWidget(b)
        boxes.addStretch()
        info.addLayout(boxes)
        head.lay.addLayout(info, 1)
        lay.addWidget(head)

        for key in report.get("new_badges", []):
            b = next((x for x in gamify.badges([]) if x.key == key), None)
            if b:
                lay.addWidget(_banner(tr("{icon}  Huy hiệu mới: <b>{name}</b> — {desc}").format(
                    icon=b.icon, name=pick(b.name_vi, b.name_en), desc=pick(b.desc_vi, b.desc_en)), "gold"))

        # --- theo dự án
        per = QGridLayout()
        per.setSpacing(16)
        for i, project in enumerate(session.exam.projects):
            rs = [r for r in self.results if r.project == project.name and r.correct is not None]
            ok = sum(r.correct is True for r in rs)
            c = Card(padding=16, spacing=8)
            row = QHBoxLayout()
            row.addWidget(label(pick(project.name, project.name_en), "h3"), 1)
            row.addWidget(label(f"{ok}/{len(rs)}", "muted"))
            c.lay.addLayout(row)
            ratio = ok / len(rs) if rs else 0
            c.lay.addWidget(T.progress(int(ratio * 100), 100, None if ratio >= 0.7 else "warn"))
            per.addWidget(c, i // 4, i % 4)
        lay.addLayout(per)

        # --- chi tiết
        lay.addWidget(label(tr("Chi tiết từng nhiệm vụ"), "h2"))
        self.table = T.table([(tr("Kết quả"), 130), (tr("Dự án"), 200), (tr("Nhiệm vụ"), None)])
        verdict = {True: "✓  " + tr("Đúng"), False: "✗  " + tr("Sai"), None: "–  " + tr("Tự kiểm tra")}
        rows = [[verdict[r.correct], pick(r.project, r.project_en).split("–")[-1].strip(), pick(r.task, r.task_en)]
                for r in self.results]
        tones = [{True: "ok", False: "bad", None: "muted"}[r.correct] for r in self.results]
        T.set_rows(self.table, rows, tones, tone_cols={0})
        self.table.setMinimumHeight(min(60 + 42 * len(rows), 440))
        self.table.itemSelectionChanged.connect(self.on_select)
        lay.addWidget(self.table)
        self.hint = _banner(tr("Bấm vào một nhiệm vụ ở bảng trên để xem cách làm."), "warn")
        self.hint_label = self.hint.findChild(QLabel)
        lay.addWidget(self.hint)

        btns = QHBoxLayout()
        btns.addWidget(button(tr("Mở thư mục bài làm"), lambda: core.open_in_office(session.workdir)))
        btns.addStretch()
        btns.addWidget(button(tr("Làm lại đề này"), lambda: shell.start_exam(session.exam, session.mode)))
        btns.addWidget(button(tr("Về trang chủ"), lambda: shell.go("home"), "primary"))
        lay.addLayout(btns)

        _clear_page(self, body)
        if passed:
            QTimer.singleShot(250, lambda: T.Confetti(self.viewport()))

    def on_select(self):
        rows = self.table.selectionModel().selectedRows()
        if rows:
            r = self.results[rows[0].row()]
            text = f"<b>{tr('Cách làm:')}</b> {esc(pick(r.hint, r.hint_en))}"
            if r.error:
                text += "<br><i>" + tr("(Lỗi khi đọc file: {err})").format(err=esc(r.error)) + "</i>"
            self.hint_label.setText(text)


# ====================================================================== thanh làm bài


class TaskCard(QFrame):
    def __init__(self, bar: "ExamBar", key, number: int, task):
        super().__init__()
        self.bar, self.key, self.task = bar, key, task
        self.setObjectName("Card")
        outer = QVBoxLayout(self)
        outer.setContentsMargins(14, 10, 14, 10)
        outer.setSpacing(8)
        row = QHBoxLayout()
        row.setSpacing(12)
        self.num = QLabel(str(number))
        self.num.setFixedSize(28, 28)
        self.num.setAlignment(Qt.AlignCenter)
        row.addWidget(self.num, 0, Qt.AlignVCenter)
        text = label(pick(task.title, task.title_en), wrap=True)
        text.setStyleSheet("font-size: 10.5pt;")
        text.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Preferred)
        text.setTextInteractionFlags(Qt.TextSelectableByMouse)
        row.addWidget(text, 1)
        self.status = chip("", SUCCESS, SUCCESS_SOFT)
        self.status.hide()
        row.addWidget(self.status, 0, Qt.AlignVCenter)
        if bar.training:
            self.hint_btn = button("💡 " + tr("Gợi ý"), self.toggle_hint, "ghost", "sm")
            row.addWidget(self.hint_btn, 0, Qt.AlignVCenter)
        self.mark = button("⚑ " + tr("Đánh dấu"), None, None, "sm")
        self.mark.setCheckable(True)
        self.mark.setProperty("toggle", "mark")
        self.mark.toggled.connect(self._on_mark)
        row.addWidget(self.mark, 0, Qt.AlignVCenter)
        self.done = button("✓ " + tr("Đã làm"), None, None, "sm")
        self.done.setCheckable(True)
        self.done.setProperty("toggle", "done")
        self.done.toggled.connect(self._on_done)
        row.addWidget(self.done, 0, Qt.AlignVCenter)
        outer.addLayout(row)
        self.hint = label(f"<b>{tr('Cách làm:')}</b> {esc(pick(task.hint, task.hint_en))}", wrap=True)
        self.hint.setStyleSheet(f"background:{GOLD_SOFT}; color:#5C4410; border-radius:10px; padding:8px 12px;")
        self.hint.hide()
        outer.addWidget(self.hint)
        self.mark.setChecked(key in bar.s.marked)
        self.done.setChecked(key in bar.s.done)
        self._paint_num()
        if key in bar.checked:
            self.set_status(bar.checked[key])
        if key in bar.hints_open and bar.training:
            self.toggle_hint()

    def _paint_num(self):
        c = self.bar.color
        on = self.done.isChecked()
        self.num.setStyleSheet(f"border-radius:14px; border:2px solid {c}; font-weight:800;"
                               f"background:{c if on else 'white'}; color:{'white' if on else c};")

    def _on_mark(self, on):
        (self.bar.s.marked.add if on else self.bar.s.marked.discard)(self.key)
        self.bar.update_progress()

    def _on_done(self, on):
        (self.bar.s.done.add if on else self.bar.s.done.discard)(self.key)
        self._paint_num()
        self.bar.update_progress()

    def toggle_hint(self):
        self.hint.setVisible(not self.hint.isVisible())
        (self.bar.hints_open.add if self.hint.isVisible() else self.bar.hints_open.discard)(self.key)
        self.hint_btn.setText(("🙈 " + tr("Ẩn gợi ý")) if self.hint.isVisible() else ("💡 " + tr("Gợi ý")))

    def set_status(self, ok: bool | None):
        if ok is None:
            T.set_chip(self.status, tr("Tự kiểm tra"), WARN, WARN_SOFT)
        else:
            T.set_chip(self.status, "✓ " + tr("Đúng") if ok else "✗ " + tr("Sai"), SUCCESS if ok else DANGER,
                       SUCCESS_SOFT if ok else DANGER_SOFT)
        self.status.show()


class ExamBar(QWidget):
    HEIGHT = 390

    def __init__(self, session: core.Session, on_finish, on_quit):
        super().__init__(None, Qt.Window | Qt.WindowStaysOnTopHint)
        self.s, self.on_finish, self.on_quit = session, on_finish, on_quit
        self.training = session.mode in ("training", "chapter")
        self.idx = 0
        self.started = time.time()
        self.seconds = 0 if self.training else session.exam.minutes * 60
        self.finished = False
        self.checked: dict = {}
        self.hints_open: set = set()
        _, self.color, self.tint, _ = T.brand(session.exam.code)
        self.setWindowIcon(app_icon())
        self.setObjectName("Page")
        self.setAttribute(Qt.WA_StyledBackground, True)
        geo = QGuiApplication.primaryScreen().availableGeometry()
        self.setGeometry(geo.x(), geo.y() + geo.height() - self.HEIGHT - 30, geo.width(), self.HEIGHT)
        outer = QVBoxLayout(self)
        outer.setContentsMargins(0, 0, 0, 0)
        self.root = None
        self.timer = QTimer(self)
        self.timer.timeout.connect(self.tick)
        self.timer.start(1000)
        self._build()
        QTimer.singleShot(0, self.open_file)

    def _build(self):
        """Dựng (lại) toàn bộ nội dung thanh – gọi lại khi đổi ngôn ngữ."""
        if self.root is not None:
            self.layout().removeWidget(self.root)
            self.root.deleteLater()
        self.root = QWidget()
        self.root.setObjectName("Page")
        self.layout().addWidget(self.root)
        session = self.s
        self.setWindowTitle(f"MOS – {exam_name(session.exam)}")
        lay = QVBoxLayout(self.root)
        lay.setContentsMargins(0, 0, 0, 0)
        lay.setSpacing(0)

        # --- đầu: môn, chế độ, tab dự án, ngôn ngữ, đồng hồ
        head = QWidget()
        head.setObjectName("BarHeader")
        head.setAttribute(Qt.WA_StyledBackground, True)
        hl = QHBoxLayout(head)
        hl.setContentsMargins(16, 10, 16, 10)
        hl.setSpacing(12)
        hl.addWidget(T.badge(session.exam.code, 30))
        name = label(exam_name(session.exam).replace("Microsoft ", ""))
        name.setStyleSheet(f"font-family: '{T.pick_heading_family()}'; font-weight: 800; font-size: 12pt;")
        hl.addWidget(name)
        mode_chip = {"training": tr("LUYỆN TẬP"), "chapter": tr("THEO CHƯƠNG")}.get(session.mode, tr("THI THỬ"))
        hl.addWidget(chip(mode_chip, T.TEXT,
                          T.OLIVE if self.training else "#E3C66B"))
        hl.addSpacing(14)
        self.tabs = QButtonGroup(self)
        for i, p in enumerate(session.exam.projects):
            short = pick(p.name, p.name_en).split("–")[-1].strip()
            if len(session.exam.projects) > 4 and len(short) > 14:
                short = short[:13] + "…"
            t = QPushButton(f"{i + 1}  {short}")
            t.setProperty("tab", True)
            t.setCheckable(True)
            t.setCursor(Qt.PointingHandCursor)
            t.clicked.connect(lambda _=False, k=i: self.go(k))
            self.tabs.addButton(t, i)
            hl.addWidget(t)
        hl.addStretch()
        hl.addWidget(T.lang_switch(i18n.get_lang(), self.change_language, dark=True))
        self.timer_lbl = QLabel()
        hl.addWidget(self.timer_lbl)
        lay.addWidget(head)

        # --- mô tả dự án
        intro = QWidget()
        intro.setAttribute(Qt.WA_StyledBackground, True)
        intro.setStyleSheet(f"background: {PRIMARY_SOFT};")
        il = QHBoxLayout(intro)
        il.setContentsMargins(18, 8, 18, 8)
        self.intro = label("", wrap=True)
        self.intro.setStyleSheet("background: transparent;")
        il.addWidget(self.intro, 1)
        self.file_lbl = label("")
        self.file_lbl.setStyleSheet(f"background: transparent; color:{self.color}; font-weight:700;")
        il.addWidget(self.file_lbl)
        lay.addWidget(intro)

        # --- danh sách nhiệm vụ
        self.area = QScrollArea()
        self.area.setWidgetResizable(True)
        self.area.setFrameShape(QFrame.NoFrame)
        self.area.setObjectName("Page")
        lay.addWidget(self.area, 1)

        # --- chân: nút thao tác
        foot = QWidget()
        foot.setObjectName("BarFooter")
        foot.setAttribute(Qt.WA_StyledBackground, True)
        fl = QHBoxLayout(foot)
        fl.setContentsMargins(16, 10, 16, 10)
        fl.setSpacing(8)
        fl.addWidget(button("📂 " + tr("Mở file"), self.open_file))
        fl.addWidget(button("↺ " + tr("Làm lại dự án"), self.reset_project))
        if self.training:
            fl.addWidget(button("✓ " + tr("Kiểm tra dự án"), self.check_current, "accent"))
        fl.addSpacing(12)
        self.progress_lbl = label("", "muted")
        fl.addWidget(self.progress_lbl)
        fl.addStretch()
        self.prev_btn = button("‹  " + tr("Dự án trước"), lambda: self.go(self.idx - 1))
        self.next_btn = button(tr("Dự án sau") + "  ›", lambda: self.go(self.idx + 1))
        fl.addWidget(self.prev_btn)
        fl.addWidget(self.next_btn)
        fl.addSpacing(8)
        fl.addWidget(button(tr("Nộp bài"), self.submit, "primary"))
        lay.addWidget(foot)

        self._render_timer()
        self.show_project()

    def change_language(self, code: str):
        i18n.set_lang(code)
        self._build()

    # ------------------------------------------------------------ hiển thị
    def show_project(self):
        project = self.s.exam.projects[self.idx]
        n = len(self.s.exam.projects)
        self.tabs.button(self.idx).setChecked(True)
        self.intro.setText(f"<b>{esc(pick(project.name, project.name_en))}</b> — "
                           f"{esc(pick(project.intro, project.intro_en))}  "
                           f"<i>{tr('Làm xong nhớ lưu file (Ctrl + S).')}</i>")
        self.file_lbl.setText(project.filename)
        self.prev_btn.setEnabled(self.idx > 0)
        self.next_btn.setEnabled(self.idx < n - 1)
        body = QWidget()
        body.setObjectName("Page")
        bl = QVBoxLayout(body)
        bl.setContentsMargins(16, 12, 16, 12)
        bl.setSpacing(8)
        self.cards = [TaskCard(self, (self.idx, i), i + 1, t) for i, t in enumerate(project.tasks)]
        for c in self.cards:
            bl.addWidget(c)
        bl.addStretch()
        old = self.area.takeWidget()
        if old:
            old.deleteLater()
        self.area.setWidget(body)
        self.update_progress()

    def update_progress(self):
        total = sum(len(p.tasks) for p in self.s.exam.projects)
        text = tr("Đã làm {n}/{total}").format(n=len(self.s.done), total=total)
        if self.s.marked:
            text += "  ·  " + tr("Đánh dấu {n}").format(n=len(self.s.marked))
        self.progress_lbl.setText(text)

    def _render_timer(self):
        urgent = not self.training and self.seconds < 300
        self.timer_lbl.setText(("⏱  " if self.training else "⏳  ") + T.fmt_time(self.seconds))
        self.timer_lbl.setStyleSheet(f"background:{DANGER if urgent else T.SIDEBAR_2}; color:white;"
                                     "border-radius:8px; padding:5px 12px; font-size:13pt; font-weight:800;"
                                     "font-family: Consolas, 'DejaVu Sans Mono', monospace;")

    def tick(self):
        if self.finished:
            return
        self.seconds += 1 if self.training else -1
        self._render_timer()
        if not self.training and self.seconds <= 0:
            self.timer.stop()
            T.warn(self, tr("Hết giờ"), tr("Đã hết thời gian. Bài sẽ được nộp tự động.\n"
                                           "(Các thay đổi chưa lưu sẽ không được tính.)"))
            self.finish()

    # ------------------------------------------------------------ thao tác
    def open_file(self):
        path = self.s.file_of(self.idx)
        try:
            core.open_in_office(path)
        except OSError as exc:
            T.error(self, tr("Không mở được file"), f"{exc}\n\n" + tr("Hãy tự mở file:") + f"\n{path}")

    def go(self, idx):
        if 0 <= idx < len(self.s.exam.projects) and idx != self.idx:
            self.idx = idx
            self.show_project()
            self.open_file()
        else:
            self.tabs.button(self.idx).setChecked(True)

    def reset_project(self):
        path = self.s.file_of(self.idx)
        if core.file_is_open(path):
            T.warn(self, tr("File đang mở"), tr("Hãy ĐÓNG file trong Office trước khi làm lại dự án."))
            return
        if T.confirm(self, tr("Làm lại dự án"), tr("Tạo lại file gốc cho dự án này?\n"
                                                   "(Bài cũ được lưu với tên *_cu)")):
            core.reset_project(self.s, self.idx)
            for t in range(len(self.s.exam.projects[self.idx].tasks)):
                self.checked.pop((self.idx, t), None)
            self.show_project()
            self.open_file()

    def _confirm_saved(self, path) -> bool:
        if core.file_is_open(path):
            return T.confirm(self, tr("File đang mở"), tr("File vẫn đang mở trong Office. Chỉ những gì đã LƯU mới "
                                                          "được chấm.\n\nBạn đã bấm Ctrl + S chưa? Chọn Đồng ý để "
                                                          "chấm."))
        return True

    def check_current(self):
        if not self._confirm_saved(self.s.file_of(self.idx)):
            return
        results = core.check_project(self.s, self.idx)
        for card, r in zip(self.cards, results):
            self.checked[card.key] = r.correct
            card.set_status(r.correct)
        errs = [r.error for r in results if r.error]
        if errs:
            T.warn(self, tr("Không đọc được file"), tr("Lỗi: {err}\n\nHãy lưu lại file rồi thử lại.")
                   .format(err=errs[0]))
        elif results and all(r.correct is not False for r in results):
            T.Confetti(self.area.viewport(), count=80, seconds=2.5)

    def submit(self):
        opened = [i for i in range(len(self.s.exam.projects)) if core.file_is_open(self.s.file_of(i))]
        msg = tr("Nộp bài và chấm điểm?")
        if opened:
            msg = tr("Vẫn còn file đang mở trong Office. Chỉ phần đã LƯU mới được chấm.") + "\n\n" + msg
        if self.s.marked:
            msg = tr("Bạn còn {n} nhiệm vụ đánh dấu xem lại.").format(n=len(self.s.marked)) + "\n" + msg
        if T.confirm(self, tr("Nộp bài"), msg):
            self.finish()

    def finish(self):
        self.finished = True
        self.timer.stop()
        report = core.grade(self.s)
        report["duration"] = int(time.time() - self.started)
        before = core.load_history(self.s.user)
        try:
            core.save_history(report)
        except OSError:
            pass
        entry = {k: v for k, v in report.items() if k != "results"}
        report["xp"] = gamify.attempt_xp(entry)
        report["new_badges"] = [b.key for b in gamify.new_badges(before, before + [entry])]
        self.close()
        self.on_finish(self.s, report)

    def closeEvent(self, e):
        if self.finished:
            return super().closeEvent(e)
        if T.confirm(self, tr("Thoát bài thi"), tr("Thoát bài thi? Kết quả sẽ không được chấm.")):
            self.finished = True
            self.timer.stop()
            super().closeEvent(e)
            self.on_quit()
        else:
            e.ignore()


# ====================================================================== khởi động


def main() -> None:
    if hasattr(Qt, "HighDpiScaleFactorRoundingPolicy"):
        QGuiApplication.setHighDpiScaleFactorRoundingPolicy(Qt.HighDpiScaleFactorRoundingPolicy.PassThrough)
    QApplication.setAttribute(Qt.AA_ShareOpenGLContexts)     # cần cho trình duyệt nhúng (bài SCORM)
    app = QApplication.instance() or QApplication(sys.argv)
    app.setApplicationName("Luyện thi MOS")
    app.setStyle("Fusion")
    T.load_fonts()
    family = T.pick_font_family()
    app.setFont(QFont(family, 10))
    app.setStyleSheet(T.stylesheet(family, T.pick_heading_family()))
    win = MainWindow(AccountStore())
    win.show()
    sys.exit(app.exec())
