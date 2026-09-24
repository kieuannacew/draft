"""Ứng dụng Luyện thi MOS – giao diện Qt (PySide6).

Cấu trúc màn hình:
  MainWindow
   ├─ LoginPage                 đăng nhập
   └─ Shell                     thanh bên (menu) + trang nội dung
        ├─ HomePage             chọn bài thi / chế độ → Bắt đầu
        ├─ HistoryPage          lịch sử của tôi
        ├─ ResultPage           kết quả sau khi nộp bài
        └─ (quản trị) AccountsPage / ExamsPage / ResultsPage   – xem admin.py
  ExamBar                       cửa sổ làm bài nằm ở cạnh dưới màn hình, luôn nổi trên Office
"""
from __future__ import annotations

import sys
import time

from PySide6.QtCore import Qt, QTimer
from PySide6.QtGui import QFont, QGuiApplication, QIcon, QPixmap, QPainter, QColor
from PySide6.QtWidgets import (QApplication, QButtonGroup, QDialog, QFrame, QGridLayout, QHBoxLayout, QLabel,
                               QLineEdit, QMainWindow, QPushButton, QScrollArea, QSizePolicy, QStackedWidget,
                               QVBoxLayout, QWidget)

from .. import core
from ..accounts import ROLE_NAMES, AccountError, AccountStore
from ..custom import load_custom_exams, merge_exams
from ..exams import ALL_EXAMS
from . import theme as T
from .theme import (DANGER, DANGER_SOFT, PRIMARY, PRIMARY_SOFT, SUCCESS, SUCCESS_SOFT, WARN,
                    WARN_SOFT, Card, ClickableCard, button, chip, label)


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
        outer.setAlignment(Qt.AlignCenter)

        card = Card(padding=36, spacing=0, raised=True)
        card.setFixedWidth(420)
        outer.addWidget(card, 0, Qt.AlignCenter)
        lay = card.lay

        badges = QHBoxLayout()
        badges.setSpacing(8)
        for code in ("WORD", "EXCEL", "POWERPOINT"):
            badges.addWidget(T.badge(code, 36))
        badges.addStretch()
        lay.addLayout(badges)
        lay.addSpacing(20)
        lay.addWidget(label("Luyện thi MOS", "h1"))
        lay.addWidget(label("Đăng nhập bằng tài khoản giáo viên cấp cho bạn.", "muted", wrap=True))
        lay.addSpacing(24)

        lay.addWidget(label("Tên đăng nhập", "h3"))
        lay.addSpacing(6)
        self.user = QLineEdit()
        self.user.setProperty("size", "lg")
        self.user.setPlaceholderText("vd: nguyenvana")
        lay.addWidget(self.user)
        lay.addSpacing(14)
        lay.addWidget(label("Mật khẩu", "h3"))
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
        lay.addWidget(button("Đăng nhập", self.login, "primary", "lg"))

        store = main.store
        admin = store.get("admin")
        if admin and admin.doi_mat_khau and len(store.accounts) == 1:
            lay.addSpacing(16)
            tip = QFrame()
            tip.setObjectName("Banner")
            tl = QVBoxLayout(tip)
            tl.setContentsMargins(14, 10, 14, 10)
            tl.addWidget(label("<b>Lần đầu sử dụng?</b><br>Đăng nhập <b>admin</b> / <b>admin</b> "
                               "rồi đặt mật khẩu mới.", wrap=True))
            lay.addWidget(tip)

        foot = label("Chấm điểm tự động • Word • Excel • PowerPoint", "caption")
        foot.setStyleSheet("color: rgba(255,255,255,0.6);")
        outer.addSpacing(18)
        outer.addWidget(foot, 0, Qt.AlignCenter)

        self.user.returnPressed.connect(self.pw.setFocus)
        self.pw.returnPressed.connect(self.login)

    def showEvent(self, e):
        super().showEvent(e)
        self.user.setFocus()

    def login(self):
        try:
            acc = self.main.store.authenticate(self.user.text(), self.pw.text())
        except AccountError as exc:
            self.err.setText(str(exc))
            self.err.show()
            return
        self.err.hide()
        self.pw.clear()
        self.main.enter(acc)


class ChangePasswordDialog(QDialog):
    def __init__(self, parent, store: AccountStore, acc, forced=False):
        super().__init__(parent)
        self.store, self.acc, self.forced = store, acc, forced
        self.setWindowTitle("Đổi mật khẩu")
        self.setMinimumWidth(420)
        lay = QVBoxLayout(self)
        lay.setContentsMargins(28, 24, 28, 24)
        lay.setSpacing(8)
        lay.addWidget(label("Đổi mật khẩu", "h2"))
        if forced:
            lay.addWidget(label("Đây là lần đăng nhập đầu tiên — hãy đặt mật khẩu mới cho tài khoản "
                                f"<b>{acc.username}</b>.", "muted", wrap=True))
        lay.addSpacing(8)
        self.old = self._field(lay, "Mật khẩu hiện tại") if not forced else None
        self.new = self._field(lay, "Mật khẩu mới")
        self.again = self._field(lay, "Nhập lại mật khẩu mới")
        self.err = label("", wrap=True)
        self.err.setStyleSheet(f"color:{DANGER};")
        lay.addWidget(self.err)
        row = QHBoxLayout()
        row.addStretch()
        if not forced:
            row.addWidget(button("Hủy", self.reject))
        row.addWidget(button("Lưu mật khẩu", self.save, "primary"))
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
            self.store.set_password(self.acc.username, self.new.text())
        except AccountError as exc:
            self.err.setText(str(exc))
            return
        self.accept()


# ====================================================================== khung ứng dụng


class Shell(QWidget):
    """Thanh bên + vùng nội dung."""

    def __init__(self, main: "MainWindow", account):
        super().__init__()
        self.main, self.account = main, account
        lay = QHBoxLayout(self)
        lay.setContentsMargins(0, 0, 0, 0)
        lay.setSpacing(0)
        lay.addWidget(self._sidebar())
        self.content = QStackedWidget()
        lay.addWidget(self.content, 1)
        self.current = None
        self.go("home")

    # ------------------------------------------------------------ thanh bên
    def _sidebar(self) -> QWidget:
        side = QWidget()
        side.setObjectName("Sidebar")
        side.setAttribute(Qt.WA_StyledBackground, True)
        side.setFixedWidth(248)
        lay = QVBoxLayout(side)
        lay.setContentsMargins(16, 22, 16, 18)
        lay.setSpacing(4)

        brand_row = QHBoxLayout()
        brand_row.setSpacing(10)
        logo = QLabel()
        logo.setPixmap(app_icon().pixmap(34, 34))
        brand_row.addWidget(logo)
        names = QVBoxLayout()
        names.setSpacing(0)
        title = label("Luyện thi MOS")
        title.setStyleSheet("font-size: 12pt; font-weight: 800;")
        names.addWidget(title)
        names.addWidget(label("Word · Excel · PowerPoint", "caption"))
        brand_row.addLayout(names, 1)
        lay.addLayout(brand_row)
        lay.addSpacing(26)

        self.nav = QButtonGroup(self)
        self.nav.setExclusive(True)
        self.nav_buttons = {}

        def nav(key, text):
            b = QPushButton(text)
            b.setProperty("nav", True)
            b.setCheckable(True)
            b.setCursor(Qt.PointingHandCursor)
            b.clicked.connect(lambda: self.go(key))
            self.nav.addButton(b)
            self.nav_buttons[key] = b
            lay.addWidget(b)

        lay.addWidget(label("HỌC TẬP", "overline"))
        lay.addSpacing(4)
        nav("home", "Trang chủ")
        nav("history", "Lịch sử của tôi")
        if self.account.is_admin:
            lay.addSpacing(18)
            lay.addWidget(label("QUẢN TRỊ", "overline"))
            lay.addSpacing(4)
            nav("accounts", "Tài khoản")
            nav("exams", "Đề thi")
            nav("results", "Kết quả học viên")
        lay.addStretch()

        me = QFrame()
        me.setStyleSheet(f"QFrame {{ background: {T.SIDEBAR_2}; border-radius: 12px; }}")
        ml = QVBoxLayout(me)
        ml.setContentsMargins(12, 12, 12, 12)
        top = QHBoxLayout()
        top.addWidget(T.avatar(self.account.ho_ten, 36, PRIMARY if self.account.is_admin else "#475467"))
        names = QVBoxLayout()
        names.setSpacing(0)
        n = label(self.account.ho_ten)
        n.setStyleSheet("font-weight: 700;")
        names.addWidget(n)
        names.addWidget(label(f"{self.account.username} · {ROLE_NAMES[self.account.vai_tro]}", "caption"))
        top.addLayout(names, 1)
        ml.addLayout(top)
        row = QHBoxLayout()
        row.addWidget(button("Mật khẩu", self.change_password, "side"))
        row.addWidget(button("Đăng xuất", self.main.logout, "side"))
        ml.addLayout(row)
        lay.addWidget(me)
        return side

    # ------------------------------------------------------------ điều hướng
    def go(self, key: str, **kw) -> None:
        from . import admin
        factories = {
            "home": lambda: HomePage(self),
            "history": lambda: HistoryPage(self),
            "result": lambda: ResultPage(self, **kw),
            "accounts": lambda: admin.AccountsPage(self),
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
        btn = self.nav_buttons.get(key)
        if btn:
            btn.setChecked(True)
        else:
            self.nav.setExclusive(False)
            for b in self.nav_buttons.values():
                b.setChecked(False)
            self.nav.setExclusive(True)

    def change_password(self):
        ChangePasswordDialog(self, self.main.store, self.account).exec()

    # ------------------------------------------------------------ làm bài
    def start_exam(self, exam, mode) -> None:
        try:
            session = core.new_session(exam, mode, user=self.account.username)
        except Exception:
            T.error(self, "Lỗi", "Không tạo được file bài thi:\n" + core.format_exception())
            return
        self.main.hide()
        self.bar = ExamBar(session, on_finish=self.show_result, on_quit=self.main.show)
        self.bar.show()

    def show_result(self, session, report) -> None:
        self.main.show()
        self.main.raise_()
        self.main.activateWindow()
        self.go("result", session=session, report=report)


class MainWindow(QMainWindow):
    def __init__(self, store: AccountStore):
        super().__init__()
        self.store = store
        self.setWindowTitle("Luyện thi MOS")
        self.setWindowIcon(app_icon())
        self.resize(1240, 800)
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

    def logout(self):
        self._set(LoginPage(self))

    def enter(self, acc):
        if acc.doi_mat_khau:
            if ChangePasswordDialog(self, self.store, acc, forced=True).exec() != QDialog.Accepted:
                return
        acc = self.store.get(acc.username)
        self._set(Shell(self, acc))


def app_icon() -> QIcon:
    pix = QPixmap(64, 64)
    pix.fill(Qt.transparent)
    p = QPainter(pix)
    p.setRenderHint(QPainter.Antialiasing)
    p.setBrush(QColor(PRIMARY))
    p.setPen(Qt.NoPen)
    p.drawRoundedRect(0, 0, 64, 64, 14, 14)
    p.setPen(QColor("white"))
    f = QFont()
    f.setBold(True)
    f.setPixelSize(21)
    p.setFont(f)
    p.drawText(pix.rect(), Qt.AlignCenter, "MOS")
    p.end()
    return QIcon(pix)


# ====================================================================== trang chủ


class HomePage(QScrollArea):
    def __init__(self, shell: Shell):
        super().__init__()
        self.shell = shell
        acc = shell.account
        custom, errors = load_custom_exams()
        self.exams = merge_exams(ALL_EXAMS, custom)
        history = core.load_history(acc.username)
        self.exam_idx, self.mode = 0, "training"

        body, lay = page_body()
        first = acc.ho_ten.split()[-1] if acc.ho_ten.split() else acc.username
        lay.addWidget(T.page_header(f"Xin chào, {first}!",
                                    "Chọn bài thi và chế độ rồi bấm Bắt đầu. Phần mềm sẽ mở Word / Excel / "
                                    "PowerPoint để bạn làm bài trực tiếp."))
        if errors and acc.is_admin:
            lay.addWidget(_banner("Có đề tự soạn bị lỗi nên chưa được nạp: " + errors[0] +
                                  (f" (và {len(errors) - 1} lỗi khác)" if len(errors) > 1 else "") +
                                  ". Vào Đề thi để sửa.", tone="warn"))

        stats = QHBoxLayout()
        stats.setSpacing(16)
        passed = sum(h["passed"] for h in history)
        best = max((h["score"] for h in history), default=None)
        stats.addWidget(T.stat_card("Số lần làm bài", str(len(history)), "tính cả luyện tập và thi thử"))
        stats.addWidget(T.stat_card("Điểm cao nhất", f"{best}" if best is not None else "—", "thang điểm 1000",
                                    SUCCESS))
        stats.addWidget(T.stat_card("Tỉ lệ đạt", f"{round(100 * passed / len(history))}%" if history else "—",
                                    "điểm đạt từ 700", WARN))
        lay.addLayout(stats)

        lay.addWidget(label("Chọn bài thi", "h2"))
        grid = QGridLayout()
        grid.setSpacing(16)
        self.exam_cards = []
        for i, exam in enumerate(self.exams):
            card = self._exam_card(exam, history)
            card.clicked.connect(lambda i=i: self.select_exam(i))
            grid.addWidget(card, i // 3, i % 3)
            self.exam_cards.append((card, exam))
        for c in range(3):
            grid.setColumnStretch(c, 1)
        lay.addLayout(grid)

        lay.addWidget(label("Chế độ", "h2"))
        modes = QHBoxLayout()
        modes.setSpacing(16)
        self.mode_cards = {}
        for key, title, desc in (("training", "Luyện tập", "Không giới hạn giờ · có gợi ý · kiểm tra từng dự án"),
                                 ("testing", "Thi thử", "Tính giờ như thi thật · không gợi ý · chấm khi nộp bài")):
            card = ClickableCard(padding=18, spacing=14, horizontal=True)
            dot = QLabel()
            dot.setFixedSize(20, 20)
            card.lay.addWidget(dot)
            col = QVBoxLayout()
            col.setSpacing(2)
            col.addWidget(label(title, "h3"))
            col.addWidget(label(desc, "caption"))
            card.lay.addLayout(col, 1)
            card.clicked.connect(lambda k=key: self.select_mode(k))
            modes.addWidget(card)
            self.mode_cards[key] = (card, dot)
        lay.addLayout(modes)

        foot = QHBoxLayout()
        if history:
            h = history[-1]
            foot.addWidget(label(f"Lần gần nhất: {h['exam']} — {h['score']}/1000 "
                                 f"({'Đạt' if h['passed'] else 'Chưa đạt'}) · {h['time']}", "muted"))
        else:
            foot.addWidget(label("Mẹo: làm xong mỗi dự án nhớ bấm Ctrl + S để lưu file.", "muted"))
        foot.addStretch()
        self.start_btn = button("Bắt đầu làm bài  →", self.start, "primary", "lg")
        foot.addWidget(self.start_btn)
        lay.addLayout(foot)
        lay.addStretch()

        self.setObjectName("Page")
        self.setWidgetResizable(True)
        self.setFrameShape(QFrame.NoFrame)
        self.setWidget(body)
        self.select_exam(0)
        self.select_mode("training")

    def _exam_card(self, exam, history) -> ClickableCard:
        letter, color, tint, short = T.brand(exam.code)
        card = ClickableCard(padding=20, spacing=6)
        top = QHBoxLayout()
        top.addWidget(T.badge(exam.code, 46))
        top.addStretch()
        card.check = chip("✓ Đã chọn", color, tint)
        top.addWidget(card.check, 0, Qt.AlignTop)
        card.lay.addLayout(top)
        card.lay.addSpacing(6)
        name, _, code = exam.name.replace("Microsoft ", "").partition(" (")
        card.lay.addWidget(label(name, "h2"))
        card.lay.addWidget(label(code.rstrip(")") or "Đề tự soạn", "caption"))
        n_tasks = sum(len(p.tasks) for p in exam.projects)
        card.lay.addWidget(label(f"{len(exam.projects)} dự án · {n_tasks} nhiệm vụ · {exam.minutes} phút", "muted"))
        card.lay.addSpacing(8)
        best = max((h["score"] for h in history if h["exam"] == exam.name), default=None)
        card.lay.addWidget(T.progress(best or 0, tone=None if (best or 0) >= core.PASS_SCORE else "warn"))
        card.lay.addWidget(label(f"Điểm cao nhất: {best}/1000" if best is not None else "Chưa làm bài này",
                                 "caption"))
        card.brand_color, card.tint = color, tint
        return card

    def select_exam(self, idx):
        self.exam_idx = idx
        for i, (card, _) in enumerate(self.exam_cards):
            on = i == idx
            card.setStyleSheet(f"QFrame#Card {{ background: white; border: 2px solid {card.brand_color}; "
                               "border-radius: 14px; }" if on else "")
            card.check.setVisible(on)

    def select_mode(self, key):
        self.mode = key
        for k, (card, dot) in self.mode_cards.items():
            on = k == key
            card.setStyleSheet(f"QFrame#Card {{ background: {PRIMARY_SOFT}; border: 2px solid {PRIMARY}; "
                               "border-radius: 14px; }" if on else "")
            dot.setStyleSheet(f"border-radius: 10px; border: {'6px' if on else '2px'} solid "
                              f"{PRIMARY if on else '#D0D5DD'}; background: white;")

    def start(self):
        self.shell.start_exam(self.exams[self.exam_idx], self.mode)


def _banner(text, tone="info") -> QFrame:
    f = QFrame()
    colors = {"info": ("#1E3A8A", PRIMARY_SOFT, "#D1E0FF"), "warn": (WARN, WARN_SOFT, "#FEDF89"),
              "bad": (DANGER, DANGER_SOFT, "#FECDCA"), "ok": (SUCCESS, SUCCESS_SOFT, "#ABEFC6")}[tone]
    f.setStyleSheet(f"QFrame {{ background:{colors[1]}; border:1px solid {colors[2]}; border-radius:10px; }}"
                    f"QLabel {{ color:{colors[0]}; background: transparent; }}")
    lay = QHBoxLayout(f)
    lay.setContentsMargins(14, 10, 14, 10)
    lay.addWidget(label(text, wrap=True), 1)
    return f


# ====================================================================== lịch sử


class HistoryPage(QScrollArea):
    def __init__(self, shell: Shell):
        super().__init__()
        history = core.load_history(shell.account.username)
        body, lay = page_body()
        lay.addWidget(T.page_header("Lịch sử của tôi", "Tất cả các lần bạn đã nộp bài."))
        stats = QHBoxLayout()
        stats.setSpacing(16)
        passed = sum(h["passed"] for h in history)
        stats.addWidget(T.stat_card("Số lần làm bài", str(len(history))))
        stats.addWidget(T.stat_card("Điểm cao nhất", str(max((h["score"] for h in history), default="—")),
                                    accent=SUCCESS))
        stats.addWidget(T.stat_card("Điểm trung bình",
                                    str(round(sum(h["score"] for h in history) / len(history))) if history else "—",
                                    accent="#7A5AF8"))
        stats.addWidget(T.stat_card("Tỉ lệ đạt", f"{round(100 * passed / len(history))}%" if history else "—",
                                    accent=WARN))
        lay.addLayout(stats)
        if history:
            t = T.table([("Thời gian", 150), ("Bài thi", None), ("Chế độ", 110), ("Làm trong", 100),
                         ("Điểm", 90), ("Kết quả", 110)])
            rows = [[h["time"], h["exam"], "Luyện tập" if h["mode"] == "training" else "Thi thử",
                     T.fmt_time(h.get("duration")), f'{h["score"]}', "Đạt" if h["passed"] else "Chưa đạt"]
                    for h in reversed(history)]
            T.set_rows(t, rows, ["ok" if h["passed"] else "bad" for h in reversed(history)])
            t.setMinimumHeight(360)
            lay.addWidget(t, 1)
        else:
            lay.addWidget(T.empty_state("Bạn chưa nộp bài lần nào. Vào Trang chủ để bắt đầu."))
            lay.addStretch()
        self.setObjectName("Page")
        self.setWidgetResizable(True)
        self.setFrameShape(QFrame.NoFrame)
        self.setWidget(body)


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
        info.addWidget(chip("ĐẠT" if passed else "CHƯA ĐẠT", SUCCESS if passed else DANGER,
                            SUCCESS_SOFT if passed else DANGER_SOFT), 0, Qt.AlignLeft)
        info.addWidget(label(session.exam.name, "h1"))
        info.addWidget(label("Chúc mừng! Bạn đã vượt qua điểm đạt." if passed else
                             f"Cần thêm {core.PASS_SCORE - report['score']} điểm để đạt. "
                             "Xem lại các nhiệm vụ sai bên dưới.", "muted", wrap=True))
        info.addSpacing(8)
        boxes = QHBoxLayout()
        boxes.setSpacing(10)
        for title, value in (("Đúng", f'{report["correct"]}/{report["total"]}'),
                             ("Thời gian", T.fmt_time(report.get("duration", 0))),
                             ("Điểm đạt", str(core.PASS_SCORE)),
                             ("Chế độ", "Luyện tập" if session.mode == "training" else "Thi thử")):
            b = QFrame()
            b.setObjectName("Soft")
            bl = QVBoxLayout(b)
            bl.setContentsMargins(14, 10, 14, 10)
            bl.setSpacing(0)
            bl.addWidget(label(title, "caption"))
            v = label(value)
            v.setStyleSheet("font-size: 13pt; font-weight: 700; background: transparent;")
            bl.addWidget(v)
            boxes.addWidget(b)
        boxes.addStretch()
        info.addLayout(boxes)
        head.lay.addLayout(info, 1)
        lay.addWidget(head)

        # --- theo dự án
        per = QHBoxLayout()
        per.setSpacing(16)
        for project in session.exam.projects:
            rs = [r for r in self.results if r.project == project.name]
            ok = sum(r.correct for r in rs)
            c = Card(padding=16, spacing=8)
            row = QHBoxLayout()
            row.addWidget(label(project.name, "h3"), 1)
            row.addWidget(label(f"{ok}/{len(rs)}", "muted"))
            c.lay.addLayout(row)
            ratio = ok / len(rs) if rs else 0
            c.lay.addWidget(T.progress(int(ratio * 100), 100, None if ratio >= 0.7 else "warn"))
            per.addWidget(c)
        lay.addLayout(per)

        # --- chi tiết
        lay.addWidget(label("Chi tiết từng nhiệm vụ", "h2"))
        self.table = T.table([("Kết quả", 110), ("Dự án", 220), ("Nhiệm vụ", None)])
        rows = [["✓  Đúng" if r.correct else "✗  Sai", r.project.split("–")[-1].strip(), r.task]
                for r in self.results]
        T.set_rows(self.table, rows, ["ok" if r.correct else "bad" for r in self.results], tone_cols={0})
        self.table.setMinimumHeight(min(60 + 42 * len(rows), 420))
        self.table.itemSelectionChanged.connect(self.on_select)
        lay.addWidget(self.table)
        self.hint = _banner("Bấm vào một nhiệm vụ ở bảng trên để xem cách làm.", "warn")
        self.hint_label = self.hint.findChild(QLabel)
        lay.addWidget(self.hint)

        btns = QHBoxLayout()
        btns.addWidget(button("Mở thư mục bài làm", lambda: core.open_in_office(session.workdir)))
        btns.addStretch()
        btns.addWidget(button("Làm lại đề này", lambda: shell.start_exam(session.exam, session.mode)))
        btns.addWidget(button("Về trang chủ", lambda: shell.go("home"), "primary"))
        lay.addLayout(btns)

        self.setObjectName("Page")
        self.setWidgetResizable(True)
        self.setFrameShape(QFrame.NoFrame)
        self.setWidget(body)

    def on_select(self):
        rows = self.table.selectionModel().selectedRows()
        if rows:
            r = self.results[rows[0].row()]
            text = f"<b>Cách làm:</b> {r.hint}"
            if r.error:
                text += f"<br><i>(Lỗi khi đọc file: {r.error})</i>"
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
        text = label(task.title, wrap=True)
        text.setStyleSheet("font-size: 10.5pt;")
        text.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Preferred)
        row.addWidget(text, 1)
        self.status = chip("", SUCCESS, SUCCESS_SOFT)
        self.status.hide()
        row.addWidget(self.status, 0, Qt.AlignVCenter)
        if bar.training:
            self.hint_btn = button("Gợi ý", self.toggle_hint, "ghost", "sm")
            row.addWidget(self.hint_btn, 0, Qt.AlignVCenter)
        self.mark = button("⚑ Đánh dấu", None, None, "sm")
        self.mark.setCheckable(True)
        self.mark.setProperty("toggle", "mark")
        self.mark.toggled.connect(self._on_mark)
        row.addWidget(self.mark, 0, Qt.AlignVCenter)
        self.done = button("✓ Đã làm", None, None, "sm")
        self.done.setCheckable(True)
        self.done.setProperty("toggle", "done")
        self.done.toggled.connect(self._on_done)
        row.addWidget(self.done, 0, Qt.AlignVCenter)
        outer.addLayout(row)
        self.hint = label(f"<b>Cách làm:</b> {task.hint}", wrap=True)
        self.hint.setStyleSheet(f"background:{WARN_SOFT}; color:#7A2E0E; border-radius:8px; padding:8px 12px;")
        self.hint.hide()
        outer.addWidget(self.hint)
        self.mark.setChecked(key in bar.s.marked)
        self.done.setChecked(key in bar.s.done)
        self._paint_num()
        if key in bar.checked:
            self.set_status(bar.checked[key])

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
        self.hint_btn.setText("Ẩn gợi ý" if self.hint.isVisible() else "Gợi ý")

    def set_status(self, ok: bool):
        T.set_chip(self.status, "✓ Đúng" if ok else "✗ Sai", SUCCESS if ok else DANGER,
                   SUCCESS_SOFT if ok else DANGER_SOFT)
        self.status.show()


class ExamBar(QWidget):
    HEIGHT = 380

    def __init__(self, session: core.Session, on_finish, on_quit):
        super().__init__(None, Qt.Window | Qt.WindowStaysOnTopHint)
        self.s, self.on_finish, self.on_quit = session, on_finish, on_quit
        self.training = session.mode == "training"
        self.idx = 0
        self.started = time.time()
        self.seconds = 0 if self.training else session.exam.minutes * 60
        self.finished = False
        self.checked: dict = {}
        _, self.color, self.tint, _ = T.brand(session.exam.code)
        self.setWindowTitle(f"MOS – {session.exam.name}")
        self.setWindowIcon(app_icon())
        self.setObjectName("Page")
        self.setAttribute(Qt.WA_StyledBackground, True)
        geo = QGuiApplication.primaryScreen().availableGeometry()
        self.setGeometry(geo.x(), geo.y() + geo.height() - self.HEIGHT - 30, geo.width(), self.HEIGHT)

        lay = QVBoxLayout(self)
        lay.setContentsMargins(0, 0, 0, 0)
        lay.setSpacing(0)

        # --- đầu: môn, chế độ, tab dự án, đồng hồ
        head = QWidget()
        head.setObjectName("BarHeader")
        head.setAttribute(Qt.WA_StyledBackground, True)
        hl = QHBoxLayout(head)
        hl.setContentsMargins(16, 10, 16, 10)
        hl.setSpacing(12)
        hl.addWidget(T.badge(session.exam.code, 30))
        name = label(session.exam.name)
        name.setStyleSheet("font-weight: 800; font-size: 11pt;")
        hl.addWidget(name)
        hl.addWidget(chip("LUYỆN TẬP" if self.training else "THI THỬ", "white",
                          PRIMARY if self.training else "#7A5AF8"))
        hl.addSpacing(18)
        self.tabs = QButtonGroup(self)
        for i, p in enumerate(session.exam.projects):
            short = p.name.split("–")[-1].strip()
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
        self.timer_lbl = QLabel()
        hl.addWidget(self.timer_lbl)
        lay.addWidget(head)

        # --- mô tả dự án
        intro = QWidget()
        intro.setAttribute(Qt.WA_StyledBackground, True)
        intro.setStyleSheet(f"background: {self.tint};")
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
        fl.addWidget(button("Mở file", self.open_file))
        fl.addWidget(button("Làm lại dự án", self.reset_project))
        if self.training:
            fl.addWidget(button("✓ Kiểm tra dự án", self.check_current, "ghost"))
        fl.addSpacing(12)
        self.progress_lbl = label("", "muted")
        fl.addWidget(self.progress_lbl)
        fl.addStretch()
        self.prev_btn = button("‹  Dự án trước", lambda: self.go(self.idx - 1))
        self.next_btn = button("Dự án sau  ›", lambda: self.go(self.idx + 1))
        fl.addWidget(self.prev_btn)
        fl.addWidget(self.next_btn)
        fl.addSpacing(8)
        fl.addWidget(button("Nộp bài", self.submit, "primary"))
        lay.addWidget(foot)

        self.timer = QTimer(self)
        self.timer.timeout.connect(self.tick)
        self.timer.start(1000)
        self._render_timer()
        self.show_project()
        QTimer.singleShot(0, self.open_file)

    # ------------------------------------------------------------ hiển thị
    def show_project(self):
        project = self.s.exam.projects[self.idx]
        n = len(self.s.exam.projects)
        self.tabs.button(self.idx).setChecked(True)
        self.intro.setText(f"<b>{project.name}</b> — {project.intro}  <i>Làm xong nhớ lưu file (Ctrl + S).</i>")
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
        text = f"Đã làm {len(self.s.done)}/{total}"
        if self.s.marked:
            text += f"  ·  Đánh dấu {len(self.s.marked)}"
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
            T.warn(self, "Hết giờ", "Đã hết thời gian. Bài sẽ được nộp tự động.\n"
                                    "(Các thay đổi chưa lưu sẽ không được tính.)")
            self.finish()

    # ------------------------------------------------------------ thao tác
    def open_file(self):
        path = self.s.file_of(self.idx)
        try:
            core.open_in_office(path)
        except OSError as exc:
            T.error(self, "Không mở được file", f"{exc}\n\nHãy tự mở file:\n{path}")

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
            T.warn(self, "File đang mở", "Hãy ĐÓNG file trong Office trước khi làm lại dự án.")
            return
        if T.confirm(self, "Làm lại dự án", "Tạo lại file gốc cho dự án này?\n(Bài cũ được lưu với tên *_cu)"):
            core.reset_project(self.s, self.idx)
            for t in range(len(self.s.exam.projects[self.idx].tasks)):
                self.checked.pop((self.idx, t), None)
            self.show_project()
            self.open_file()

    def _confirm_saved(self, path) -> bool:
        if core.file_is_open(path):
            return T.confirm(self, "File đang mở", "File vẫn đang mở trong Office. Chỉ những gì đã LƯU mới "
                                                    "được chấm.\n\nBạn đã bấm Ctrl + S chưa? Chọn Đồng ý để chấm.")
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
            T.warn(self, "Không đọc được file", f"Lỗi: {errs[0]}\n\nHãy lưu lại file rồi thử lại.")

    def submit(self):
        opened = [i for i in range(len(self.s.exam.projects)) if core.file_is_open(self.s.file_of(i))]
        msg = "Nộp bài và chấm điểm?"
        if opened:
            msg = "Vẫn còn file đang mở trong Office. Chỉ phần đã LƯU mới được chấm.\n\n" + msg
        if self.s.marked:
            msg = f"Bạn còn {len(self.s.marked)} nhiệm vụ đánh dấu xem lại.\n" + msg
        if T.confirm(self, "Nộp bài", msg):
            self.finish()

    def finish(self):
        self.finished = True
        self.timer.stop()
        report = core.grade(self.s)
        report["duration"] = int(time.time() - self.started)
        try:
            core.save_history(report)
        except OSError:
            pass
        self.close()
        self.on_finish(self.s, report)

    def closeEvent(self, e):
        if self.finished:
            return super().closeEvent(e)
        if T.confirm(self, "Thoát bài thi", "Thoát bài thi? Kết quả sẽ không được chấm."):
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
    app = QApplication.instance() or QApplication(sys.argv)
    app.setApplicationName("Luyện thi MOS")
    app.setStyle("Fusion")
    family = T.pick_font_family()
    app.setFont(QFont(family, 10))
    app.setStyleSheet(T.stylesheet(family))
    win = MainWindow(AccountStore())
    win.show()
    sys.exit(app.exec())
