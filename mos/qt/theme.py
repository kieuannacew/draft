"""Hệ thống thiết kế: màu, font, stylesheet (QSS) và các widget nhỏ dùng chung."""
from __future__ import annotations

from PySide6.QtCore import QRectF, Qt, Signal
from PySide6.QtGui import QColor, QFont, QFontDatabase, QPainter, QPen
from PySide6.QtWidgets import (QAbstractItemView, QApplication, QFrame, QGraphicsDropShadowEffect, QHBoxLayout,
                               QHeaderView, QLabel, QMessageBox, QProgressBar, QPushButton, QSizePolicy,
                               QTableWidget, QTableWidgetItem, QVBoxLayout, QWidget)

# ---------------------------------------------------------------- màu
BG = "#F4F6FA"
SURFACE = "#FFFFFF"
SURFACE_2 = "#F9FAFB"
BORDER = "#E4E7EC"
TEXT = "#101828"
TEXT_2 = "#344054"
MUTED = "#667085"
FAINT = "#98A2B3"
PRIMARY = "#2563EB"
PRIMARY_HOVER = "#1D4ED8"
PRIMARY_SOFT = "#EFF4FF"
SIDEBAR = "#0B1220"
SIDEBAR_2 = "#141C2F"
SUCCESS, SUCCESS_SOFT = "#079455", "#ECFDF3"
DANGER, DANGER_SOFT = "#D92D20", "#FEF3F2"
WARN, WARN_SOFT = "#B54708", "#FFFAEB"

# Màu nhận diện từng môn: (chữ cái, màu chính, nền nhạt, tên ngắn)
BRAND = {
    "WORD": ("W", "#2B579A", "#EAF0FA", "Word"),
    "EXCEL": ("X", "#1D6F42", "#E8F4EC", "Excel"),
    "POWERPOINT": ("P", "#C43E1C", "#FCEEE9", "PowerPoint"),
}


def brand(code: str):
    return BRAND.get(code, ("?", PRIMARY, PRIMARY_SOFT, code))


def pick_font_family() -> str:
    families = set(QFontDatabase.families())
    for fam in ("Segoe UI Variable Text", "Segoe UI", "Inter", "SF Pro Text", "Noto Sans", "DejaVu Sans"):
        if fam in families:
            return fam
    return QApplication.font().family()


def _assets() -> dict[str, str]:
    """Ghi vài icon SVG nhỏ ra thư mục tạm để QSS dùng (url(...))."""
    import tempfile
    from pathlib import Path
    folder = Path(tempfile.gettempdir()) / "mos_qt_assets"
    folder.mkdir(exist_ok=True)
    icons = {
        "chevron": f'<svg xmlns="http://www.w3.org/2000/svg" width="12" height="12" viewBox="0 0 12 12">'
                   f'<path d="M2.5 4.5 6 8l3.5-3.5" fill="none" stroke="{MUTED}" stroke-width="1.6" '
                   f'stroke-linecap="round" stroke-linejoin="round"/></svg>',
        "check": '<svg xmlns="http://www.w3.org/2000/svg" width="12" height="12" viewBox="0 0 12 12">'
                 '<path d="M2.5 6.2 5 8.5l4.5-5" fill="none" stroke="white" stroke-width="1.8" '
                 'stroke-linecap="round" stroke-linejoin="round"/></svg>',
    }
    out = {}
    for name, svg in icons.items():
        path = folder / f"{name}.svg"
        path.write_text(svg, encoding="utf-8")
        out[name] = path.as_posix()
    return out


def stylesheet(family: str) -> str:
    icons = _assets()
    return f"""
    * {{ font-family: "{family}"; font-size: 10pt; color: {TEXT}; }}
    QMainWindow, QDialog, QWidget#Page, QScrollArea#Page > QWidget > QWidget {{ background: {BG}; }}
    QScrollArea {{ border: none; background: transparent; }}
    QToolTip {{ background: {TEXT}; color: white; border: none; padding: 6px 8px; border-radius: 6px; }}

    /* ---------- chữ */
    QLabel[role="h1"] {{ font-size: 20pt; font-weight: 700; }}
    QLabel[role="h2"] {{ font-size: 13pt; font-weight: 700; }}
    QLabel[role="h3"] {{ font-size: 11pt; font-weight: 700; }}
    QLabel[role="muted"] {{ color: {MUTED}; }}
    QLabel[role="caption"] {{ color: {MUTED}; font-size: 9pt; }}
    QLabel[role="overline"] {{ color: {FAINT}; font-size: 8pt; font-weight: 700; letter-spacing: 1px; }}
    QLabel[role="stat"] {{ font-size: 20pt; font-weight: 700; }}

    /* ---------- thẻ */
    QFrame#Card {{ background: {SURFACE}; border: 1px solid {BORDER}; border-radius: 14px; }}
    QFrame#Card QLabel {{ background: transparent; }}
    QFrame#Soft {{ background: {SURFACE_2}; border: 1px solid {BORDER}; border-radius: 10px; }}
    QFrame#Banner {{ background: {PRIMARY_SOFT}; border: 1px solid #D1E0FF; border-radius: 10px; }}
    QFrame#Banner QLabel {{ color: #1E3A8A; background: transparent; }}

    /* ---------- nút */
    QPushButton {{ background: {SURFACE}; border: 1px solid {BORDER}; border-radius: 9px;
                   padding: 8px 16px; font-weight: 600; color: {TEXT_2}; }}
    QPushButton:hover {{ background: {SURFACE_2}; border-color: #D0D5DD; }}
    QPushButton:pressed {{ background: #EAECF0; }}
    QPushButton:disabled {{ color: {FAINT}; background: {SURFACE_2}; border-color: {BORDER}; }}
    QPushButton[kind="primary"] {{ background: {PRIMARY}; color: white; border: 1px solid {PRIMARY}; }}
    QPushButton[kind="primary"]:hover {{ background: {PRIMARY_HOVER}; border-color: {PRIMARY_HOVER}; }}
    QPushButton[kind="primary"]:disabled {{ background: #B2CCFF; border-color: #B2CCFF; color: white; }}
    QPushButton[kind="danger"] {{ color: {DANGER}; }}
    QPushButton[kind="danger"]:hover {{ background: {DANGER_SOFT}; border-color: #FDA29B; }}
    QPushButton[kind="ghost"] {{ background: transparent; border: none; color: {PRIMARY}; padding: 6px 10px; }}
    QPushButton[kind="ghost"]:hover {{ background: {PRIMARY_SOFT}; }}
    QPushButton[size="lg"] {{ padding: 12px 26px; font-size: 11pt; border-radius: 10px; }}
    QPushButton[size="sm"] {{ padding: 5px 10px; font-size: 9pt; border-radius: 7px; }}

    /* ---------- ô nhập */
    QLineEdit, QPlainTextEdit, QTextEdit, QComboBox, QSpinBox {{
        background: {SURFACE}; border: 1px solid #D0D5DD; border-radius: 9px; padding: 7px 10px;
        selection-background-color: #D1E0FF; selection-color: {TEXT}; }}
    QLineEdit:focus, QPlainTextEdit:focus, QTextEdit:focus, QComboBox:focus {{ border: 1px solid {PRIMARY}; }}
    QLineEdit:disabled, QPlainTextEdit:disabled, QComboBox:disabled {{ background: {SURFACE_2}; color: {FAINT}; }}
    QLineEdit[size="lg"] {{ padding: 10px 12px; font-size: 11pt; }}
    QComboBox::drop-down {{ border: none; width: 26px; }}
    QComboBox::down-arrow {{ image: url("{icons['chevron']}"); width: 12px; height: 12px; margin-right: 10px; }}
    QComboBox QAbstractItemView {{ background: {SURFACE}; border: 1px solid {BORDER}; border-radius: 8px;
                                   padding: 4px; selection-background-color: {PRIMARY_SOFT}; selection-color: {TEXT}; outline: 0; }}

    /* ---------- bảng / danh sách */
    QTableWidget, QListWidget {{ background: {SURFACE}; border: 1px solid {BORDER}; border-radius: 12px;
                                 gridline-color: transparent; outline: 0; alternate-background-color: {SURFACE_2}; }}
    QTableWidget::item {{ padding: 6px 8px; border-bottom: 1px solid #F2F4F7; }}
    QTableWidget::item:selected, QListWidget::item:selected {{ background: {PRIMARY_SOFT}; color: {TEXT}; }}
    QListWidget {{ padding: 4px; }}
    QListWidget::item {{ padding: 8px 10px; border-radius: 8px; margin: 1px 0; }}
    QListWidget::item:hover {{ background: {SURFACE_2}; }}
    QHeaderView {{ background: transparent; }}
    QHeaderView::section {{ background: {SURFACE_2}; color: {MUTED}; font-size: 9pt; font-weight: 700;
                            border: none; border-bottom: 1px solid {BORDER}; padding: 9px 8px; }}
    QHeaderView::section:first {{ border-top-left-radius: 12px; }}
    QHeaderView::section:last {{ border-top-right-radius: 12px; }}

    /* ---------- thanh cuộn */
    QScrollBar:vertical {{ background: transparent; width: 10px; margin: 4px 2px; }}
    QScrollBar::handle:vertical {{ background: #D0D5DD; border-radius: 3px; min-height: 36px; }}
    QScrollBar::handle:vertical:hover {{ background: {FAINT}; }}
    QScrollBar:horizontal {{ background: transparent; height: 10px; margin: 2px 4px; }}
    QScrollBar::handle:horizontal {{ background: #D0D5DD; border-radius: 3px; min-width: 36px; }}
    QScrollBar::add-line, QScrollBar::sub-line {{ width: 0; height: 0; }}
    QScrollBar::add-page, QScrollBar::sub-page {{ background: transparent; }}

    /* ---------- thanh tiến độ */
    QProgressBar {{ background: #EAECF0; border: none; border-radius: 4px; max-height: 8px; min-height: 8px; }}
    QProgressBar::chunk {{ background: {SUCCESS}; border-radius: 4px; }}
    QProgressBar[tone="warn"]::chunk {{ background: #F79009; }}
    QProgressBar[tone="danger"]::chunk {{ background: #F04438; }}

    QCheckBox {{ spacing: 8px; }}
    QCheckBox::indicator {{ width: 16px; height: 16px; border: 1px solid #D0D5DD; border-radius: 5px; background: white; }}
    QCheckBox::indicator:checked {{ background: {PRIMARY}; border-color: {PRIMARY}; image: url("{icons['check']}"); }}

    /* ---------- thanh bên (sidebar) */
    QWidget#Sidebar {{ background: {SIDEBAR}; }}
    QWidget#Sidebar QLabel {{ color: white; background: transparent; }}
    QWidget#Sidebar QLabel[role="caption"] {{ color: #8A94A6; }}
    QWidget#Sidebar QLabel[role="overline"] {{ color: #5B667A; }}
    QPushButton[nav="true"] {{ background: transparent; color: #C4CBD8; border: none; text-align: left;
                               padding: 10px 14px; border-radius: 9px; font-weight: 600; }}
    QPushButton[nav="true"]:hover {{ background: {SIDEBAR_2}; color: white; }}
    QPushButton[nav="true"]:checked {{ background: {PRIMARY}; color: white; }}
    QPushButton[kind="side"] {{ background: transparent; color: #AEB6C4; border: 1px solid #243049;
                                padding: 6px 10px; border-radius: 8px; font-size: 9pt; }}
    QPushButton[kind="side"]:hover {{ background: {SIDEBAR_2}; color: white; }}

    /* ---------- trang đăng nhập */
    QWidget#Login {{ background: qlineargradient(x1:0, y1:0, x2:1, y2:1, stop:0 #0B1220, stop:0.55 #16255A,
                                                 stop:1 #2563EB); }}

    /* ---------- thanh làm bài */
    QWidget#BarHeader {{ background: {SIDEBAR}; }}
    QWidget#BarHeader QLabel {{ color: white; background: transparent; }}
    QPushButton[tab="true"] {{ background: {SIDEBAR_2}; color: #C4CBD8; border: none; border-radius: 8px;
                               padding: 6px 14px; font-size: 9pt; font-weight: 700; }}
    QPushButton[tab="true"]:hover {{ color: white; }}
    QPushButton[tab="true"]:checked {{ background: white; color: {TEXT}; }}
    QWidget#BarFooter {{ background: {SURFACE}; border-top: 1px solid {BORDER}; }}
    QPushButton[toggle="mark"]:checked {{ background: {WARN_SOFT}; color: {WARN}; border-color: #FEC84B; }}
    QPushButton[toggle="done"]:checked {{ background: {SUCCESS_SOFT}; color: {SUCCESS}; border-color: #75E0A7; }}

    QSplitter::handle {{ background: transparent; width: 12px; }}
    QMessageBox {{ background: {SURFACE}; }}
    QMessageBox QLabel {{ font-size: 10pt; }}
    """


# ---------------------------------------------------------------- widget nhỏ


def label(text: str = "", role: str | None = None, wrap: bool = False, parent=None) -> QLabel:
    lb = QLabel(text, parent)
    if role:
        lb.setProperty("role", role)
    lb.setWordWrap(wrap)
    return lb


def button(text: str, on_click=None, kind: str | None = None, size: str | None = None, parent=None) -> QPushButton:
    b = QPushButton(text, parent)
    b.setCursor(Qt.PointingHandCursor)
    if kind:
        b.setProperty("kind", kind)
    if size:
        b.setProperty("size", size)
    if on_click:
        b.clicked.connect(lambda *_: on_click())
    return b


def repolish(w: QWidget) -> None:
    """Áp lại QSS sau khi đổi property."""
    w.style().unpolish(w)
    w.style().polish(w)
    w.update()


def shadow(w: QWidget, blur=28, y=6, alpha=18) -> QWidget:
    eff = QGraphicsDropShadowEffect(w)
    eff.setBlurRadius(blur)
    eff.setOffset(0, y)
    eff.setColor(QColor(16, 24, 40, alpha))
    w.setGraphicsEffect(eff)
    return w


class Card(QFrame):
    def __init__(self, parent=None, padding=20, spacing=10, raised=False, horizontal=False):
        super().__init__(parent)
        self.setObjectName("Card")
        lay = QHBoxLayout(self) if horizontal else QVBoxLayout(self)
        lay.setContentsMargins(padding, padding, padding, padding)
        lay.setSpacing(spacing)
        if raised:
            shadow(self)

    @property
    def lay(self):
        return self.layout()


class ClickableCard(Card):
    clicked = Signal()

    def __init__(self, *a, **kw):
        super().__init__(*a, **kw)
        self.setCursor(Qt.PointingHandCursor)

    def mouseReleaseEvent(self, e):
        if e.button() == Qt.LeftButton:
            self.clicked.emit()
        super().mouseReleaseEvent(e)


def chip(text: str, fg: str, bg: str, bold=True, parent=None) -> QLabel:
    lb = QLabel(text, parent)
    lb.setStyleSheet(f"background:{bg}; color:{fg}; border-radius:10px; padding:3px 10px;"
                     f"font-size:9pt; font-weight:{'700' if bold else '500'};")
    lb.setSizePolicy(QSizePolicy.Maximum, QSizePolicy.Fixed)
    return lb


def set_chip(lb: QLabel, text: str, fg: str, bg: str) -> None:
    lb.setText(text)
    lb.setStyleSheet(f"background:{bg}; color:{fg}; border-radius:10px; padding:3px 10px;"
                     "font-size:9pt; font-weight:700;")


def badge(code: str, size=44, parent=None) -> QLabel:
    letter, color, _, _ = brand(code)
    lb = QLabel(letter, parent)
    lb.setFixedSize(size, size)
    lb.setAlignment(Qt.AlignCenter)
    lb.setStyleSheet(f"background:{color}; color:white; border-radius:{size // 4}px;"
                     f"font-size:{max(9, int(size / 2.6))}pt; font-weight:800;")
    return lb


def avatar(name: str, size=36, color=PRIMARY, parent=None) -> QLabel:
    parts = [p for p in name.split() if p]
    initials = (parts[0][0] + (parts[-1][0] if len(parts) > 1 else "")).upper() if parts else "?"
    lb = QLabel(initials, parent)
    lb.setFixedSize(size, size)
    lb.setAlignment(Qt.AlignCenter)
    lb.setStyleSheet(f"background:{color}; color:white; border-radius:{size // 2}px; font-weight:700;"
                     f"font-size:{int(size / 3.2)}pt;")
    return lb


def progress(value: int, maximum=1000, tone: str | None = None) -> QProgressBar:
    bar = QProgressBar()
    bar.setRange(0, maximum)
    bar.setValue(value)
    bar.setTextVisible(False)
    if tone:
        bar.setProperty("tone", tone)
    return bar


def hline() -> QFrame:
    f = QFrame()
    f.setFixedHeight(1)
    f.setStyleSheet(f"background:{BORDER}; border:none;")
    return f


def stat_card(title: str, value: str, hint: str = "", accent: str = PRIMARY) -> Card:
    c = Card(padding=18, spacing=4)
    top = QHBoxLayout()
    dot = QLabel()
    dot.setFixedSize(8, 8)
    dot.setStyleSheet(f"background:{accent}; border-radius:4px;")
    top.addWidget(dot)
    top.addWidget(label(title, "caption"))
    top.addStretch()
    c.lay.addLayout(top)
    c.lay.addWidget(label(value, "stat"))
    if hint:
        c.lay.addWidget(label(hint, "caption"))
    return c


def page_header(title: str, subtitle: str = "", actions: list[QWidget] | None = None) -> QWidget:
    w = QWidget()
    lay = QHBoxLayout(w)
    lay.setContentsMargins(0, 0, 0, 0)
    col = QVBoxLayout()
    col.setSpacing(2)
    col.addWidget(label(title, "h1"))
    if subtitle:
        col.addWidget(label(subtitle, "muted", wrap=True))
    lay.addLayout(col, 1)
    for a in actions or []:
        lay.addWidget(a, 0, Qt.AlignBottom)
    return w


class ScoreRing(QWidget):
    """Vòng tròn điểm (0–1000)."""

    def __init__(self, score: int, passed: bool, size=168, parent=None):
        super().__init__(parent)
        self.score, self.passed = score, passed
        self.setFixedSize(size, size)

    def paintEvent(self, _):
        p = QPainter(self)
        p.setRenderHint(QPainter.Antialiasing)
        w = 14
        rect = QRectF(w / 2 + 2, w / 2 + 2, self.width() - w - 4, self.height() - w - 4)
        p.setPen(QPen(QColor("#EAECF0"), w, Qt.SolidLine, Qt.RoundCap))
        p.drawEllipse(rect)
        if self.score:
            p.setPen(QPen(QColor(SUCCESS if self.passed else "#F04438"), w, Qt.SolidLine, Qt.RoundCap))
            p.drawArc(rect, 90 * 16, int(-360 * 16 * min(self.score, 1000) / 1000))
        p.setPen(QColor(TEXT))
        f = QFont(self.font())
        f.setPointSize(26 if self.score < 1000 else 22)
        f.setBold(True)
        p.setFont(f)
        p.drawText(QRectF(0, -8, self.width(), self.height()), Qt.AlignCenter, str(self.score))
        f.setPointSize(9)
        f.setBold(False)
        p.setFont(f)
        p.setPen(QColor(MUTED))
        p.drawText(QRectF(0, 30, self.width(), self.height()), Qt.AlignCenter, "/ 1000 điểm")


# ---------------------------------------------------------------- hộp thoại thông báo


def _box(parent, icon, title, text, buttons, default):
    box = QMessageBox(icon, title, text, buttons, parent)
    box.setDefaultButton(default)
    for b in box.buttons():
        b.setCursor(Qt.PointingHandCursor)
        role = box.buttonRole(b)
        if role in (QMessageBox.AcceptRole, QMessageBox.YesRole):
            b.setProperty("kind", "primary")
        b.setMinimumWidth(90)
    tr = {QMessageBox.Yes: "Đồng ý", QMessageBox.No: "Không", QMessageBox.Ok: "OK", QMessageBox.Cancel: "Hủy"}
    for std, txt in tr.items():
        b = box.button(std)
        if b:
            b.setText(txt)
    return box.exec()


def info(parent, title, text):
    _box(parent, QMessageBox.Information, title, text, QMessageBox.Ok, QMessageBox.Ok)


def warn(parent, title, text):
    _box(parent, QMessageBox.Warning, title, text, QMessageBox.Ok, QMessageBox.Ok)


def error(parent, title, text):
    _box(parent, QMessageBox.Critical, title, text, QMessageBox.Ok, QMessageBox.Ok)


def confirm(parent, title, text) -> bool:
    return _box(parent, QMessageBox.Question, title, text, QMessageBox.Yes | QMessageBox.No,
                QMessageBox.Yes) == QMessageBox.Yes


# ---------------------------------------------------------------- bảng


TONES = {"ok": SUCCESS, "bad": DANGER, "warn": WARN, "muted": FAINT}


def table(columns: list[tuple[str, int | None]], parent=None) -> QTableWidget:
    """columns: [(tiêu đề, độ rộng)] – độ rộng None = co giãn."""
    t = QTableWidget(0, len(columns), parent)
    t.setHorizontalHeaderLabels([c[0] for c in columns])
    t.verticalHeader().hide()
    t.setSelectionBehavior(QAbstractItemView.SelectRows)
    t.setSelectionMode(QAbstractItemView.SingleSelection)
    t.setEditTriggers(QAbstractItemView.NoEditTriggers)
    t.setShowGrid(False)
    t.setFocusPolicy(Qt.NoFocus)
    t.setWordWrap(False)
    t.verticalHeader().setDefaultSectionSize(42)
    head = t.horizontalHeader()
    head.setHighlightSections(False)
    head.setDefaultAlignment(Qt.AlignLeft | Qt.AlignVCenter)
    for i, (_, width) in enumerate(columns):
        if width is None:
            head.setSectionResizeMode(i, QHeaderView.Stretch)
        else:
            head.setSectionResizeMode(i, QHeaderView.Interactive)
            t.setColumnWidth(i, width)
    return t


def set_rows(t: QTableWidget, rows: list[list], tones: list[str | None] | None = None,
             tone_cols: set[int] | None = None, keys: list | None = None) -> None:
    """Đổ dữ liệu vào bảng. tones[i] tô màu các cột tone_cols (mặc định: cột cuối) của dòng i."""
    t.setRowCount(0)
    t.setRowCount(len(rows))
    for r, row in enumerate(rows):
        tone = tones[r] if tones else None
        for c, value in enumerate(row):
            item = QTableWidgetItem("" if value is None else str(value))
            if keys is not None and c == 0:
                item.setData(Qt.UserRole, keys[r])
            if tone and c in (tone_cols if tone_cols is not None else {len(row) - 1}):
                item.setForeground(QColor(TONES[tone]))
                f = item.font()
                f.setBold(True)
                item.setFont(f)
            t.setItem(r, c, item)


def selected_key(t: QTableWidget):
    rows = t.selectionModel().selectedRows()
    if not rows:
        return None
    item = t.item(rows[0].row(), 0)
    return item.data(Qt.UserRole) if item else None


def empty_state(text: str) -> QWidget:
    w = QFrame()
    w.setObjectName("Soft")
    lay = QVBoxLayout(w)
    lay.setContentsMargins(24, 28, 24, 28)
    lb = label(text, "muted", wrap=True)
    lb.setAlignment(Qt.AlignCenter)
    lay.addWidget(lb)
    return w


def fmt_time(seconds) -> str:
    if seconds is None:
        return "—"
    m, s = divmod(max(int(seconds), 0), 60)
    return f"{m:02d}:{s:02d}"
