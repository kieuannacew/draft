"""Hệ thống thiết kế "Forest Canopy": màu, font, stylesheet (QSS) và các widget nhỏ dùng chung.

Bảng màu (theme-factory · Forest Canopy):
    Forest Green #2d4a2b – màu chủ đạo (nút chính, thanh bên)
    Sage         #7d8471 – chữ phụ, đường nét dịu
    Olive        #a4ac86 – điểm nhấn sáng (mục đang chọn, huy hiệu)
    Ivory        #faf9f6 – nền
Chữ: Be Vietnam Pro (đóng gói kèm app trong mos/qt/fonts, thiết kế cho tiếng Việt, giấy phép OFL).
"""
from __future__ import annotations

import math
import random

from PySide6.QtCore import QPointF, QRectF, Qt, QTimer, Signal
from PySide6.QtGui import QColor, QFont, QFontDatabase, QPainter, QPainterPath, QPen
from PySide6.QtWidgets import (QAbstractItemView, QApplication, QButtonGroup, QFrame, QGraphicsDropShadowEffect,
                               QHBoxLayout, QHeaderView, QLabel, QMessageBox, QProgressBar, QPushButton,
                               QSizePolicy, QTableWidget, QTableWidgetItem, QVBoxLayout, QWidget)

from ..i18n import tr

# ---------------------------------------------------------------- màu
FOREST = "#2D4A2B"
SAGE = "#7D8471"
OLIVE = "#A4AC86"
IVORY = "#FAF9F6"

BG = IVORY
SURFACE = "#FFFFFF"
SURFACE_2 = "#F3F2EC"
BORDER = "#E3E1D6"
TEXT = "#1E2A1D"
TEXT_2 = "#3A4536"
MUTED = "#5F6857"          # sage đậm hơn một chút để đủ tương phản trên nền ngà
FAINT = "#9CA38A"
PRIMARY = FOREST
PRIMARY_HOVER = "#223A20"
PRIMARY_SOFT = "#ECEFE3"
PRIMARY_LINE = "#CDD3BC"
SIDEBAR = "#1F3520"
SIDEBAR_2 = "#2D4A2B"
SUCCESS, SUCCESS_SOFT = "#3E7B3A", "#EAF3E4"
DANGER, DANGER_SOFT = "#A63D2B", "#F8ECE8"
WARN, WARN_SOFT = "#94650F", "#FBF3E0"
GOLD, GOLD_SOFT = "#C99A2E", "#FBF4E2"

# Màu nhận diện từng môn: (chữ cái, màu chính, nền nhạt, tên ngắn)
BRAND = {
    "WORD": ("W", "#2B579A", "#EAF0FA", "Word"),
    "EXCEL": ("X", "#1D6F42", "#E8F4EC", "Excel"),
    "POWERPOINT": ("P", "#C43E1C", "#FCEEE9", "PowerPoint"),
}


def brand(code: str):
    return BRAND.get(code, ("?", PRIMARY, PRIMARY_SOFT, code))


def _first_family(candidates) -> str | None:
    families = set(QFontDatabase.families())
    return next((f for f in candidates if f in families), None)


FONT_FAMILY = "Be Vietnam Pro"


def load_fonts() -> None:
    """Nạp font đóng gói kèm app (mos/qt/fonts/*.ttf) để máy nào cũng hiển thị giống nhau."""
    from pathlib import Path
    for f in sorted((Path(__file__).parent / "fonts").glob("*.ttf")):
        QFontDatabase.addApplicationFont(str(f))


def pick_font_family() -> str:
    """Font nội dung (ưu tiên font đóng gói kèm)."""
    return _first_family((FONT_FAMILY, "Segoe UI", "Arial", "Noto Sans", "DejaVu Sans")) or \
        QApplication.font().family()


def pick_heading_family() -> str:
    """Font tiêu đề: cùng họ chữ, in đậm qua QSS."""
    return pick_font_family()


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


def stylesheet(family: str, heading: str | None = None) -> str:
    icons = _assets()
    heading = heading or family
    return f"""
    * {{ font-family: "{family}"; font-size: 10pt; color: {TEXT}; }}
    QMainWindow, QDialog, QWidget#Page, QScrollArea#Page > QWidget > QWidget {{ background: {BG}; }}
    QScrollArea {{ border: none; background: transparent; }}
    QToolTip {{ background: {TEXT}; color: white; border: none; padding: 6px 8px; border-radius: 6px; }}

    /* ---------- chữ */
    QLabel[role="h1"] {{ font-family: "{heading}"; font-size: 22pt; font-weight: 700; color: {FOREST}; }}
    QLabel[role="h2"] {{ font-family: "{heading}"; font-size: 14pt; font-weight: 700; color: {FOREST}; }}
    QLabel[role="h3"] {{ font-size: 11pt; font-weight: 700; }}
    QLabel[role="muted"] {{ color: {MUTED}; }}
    QLabel[role="caption"] {{ color: {MUTED}; font-size: 9pt; }}
    QLabel[role="overline"] {{ color: {FAINT}; font-size: 8pt; font-weight: 700; letter-spacing: 1px; }}
    QLabel[role="stat"] {{ font-family: "{heading}"; font-size: 22pt; font-weight: 700; color: {FOREST}; }}
    QLabel[role="display"] {{ font-family: "{heading}"; font-size: 26pt; font-weight: 700; color: white; }}

    /* ---------- thẻ */
    QFrame#Card {{ background: {SURFACE}; border: 1px solid {BORDER}; border-radius: 16px; }}
    QFrame#Card QLabel {{ background: transparent; }}
    QFrame#Soft {{ background: {SURFACE_2}; border: 1px solid {BORDER}; border-radius: 12px; }}
    QFrame#Soft QLabel {{ background: transparent; }}
    QFrame#Banner {{ background: {PRIMARY_SOFT}; border: 1px solid {PRIMARY_LINE}; border-radius: 12px; }}
    QFrame#Banner QLabel {{ color: {FOREST}; background: transparent; }}
    QFrame#Hero {{ background: qlineargradient(x1:0, y1:0, x2:1, y2:1, stop:0 #1F3520, stop:0.6 {FOREST},
                                               stop:1 #4E6B45); border-radius: 20px; }}
    QFrame#Hero QLabel {{ color: #F4F6EE; background: transparent; }}
    QFrame#Hero QLabel[role="caption"] {{ color: #C9D0B6; }}
    QFrame#HeroChip {{ background: rgba(255,255,255,0.10); border: 1px solid rgba(255,255,255,0.18);
                       border-radius: 12px; }}
    QFrame#Tile {{ background: {SURFACE}; border: 1px solid {BORDER}; border-radius: 14px; }}
    QFrame#Tile[earned="true"] {{ background: {GOLD_SOFT}; border: 1px solid #EAD9A6; }}
    QFrame#Tile QLabel {{ background: transparent; }}

    /* ---------- nút */
    QPushButton {{ background: {SURFACE}; border: 1px solid {BORDER}; border-radius: 10px;
                   padding: 8px 16px; font-weight: 600; color: {TEXT_2}; }}
    QPushButton:hover {{ background: {SURFACE_2}; border-color: {PRIMARY_LINE}; }}
    QPushButton:pressed {{ background: #E8E6DC; }}
    QPushButton:disabled {{ color: {FAINT}; background: {SURFACE_2}; border-color: {BORDER}; }}
    QPushButton[kind="primary"] {{ background: {PRIMARY}; color: white; border: 1px solid {PRIMARY}; }}
    QPushButton[kind="primary"]:hover {{ background: {PRIMARY_HOVER}; border-color: {PRIMARY_HOVER}; }}
    QPushButton[kind="primary"]:disabled {{ background: {OLIVE}; border-color: {OLIVE}; color: white; }}
    QPushButton[kind="accent"] {{ background: {OLIVE}; color: {TEXT}; border: 1px solid {OLIVE}; }}
    QPushButton[kind="accent"]:hover {{ background: #B4BB97; }}
    QPushButton[kind="danger"] {{ color: {DANGER}; }}
    QPushButton[kind="danger"]:hover {{ background: {DANGER_SOFT}; border-color: #E2B5AA; }}
    QPushButton[kind="ghost"] {{ background: transparent; border: none; color: {PRIMARY}; padding: 6px 10px; }}
    QPushButton[kind="ghost"]:hover {{ background: {PRIMARY_SOFT}; }}
    QPushButton[size="lg"] {{ padding: 12px 28px; font-size: 11pt; border-radius: 12px; }}
    QPushButton[size="sm"] {{ padding: 5px 10px; font-size: 9pt; border-radius: 8px; }}

    /* ---------- nút chuyển ngôn ngữ */
    QFrame#Seg {{ background: rgba(0,0,0,0.06); border-radius: 10px; }}
    QFrame#Seg[dark="true"] {{ background: rgba(255,255,255,0.10); }}
    QPushButton[seg="true"] {{ background: transparent; border: none; border-radius: 8px; padding: 4px 10px;
                               font-size: 9pt; font-weight: 700; color: {MUTED}; min-width: 28px; }}
    QFrame#Seg[dark="true"] QPushButton[seg="true"] {{ color: #C9D0B6; }}
    QPushButton[seg="true"]:checked, QFrame#Seg[dark="true"] QPushButton[seg="true"]:checked {{
        background: {OLIVE}; color: {TEXT}; }}

    /* ---------- ô nhập */
    QLineEdit, QPlainTextEdit, QTextEdit, QComboBox, QSpinBox {{
        background: {SURFACE}; border: 1px solid #D5D2C4; border-radius: 10px; padding: 7px 10px;
        selection-background-color: #DCE3C8; selection-color: {TEXT}; }}
    QLineEdit:focus, QPlainTextEdit:focus, QTextEdit:focus, QComboBox:focus {{ border: 1px solid {PRIMARY}; }}
    QLineEdit:disabled, QPlainTextEdit:disabled, QComboBox:disabled {{ background: {SURFACE_2}; color: {FAINT}; }}
    QLineEdit[size="lg"] {{ padding: 10px 12px; font-size: 11pt; }}
    QComboBox::drop-down {{ border: none; width: 26px; }}
    QComboBox::down-arrow {{ image: url("{icons['chevron']}"); width: 12px; height: 12px; margin-right: 10px; }}
    QComboBox QAbstractItemView {{ background: {SURFACE}; border: 1px solid {BORDER}; border-radius: 8px;
                                   padding: 4px; selection-background-color: {PRIMARY_SOFT}; selection-color: {TEXT}; outline: 0; }}

    /* ---------- bảng / danh sách */
    QTableWidget, QListWidget {{ background: {SURFACE}; border: 1px solid {BORDER}; border-radius: 14px;
                                 gridline-color: transparent; outline: 0; alternate-background-color: {SURFACE_2}; }}
    QTableWidget::item {{ padding: 6px 8px; border-bottom: 1px solid #F0EEE6; }}
    QTableWidget::item:selected, QListWidget::item:selected {{ background: {PRIMARY_SOFT}; color: {TEXT}; }}
    QListWidget {{ padding: 4px; }}
    QListWidget::item {{ padding: 8px 10px; border-radius: 8px; margin: 1px 0; }}
    QListWidget::item:hover {{ background: {SURFACE_2}; }}
    QHeaderView {{ background: transparent; }}
    QHeaderView::section {{ background: {SURFACE_2}; color: {MUTED}; font-size: 9pt; font-weight: 700;
                            border: none; border-bottom: 1px solid {BORDER}; padding: 9px 8px; }}
    QHeaderView::section:first {{ border-top-left-radius: 14px; }}
    QHeaderView::section:last {{ border-top-right-radius: 14px; }}

    /* ---------- thanh cuộn */
    QScrollBar:vertical {{ background: transparent; width: 10px; margin: 4px 2px; }}
    QScrollBar::handle:vertical {{ background: #D5D2C4; border-radius: 3px; min-height: 36px; }}
    QScrollBar::handle:vertical:hover {{ background: {FAINT}; }}
    QScrollBar:horizontal {{ background: transparent; height: 10px; margin: 2px 4px; }}
    QScrollBar::handle:horizontal {{ background: #D5D2C4; border-radius: 3px; min-width: 36px; }}
    QScrollBar::add-line, QScrollBar::sub-line {{ width: 0; height: 0; }}
    QScrollBar::add-page, QScrollBar::sub-page {{ background: transparent; }}

    /* ---------- thanh tiến độ */
    QProgressBar {{ background: #E8E6DC; border: none; border-radius: 4px; max-height: 8px; min-height: 8px; }}
    QProgressBar::chunk {{ background: {SUCCESS}; border-radius: 4px; }}
    QProgressBar[tone="warn"]::chunk {{ background: #C98A1B; }}
    QProgressBar[tone="danger"]::chunk {{ background: {DANGER}; }}
    QProgressBar[tone="xp"] {{ background: rgba(255,255,255,0.18); max-height: 10px; min-height: 10px;
                              border-radius: 5px; }}
    QProgressBar[tone="xp"]::chunk {{ background: qlineargradient(x1:0, y1:0, x2:1, y2:0, stop:0 {OLIVE},
                                      stop:1 #E3C66B); border-radius: 5px; }}

    QCheckBox {{ spacing: 8px; }}
    QCheckBox::indicator {{ width: 16px; height: 16px; border: 1px solid #D5D2C4; border-radius: 5px; background: white; }}
    QCheckBox::indicator:checked {{ background: {PRIMARY}; border-color: {PRIMARY}; image: url("{icons['check']}"); }}

    /* ---------- thanh bên (sidebar) */
    QWidget#Sidebar {{ background: qlineargradient(x1:0, y1:0, x2:0, y2:1, stop:0 {SIDEBAR}, stop:1 #26402A); }}
    QWidget#Sidebar QLabel {{ color: #F4F6EE; background: transparent; }}
    QWidget#Sidebar QLabel[role="caption"] {{ color: #AEB59A; }}
    QWidget#Sidebar QLabel[role="overline"] {{ color: #8A9277; }}
    QPushButton[nav="true"] {{ background: transparent; color: #D3D8C4; border: none; text-align: left;
                               padding: 10px 14px; border-radius: 10px; font-weight: 600; }}
    QPushButton[nav="true"]:hover {{ background: {SIDEBAR_2}; color: white; }}
    QPushButton[nav="true"]:checked {{ background: {OLIVE}; color: {TEXT}; }}
    QPushButton[kind="side"] {{ background: transparent; color: #C9D0B6; border: 1px solid #3E5A3A;
                                padding: 6px 10px; border-radius: 8px; font-size: 9pt; }}
    QPushButton[kind="side"]:hover {{ background: {SIDEBAR_2}; color: white; }}

    /* ---------- trang đăng nhập */
    QWidget#Login {{ background: qlineargradient(x1:0, y1:0, x2:1, y2:1, stop:0 #16271A, stop:0.5 {FOREST},
                                                 stop:1 {SAGE}); }}

    /* ---------- thanh làm bài */
    QWidget#BarHeader {{ background: {SIDEBAR}; }}
    QWidget#BarHeader QLabel {{ color: #F4F6EE; background: transparent; }}
    QPushButton[tab="true"] {{ background: {SIDEBAR_2}; color: #D3D8C4; border: none; border-radius: 8px;
                               padding: 6px 14px; font-size: 9pt; font-weight: 700; }}
    QPushButton[tab="true"]:hover {{ color: white; }}
    QPushButton[tab="true"]:checked {{ background: {OLIVE}; color: {TEXT}; }}
    QWidget#BarFooter {{ background: {SURFACE}; border-top: 1px solid {BORDER}; }}
    QPushButton[toggle="mark"]:checked {{ background: {WARN_SOFT}; color: {WARN}; border-color: #E8C98A; }}
    QPushButton[toggle="done"]:checked {{ background: {SUCCESS_SOFT}; color: {SUCCESS}; border-color: #A9CFA0; }}

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


def shadow(w: QWidget, blur=30, y=8, alpha=22) -> QWidget:
    eff = QGraphicsDropShadowEffect(w)
    eff.setBlurRadius(blur)
    eff.setOffset(0, y)
    eff.setColor(QColor(31, 53, 32, alpha))
    w.setGraphicsEffect(eff)
    return w


class Card(QFrame):
    def __init__(self, parent=None, padding=20, spacing=10, raised=False, horizontal=False, name="Card"):
        super().__init__(parent)
        self.setObjectName(name)
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


def stat_card(title: str, value: str, hint: str = "", accent: str = PRIMARY, icon: str = "") -> Card:
    c = Card(padding=18, spacing=4)
    top = QHBoxLayout()
    if icon:
        ic = QLabel(icon)
        ic.setStyleSheet("font-size: 13pt; background: transparent;")
        top.addWidget(ic)
    else:
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


def stars(n: int, size_pt=13) -> QLabel:
    """0–3 ngôi sao (vàng = đạt được)."""
    lb = QLabel("".join(f"<span style='color:{GOLD if i < n else '#D9D6C8'}'>★</span>" for i in range(3)))
    lb.setStyleSheet(f"font-size:{size_pt}pt; background: transparent; letter-spacing: 1px;")
    lb.setTextFormat(Qt.RichText)
    return lb


def lang_switch(current: str, on_change, dark: bool = False) -> QFrame:
    """Nút gạt VI | EN."""
    box = QFrame()
    box.setObjectName("Seg")
    box.setProperty("dark", dark)
    lay = QHBoxLayout(box)
    lay.setContentsMargins(3, 3, 3, 3)
    lay.setSpacing(2)
    group = QButtonGroup(box)
    for code, text in (("vi", "VI"), ("en", "EN")):
        b = QPushButton(text)
        b.setProperty("seg", True)
        b.setCheckable(True)
        b.setChecked(code == current)
        b.setCursor(Qt.PointingHandCursor)
        b.setToolTip("Tiếng Việt" if code == "vi" else "English")
        b.clicked.connect(lambda _=False, c=code: c != current and on_change(c))
        group.addButton(b)
        lay.addWidget(b)
    box.group = group
    return box


def badge_tile(icon: str, name: str, desc: str, earned: bool) -> QFrame:
    """Ô huy hiệu: sáng màu khi đã đạt, mờ khi chưa."""
    f = QFrame()
    f.setObjectName("Tile")
    f.setProperty("earned", earned)
    f.setToolTip(desc)
    f.setMinimumWidth(96)
    lay = QVBoxLayout(f)
    lay.setContentsMargins(12, 10, 12, 10)
    lay.setSpacing(2)
    ic = QLabel(icon if earned else "🔒")
    ic.setAlignment(Qt.AlignCenter)
    ic.setStyleSheet(f"font-size: 20pt; background: transparent; {'' if earned else 'color: #B9B6A8;'}")
    lay.addWidget(ic)
    nm = label(name, "h3")
    nm.setAlignment(Qt.AlignCenter)
    nm.setWordWrap(True)
    if not earned:
        nm.setStyleSheet(f"color: {MUTED};")
    lay.addWidget(nm)
    ds = label(desc, "caption", wrap=True)
    ds.setAlignment(Qt.AlignCenter)
    lay.addWidget(ds)
    return f


class LevelBadge(QWidget):
    """Huy hiệu cấp độ: vòng tròn tiến độ XP quanh số cấp."""

    def __init__(self, level: int, ratio: float, size=96, parent=None):
        super().__init__(parent)
        self.level, self.ratio = level, max(0.0, min(ratio, 1.0))
        self.setFixedSize(size, size)

    def paintEvent(self, _):
        p = QPainter(self)
        p.setRenderHint(QPainter.Antialiasing)
        w = 8
        rect = QRectF(w / 2 + 2, w / 2 + 2, self.width() - w - 4, self.height() - w - 4)
        p.setPen(QPen(QColor(255, 255, 255, 50), w, Qt.SolidLine, Qt.RoundCap))
        p.drawEllipse(rect)
        p.setPen(QPen(QColor("#E3C66B"), w, Qt.SolidLine, Qt.RoundCap))
        p.drawArc(rect, 90 * 16, int(-360 * 16 * self.ratio))
        p.setPen(QColor("white"))
        f = QFont(self.font())
        f.setPointSize(8)
        f.setBold(True)
        p.setFont(f)
        p.drawText(QRectF(0, self.height() * 0.22, self.width(), 16), Qt.AlignCenter, tr("CẤP"))
        f.setPointSize(22)
        p.setFont(f)
        p.drawText(QRectF(0, 6, self.width(), self.height()), Qt.AlignCenter, str(self.level))


class ScoreRing(QWidget):
    """Vòng tròn điểm (0–1000)."""

    def __init__(self, score: int, passed: bool, size=168, parent=None):
        super().__init__(parent)
        self.score, self.passed = score, passed
        self.setFixedSize(size, size)
        self.shown = 0
        self._anim = QTimer(self)
        self._anim.timeout.connect(self._step)
        self._anim.start(16)

    def _step(self):
        self.shown = min(self.score, self.shown + max(8, self.score // 45))
        if self.shown >= self.score:
            self._anim.stop()
        self.update()

    def paintEvent(self, _):
        p = QPainter(self)
        p.setRenderHint(QPainter.Antialiasing)
        w = 14
        rect = QRectF(w / 2 + 2, w / 2 + 2, self.width() - w - 4, self.height() - w - 4)
        p.setPen(QPen(QColor("#E8E6DC"), w, Qt.SolidLine, Qt.RoundCap))
        p.drawEllipse(rect)
        if self.shown:
            p.setPen(QPen(QColor(SUCCESS if self.passed else DANGER), w, Qt.SolidLine, Qt.RoundCap))
            p.drawArc(rect, 90 * 16, int(-360 * 16 * min(self.shown, 1000) / 1000))
        p.setPen(QColor(FOREST))
        f = QFont(pick_heading_family())
        f.setPointSize(28 if self.score < 1000 else 24)
        f.setBold(True)
        p.setFont(f)
        p.drawText(QRectF(0, -8, self.width(), self.height()), Qt.AlignCenter, str(self.shown))
        f = QFont(self.font())
        f.setPointSize(9)
        p.setFont(f)
        p.setPen(QColor(MUTED))
        p.drawText(QRectF(0, 32, self.width(), self.height()), Qt.AlignCenter, tr("/ 1000 điểm"))


class Confetti(QWidget):
    """Pháo giấy rơi phủ lên widget cha trong vài giây (khi đạt bài)."""

    COLORS = [OLIVE, "#E3C66B", "#C98A1B", "#6E9B5C", "#F4F6EE", SAGE, "#D98E5F"]

    def __init__(self, parent: QWidget, count=140, seconds=3.8):
        super().__init__(parent)
        self.setAttribute(Qt.WA_TransparentForMouseEvents)
        self.setAttribute(Qt.WA_NoSystemBackground)
        self.setGeometry(parent.rect())
        rnd = random.Random()
        w = max(parent.width(), 400)
        self.bits = [dict(x=rnd.uniform(0, w), y=rnd.uniform(-parent.height() * 0.6, -10),
                          vx=rnd.uniform(-1.2, 1.2), vy=rnd.uniform(2.5, 6.0), a=rnd.uniform(0, 360),
                          va=rnd.uniform(-8, 8), s=rnd.uniform(5, 10), c=rnd.choice(self.COLORS),
                          round=rnd.random() < 0.3)
                     for _ in range(count)]
        self.frames = int(seconds * 60)
        self.timer = QTimer(self)
        self.timer.timeout.connect(self._tick)
        self.timer.start(16)
        self.show()
        self.raise_()

    def _tick(self):
        self.frames -= 1
        for b in self.bits:
            b["x"] += b["vx"] + math.sin(b["y"] / 40) * 0.6
            b["y"] += b["vy"]
            b["a"] += b["va"]
        if self.frames <= 0:
            self.timer.stop()
            self.deleteLater()
            return
        self.update()

    def paintEvent(self, _):
        p = QPainter(self)
        p.setRenderHint(QPainter.Antialiasing)
        fade = min(1.0, self.frames / 40)
        for b in self.bits:
            c = QColor(b["c"])
            c.setAlphaF(0.95 * fade)
            p.save()
            p.translate(QPointF(b["x"], b["y"]))
            p.rotate(b["a"])
            p.setBrush(c)
            p.setPen(Qt.NoPen)
            if b["round"]:
                p.drawEllipse(QRectF(-b["s"] / 2, -b["s"] / 2, b["s"], b["s"]))
            else:
                path = QPainterPath()
                path.addRoundedRect(QRectF(-b["s"] / 2, -b["s"] / 4, b["s"], b["s"] / 2), 1.5, 1.5)
                p.drawPath(path)
            p.restore()


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
    names = {QMessageBox.Yes: tr("Đồng ý"), QMessageBox.No: tr("Không"), QMessageBox.Ok: "OK",
             QMessageBox.Cancel: tr("Hủy")}
    for std, txt in names.items():
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


def empty_state(text: str, icon: str = "🌿") -> QWidget:
    w = QFrame()
    w.setObjectName("Soft")
    lay = QVBoxLayout(w)
    lay.setContentsMargins(24, 28, 24, 28)
    ic = QLabel(icon)
    ic.setAlignment(Qt.AlignCenter)
    ic.setStyleSheet("font-size: 26pt; background: transparent;")
    lay.addWidget(ic)
    lb = label(text, "muted", wrap=True)
    lb.setAlignment(Qt.AlignCenter)
    lay.addWidget(lb)
    return w


def fmt_time(seconds) -> str:
    if seconds is None:
        return "—"
    m, s = divmod(max(int(seconds), 0), 60)
    return f"{m:02d}:{s:02d}"
