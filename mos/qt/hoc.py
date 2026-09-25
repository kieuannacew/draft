"""Trang "Tài liệu học": danh sách bài giảng, trình xem slide (không có nút tải / lưu) và nhập bài từ .pptx."""
from __future__ import annotations

from pathlib import Path

from PySide6.QtCore import QSize, Qt
from PySide6.QtGui import QPixmap
from PySide6.QtWidgets import (QApplication, QComboBox, QDialog, QFileDialog, QFormLayout, QGridLayout,
                               QHBoxLayout, QLabel, QLineEdit, QListWidget, QListWidgetItem, QScrollArea,
                               QSizePolicy, QVBoxLayout, QWidget)

from .. import chuong, tai_lieu
from ..i18n import pick, tr
from . import theme as T
from .app import _banner, _clear_page, page_body
from .theme import PRIMARY_SOFT, ClickableCard, button, chip, label

SUBJECTS = ("WORD", "EXCEL", "POWERPOINT")


def lesson_name(lesson) -> str:
    return pick(lesson.ten, lesson.ten_en)


# ====================================================================== danh sách bài giảng


class LessonsPage(QScrollArea):
    def __init__(self, shell, code: str | None = None):
        super().__init__()
        self.shell = shell
        self.code = code
        self.lessons = tai_lieu.list_lessons()
        self.progress = tai_lieu.load_progress(shell.account.username)
        body, lay = page_body()
        actions = []
        if shell.account.is_admin:
            actions = [button(tr("Nhập bài giảng PPTX…"), self.import_lesson, "primary")]
        lay.addWidget(T.page_header(tr("Tài liệu học"),
                                    tr("Xem bài giảng ngay trong app theo từng chương. Tài liệu chỉ xem, không tải về."),
                                    actions))

        # --- lọc theo môn
        row = QHBoxLayout()
        row.setSpacing(8)
        for code in (None,) + SUBJECTS:
            text = tr("Tất cả") if code is None else T.brand(code)[3]
            b = button(text, lambda c=code: self.shell.go("lessons", code=c), "primary" if code == self.code
                       else None, "sm")
            row.addWidget(b)
        row.addStretch()
        lay.addLayout(row)

        shown = [x for x in self.lessons if self.code is None or x.mon == self.code]
        if not shown:
            msg = tr("Chưa có bài giảng nào.")
            if shell.account.is_admin:
                msg += " " + tr("Bấm “Nhập bài giảng PPTX…” để thêm.")
            lay.addWidget(T.empty_state(msg, "📖"))
        for code in SUBJECTS:
            group = [x for x in shown if x.mon == code]
            if not group:
                continue
            head = QHBoxLayout()
            head.addWidget(T.badge(code, 30))
            head.addWidget(label(T.brand(code)[3], "h2"))
            head.addStretch()
            lay.addLayout(head)
            grid = QGridLayout()
            grid.setSpacing(14)
            for i, lesson in enumerate(group):
                grid.addWidget(self._card(lesson), i // 3, i % 3)
            for c in range(3):
                grid.setColumnStretch(c, 1)
            lay.addLayout(grid)
        lay.addStretch()
        _clear_page(self, body)

    def _card(self, lesson) -> ClickableCard:
        card = ClickableCard(padding=16, spacing=6)
        top = QHBoxLayout()
        ch = chuong.chapter(lesson.mon, lesson.chuong)
        top.addWidget(chip(tr("Chương {n}").format(n=lesson.chuong) if ch else tr("Chung"), T.FOREST, PRIMARY_SOFT))
        top.addStretch()
        seen = self.progress.get(lesson.id, 0)
        if lesson.count and seen >= lesson.count:
            top.addWidget(chip("✓ " + tr("Đã học"), T.SUCCESS, T.SUCCESS_SOFT))
        card.lay.addLayout(top)
        try:
            pix = QPixmap()
            pix.loadFromData(lesson.image(1))
            thumb = QLabel()
            thumb.setPixmap(pix.scaledToWidth(360, Qt.SmoothTransformation))
            thumb.setStyleSheet(f"border: 1px solid {T.BORDER}; border-radius: 8px;")
            thumb.setScaledContents(False)
            card.lay.addWidget(thumb)
        except Exception:  # gói ảnh hỏng: vẫn hiện thẻ
            pass
        card.lay.addWidget(label(lesson_name(lesson), "h3", wrap=True))
        if ch:
            card.lay.addWidget(label(pick(ch.name_vi, ch.name_en), "caption", wrap=True))
        card.lay.addWidget(T.progress(min(seen, lesson.count), max(lesson.count, 1)))
        foot = QHBoxLayout()
        foot.addWidget(label(tr("{n}/{total} slide").format(n=min(seen, lesson.count), total=lesson.count),
                             "caption"))
        foot.addStretch()
        if self.shell.account.is_admin:
            foot.addWidget(button(tr("Xóa"), lambda l=lesson: self.delete_lesson(l), "danger", "sm"))
        card.lay.addLayout(foot)
        card.clicked.connect(lambda l=lesson: self.shell.go("lesson", lesson=l,
                                                            start=min(max(seen, 1), l.count)))
        return card

    def import_lesson(self):
        dlg = ImportLessonDialog(self)
        if dlg.exec():
            self.shell.go("lessons", code=self.code)

    def delete_lesson(self, lesson):
        if T.confirm(self, tr("Xóa bài giảng"), tr("Xóa bài giảng “{name}”?").format(name=lesson_name(lesson))):
            tai_lieu.delete_lesson(lesson)
            self.shell.go("lessons", code=self.code)


# ====================================================================== trình xem slide


class SlideView(QLabel):
    """Ảnh slide co giãn theo khung, giữ tỉ lệ. Không có menu chuột phải / kéo thả."""

    def __init__(self):
        super().__init__()
        self.setAlignment(Qt.AlignCenter)
        self.setMinimumSize(QSize(480, 270))
        self.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        self.setContextMenuPolicy(Qt.NoContextMenu)
        self._pix = QPixmap()
        self.setStyleSheet("background: #1F3520; border-radius: 14px;")

    def set_image(self, data: bytes):
        self._pix = QPixmap()
        self._pix.loadFromData(data)
        self._rescale()

    def resizeEvent(self, e):
        super().resizeEvent(e)
        self._rescale()

    def _rescale(self):
        if not self._pix.isNull():
            self.setPixmap(self._pix.scaled(self.size() - QSize(24, 24), Qt.KeepAspectRatio,
                                            Qt.SmoothTransformation))


class LessonViewer(QWidget):
    def __init__(self, shell, lesson, start: int = 1):
        super().__init__()
        self.shell, self.lesson = shell, lesson
        self.setObjectName("Page")
        self.setAttribute(Qt.WA_StyledBackground, True)
        self.setFocusPolicy(Qt.StrongFocus)
        self.setContextMenuPolicy(Qt.NoContextMenu)
        lay = QVBoxLayout(self)
        lay.setContentsMargins(28, 20, 28, 20)
        lay.setSpacing(12)

        top = QHBoxLayout()
        top.addWidget(button("‹  " + tr("Tài liệu học"), lambda: shell.go("lessons", code=lesson.mon), "ghost"))
        top.addSpacing(8)
        top.addWidget(T.badge(lesson.mon, 28))
        name = label(lesson_name(lesson), "h2")
        top.addWidget(name, 1)
        self.counter = label("", "muted")
        top.addWidget(self.counter)
        lay.addLayout(top)

        mid = QHBoxLayout()
        mid.setSpacing(14)
        self.view = SlideView()
        mid.addWidget(self.view, 1)
        self.thumbs = QListWidget()
        self.thumbs.setIconSize(QSize(160, 90))
        self.thumbs.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        self.thumbs.setFixedWidth(206)
        self.thumbs.setSpacing(4)
        self.thumbs.setContextMenuPolicy(Qt.NoContextMenu)
        for i in range(1, lesson.count + 1):
            item = QListWidgetItem(f"{i}")
            try:
                pix = QPixmap()
                pix.loadFromData(lesson.image(i))
                item.setIcon(pix.scaled(160, 90, Qt.KeepAspectRatio, Qt.SmoothTransformation))
            except Exception:  # ảnh hỏng: chỉ hiện số
                pass
            self.thumbs.addItem(item)
        self.thumbs.currentRowChanged.connect(lambda r: r >= 0 and self.show_slide(r + 1))
        mid.addWidget(self.thumbs)
        lay.addLayout(mid, 1)

        self.notes = _banner("", "gold")
        self.notes_label = self.notes.findChild(QLabel)
        lay.addWidget(self.notes)

        foot = QHBoxLayout()
        self.prev_btn = button("‹  " + tr("Slide trước"), lambda: self.show_slide(self.index - 1))
        self.next_btn = button(tr("Slide sau") + "  ›", lambda: self.show_slide(self.index + 1), "primary")
        foot.addWidget(label(tr("Dùng phím ← → để chuyển slide."), "caption"))
        foot.addStretch()
        foot.addWidget(self.prev_btn)
        foot.addWidget(self.next_btn)
        lay.addLayout(foot)

        self.index = 0
        self.show_slide(max(1, min(start, lesson.count)))

    def show_slide(self, n: int):
        if not (1 <= n <= self.lesson.count) or n == self.index:
            return
        self.index = n
        try:
            self.view.set_image(self.lesson.image(n))
        except Exception as exc:  # gói ảnh hỏng
            self.view.setText(tr("Không đọc được slide: {err}").format(err=exc))
        self.counter.setText(tr("Slide {n}/{total}").format(n=n, total=self.lesson.count))
        self.prev_btn.setEnabled(n > 1)
        self.next_btn.setEnabled(n < self.lesson.count)
        notes = self.lesson.slides[n - 1].get("notes", "") if n - 1 < len(self.lesson.slides) else ""
        self.notes.setVisible(bool(notes))
        self.notes_label.setText("📝  " + notes)
        self.thumbs.blockSignals(True)
        self.thumbs.setCurrentRow(n - 1)
        self.thumbs.blockSignals(False)
        tai_lieu.save_progress(self.shell.account.username, self.lesson.id, n)
        if n == self.lesson.count and n > 1:
            T.Confetti(self.view, count=60, seconds=2.2)

    def keyPressEvent(self, e):
        if e.key() in (Qt.Key_Right, Qt.Key_Down, Qt.Key_PageDown, Qt.Key_Space):
            self.show_slide(self.index + 1)
        elif e.key() in (Qt.Key_Left, Qt.Key_Up, Qt.Key_PageUp):
            self.show_slide(self.index - 1)
        else:
            super().keyPressEvent(e)

    def showEvent(self, e):
        super().showEvent(e)
        self.setFocus()


# ====================================================================== nhập bài giảng


class ImportLessonDialog(QDialog):
    def __init__(self, parent):
        super().__init__(parent)
        self.setWindowTitle(tr("Nhập bài giảng PPTX"))
        self.setMinimumWidth(560)
        self.path: Path | None = None
        lay = QVBoxLayout(self)
        lay.setContentsMargins(26, 22, 26, 22)
        lay.setSpacing(10)
        lay.addWidget(label(tr("Nhập bài giảng PPTX"), "h2"))
        lay.addWidget(label(tr("Trên máy có PowerPoint, slide được chuyển thành ảnh giống hệt bản gốc; nếu không, app "
                               "tự dựng ảnh đơn giản. File .pptx gốc không được chép vào app, học viên chỉ xem được "
                               "trong app."), "muted", wrap=True))
        form = QFormLayout()
        form.setSpacing(10)
        pick_row = QHBoxLayout()
        self.file_lbl = label(tr("Chưa chọn file"), "muted")
        pick_row.addWidget(self.file_lbl, 1)
        pick_row.addWidget(button(tr("Chọn file…"), self.choose, size="sm"))
        form.addRow(label(tr("File PPTX"), "caption"), pick_row)
        self.name = QLineEdit()
        self.name.setPlaceholderText(tr("vd: Word – Chương 3: Bảng và danh sách"))
        form.addRow(label(tr("Tên bài"), "caption"), self.name)
        self.name_en = QLineEdit()
        self.name_en.setPlaceholderText(tr("Tên tiếng Anh (tuỳ chọn)"))
        form.addRow(label(tr("Tên tiếng Anh"), "caption"), self.name_en)
        self.subject = QComboBox()
        for code in SUBJECTS:
            self.subject.addItem(T.brand(code)[3], code)
        self.subject.currentIndexChanged.connect(self._fill_chapters)
        form.addRow(label(tr("Môn"), "caption"), self.subject)
        self.chapter = QComboBox()
        form.addRow(label(tr("Chương"), "caption"), self.chapter)
        lay.addLayout(form)
        self._fill_chapters()
        self.err = label("", wrap=True)
        self.err.setStyleSheet(f"color:{T.DANGER};")
        lay.addWidget(self.err)
        row = QHBoxLayout()
        row.addStretch()
        row.addWidget(button(tr("Hủy"), self.reject))
        row.addWidget(button(tr("Nhập bài"), self.do_import, "primary"))
        lay.addLayout(row)

    def _fill_chapters(self):
        self.chapter.clear()
        self.chapter.addItem(tr("(chung, không theo chương)"), None)
        for ch in chuong.chapters(self.subject.currentData()):
            self.chapter.addItem(f"{ch.number}. {pick(ch.name_vi, ch.name_en)}", ch.number)

    def choose(self):
        path, _ = QFileDialog.getOpenFileName(self, tr("Chọn bài giảng"), "", "PowerPoint (*.pptx)")
        if path:
            self.path = Path(path)
            self.file_lbl.setText(self.path.name)
            if not self.name.text().strip():
                self.name.setText(self.path.stem.replace("_", " "))

    def do_import(self):
        if self.path is None:
            self.err.setText(tr("Hãy chọn file .pptx."))
            return
        if not self.name.text().strip():
            self.err.setText(tr("Hãy đặt tên bài."))
            return
        QApplication.setOverrideCursor(Qt.WaitCursor)
        try:
            tai_lieu.import_pptx(self.path, self.name.text().strip(), self.subject.currentData(),
                                 self.chapter.currentData(), ten_en=self.name_en.text().strip())
        except Exception as exc:  # file hỏng / không đọc được
            QApplication.restoreOverrideCursor()
            self.err.setText(tr("Không nhập được: {err}").format(err=exc))
            return
        QApplication.restoreOverrideCursor()
        self.accept()
