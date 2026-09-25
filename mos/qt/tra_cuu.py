"""Trang "Tra từ khóa": thuật ngữ MOS + câu trong đề + slide bài giảng có chứa từ khóa."""
from __future__ import annotations

from dataclasses import replace

from PySide6.QtCore import QTimer
from PySide6.QtWidgets import QComboBox, QGridLayout, QHBoxLayout, QLineEdit, QScrollArea, QVBoxLayout, QWidget

from .. import core, tai_lieu, tu_khoa
from ..custom import load_custom_exams, merge_exams
from ..exams import ALL_EXAMS
from ..i18n import is_en, pick, tr
from . import theme as T
from .app import _clear_page, esc, exam_name, page_body
from .theme import GOLD_SOFT, PRIMARY_SOFT, Card, button, chip, label

SUGGEST = ["Watermark", "mục lục", "Freeze Panes", "SmartArt", "Track Changes", "Alt Text", "section break",
           "Mail Merge", "IF", "Transitions", "caption", "hyperlink"]


class SearchPage(QScrollArea):
    def __init__(self, shell, query: str = ""):
        super().__init__()
        self.shell = shell
        custom, _ = load_custom_exams()
        self.exams = merge_exams(ALL_EXAMS, custom)
        self.lessons = tai_lieu.list_lessons()
        body, lay = page_body()
        lay.addWidget(T.page_header(tr("Tra từ khóa"), tr("Gõ tên lệnh, tính năng hoặc từ tiếng Việt (không cần dấu) – "
                                                          "xem giải thích, câu luyện liên quan và slide bài giảng.")))
        row = QHBoxLayout()
        self.box = QLineEdit(query)
        self.box.setProperty("size", "lg")
        self.box.setPlaceholderText(tr("vd: watermark, muc luc, freeze panes, bieu do…"))
        self.box.setClearButtonEnabled(True)
        row.addWidget(self.box, 1)
        self.subject = QComboBox()
        self.subject.addItem(tr("Mọi môn"), None)
        for code in ("WORD", "EXCEL", "POWERPOINT"):
            self.subject.addItem(T.brand(code)[3], code)
        self.subject.setMinimumWidth(150)
        row.addWidget(self.subject)
        lay.addLayout(row)

        chips = QGridLayout()
        chips.setSpacing(6)
        chips.addWidget(label(tr("Gợi ý:"), "caption"), 0, 0)
        for i, word in enumerate(SUGGEST):
            chips.addWidget(button(word, lambda w=word: self.box.setText(w), None, "sm"), i // 6, 1 + i % 6)
        chips.setColumnStretch(7, 1)
        lay.addLayout(chips)

        self.results = QWidget()
        self.results_lay = QVBoxLayout(self.results)
        self.results_lay.setContentsMargins(0, 0, 0, 0)
        self.results_lay.setSpacing(12)
        lay.addWidget(self.results)
        lay.addStretch()
        _clear_page(self, body)

        self.timer = QTimer(self)
        self.timer.setSingleShot(True)
        self.timer.setInterval(250)
        self.timer.timeout.connect(self.run)
        self.box.textChanged.connect(lambda *_: self.timer.start())
        self.subject.currentIndexChanged.connect(lambda *_: self.run())
        self.run()

    def showEvent(self, e):
        super().showEvent(e)
        self.box.setFocus()

    # ------------------------------------------------------------ kết quả
    def _clear(self):
        while self.results_lay.count():
            item = self.results_lay.takeAt(0)
            if item.widget():
                item.widget().deleteLater()

    def run(self):
        self._clear()
        q = self.box.text().strip()
        if not q:
            self.results_lay.addWidget(T.empty_state(tr("Nhập từ khóa để bắt đầu tra cứu."), "🔎"))
            return
        subject = self.subject.currentData()
        terms = tu_khoa.search_terms(q, subject)
        exams = [e for e in self.exams if subject is None or e.code == subject]
        tasks = tu_khoa.search_tasks(q, exams)
        lessons = [x for x in self.lessons if subject is None or x.mon == subject]
        slides = tu_khoa.search_slides(q, lessons)
        if not (terms or tasks or slides):
            self.results_lay.addWidget(T.empty_state(tr("Không tìm thấy “{q}”. Thử từ khác hoặc tên lệnh tiếng Anh.")
                                                     .format(q=q), "🤔"))
            return
        if terms:
            self.results_lay.addWidget(label(tr("Thuật ngữ ({n})").format(n=len(terms)), "h2"))
            for t in terms[:12]:
                self.results_lay.addWidget(self._term_card(t))
        if slides:
            self.results_lay.addWidget(label(tr("Trong bài giảng ({n})").format(n=len(slides)), "h2"))
            for h in slides:
                self.results_lay.addWidget(self._slide_row(h))
        if tasks:
            self.results_lay.addWidget(label(tr("Câu luyện trong đề ({n})").format(n=len(tasks)), "h2"))
            for h in tasks:
                self.results_lay.addWidget(self._task_card(h))

    def _term_card(self, t) -> Card:
        c = Card(padding=16, spacing=6)
        top = QHBoxLayout()
        top.addWidget(label(t.term, "h3"))
        if t.subject != "ALL":
            top.addWidget(chip(T.brand(t.subject)[3], "white", T.brand(t.subject)[1]))
        top.addStretch()
        top.addWidget(chip("🧭 " + t.path, T.FOREST, PRIMARY_SOFT))
        c.lay.addLayout(top)
        main, other = (t.en, t.vi) if is_en() else (t.vi, t.en)
        c.lay.addWidget(label(main, wrap=True))
        c.lay.addWidget(label(other, "caption", wrap=True))
        return c

    def _slide_row(self, h) -> Card:
        c = Card(padding=14, spacing=12, horizontal=True)
        c.lay.addWidget(T.badge(h.lesson.mon, 30))
        col = QVBoxLayout()
        col.setSpacing(2)
        col.addWidget(label(tr("{lesson} · slide {n}").format(lesson=pick(h.lesson.ten, h.lesson.ten_en), n=h.number),
                            "h3"))
        col.addWidget(label(esc(h.snippet), "caption", wrap=True))
        c.lay.addLayout(col, 1)
        c.lay.addWidget(button(tr("Xem slide"), lambda: self.shell.go("lesson", lesson=h.lesson, start=h.number),
                               "ghost", "sm"))
        return c

    def _task_card(self, h) -> Card:
        c = Card(padding=14, spacing=6)
        top = QHBoxLayout()
        top.addWidget(T.badge(h.exam.code, 26))
        top.addWidget(label(f"{exam_name(h.exam)} · {pick(h.project.name, h.project.name_en)}", "caption"), 1)
        top.addWidget(button("▶ " + tr("Luyện câu này"), lambda: self.practice(h), "accent", "sm"))
        c.lay.addLayout(top)
        c.lay.addWidget(label(esc(pick(h.task.title, h.task.title_en)), wrap=True))
        hint = label(f"<b>{tr('Cách làm:')}</b> {esc(pick(h.task.hint, h.task.hint_en))}", wrap=True)
        hint.setStyleSheet(f"background:{GOLD_SOFT}; color:#5C4410; border-radius:10px; padding:8px 12px;")
        c.lay.addWidget(hint)
        return c

    def practice(self, h):
        """Luyện riêng một câu: dùng file của dự án chứa câu đó, chế độ luyện tập."""
        project = replace(h.project, tasks=[h.task])
        exam = core.Exam(code=h.exam.code, name=f"Luyện 1 câu – {h.project.name}",
                         name_en=f"Single task – {h.project.name_en or h.project.name}", projects=[project], minutes=0)
        self.shell.start_exam(exam, "training")

