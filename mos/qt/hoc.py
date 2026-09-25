"""Trang "Tài liệu học": danh sách bài giảng; trình xem slide / video / SCORM (không có nút tải / lưu);
nhập bài từ .pptx, video hoặc gói SCORM; nút "Thực hành ngay" mở bài luyện của chương."""
from __future__ import annotations

import json
from pathlib import Path

from PySide6.QtCore import QByteArray, QBuffer, QIODevice, QSize, Qt, QUrl
from PySide6.QtGui import QPixmap
from PySide6.QtWidgets import (QApplication, QComboBox, QDialog, QFileDialog, QFormLayout, QFrame, QGridLayout,
                               QHBoxLayout, QLabel, QLineEdit, QListWidget, QListWidgetItem, QScrollArea,
                               QSizePolicy, QSlider, QVBoxLayout, QWidget)

from .. import chuong, tai_lieu
from ..custom import load_custom_exams, merge_exams
from ..exams import ALL_EXAMS
from ..i18n import pick, tr
from . import theme as T
from .app import _banner, _clear_page, page_body
from .theme import PRIMARY_SOFT, ClickableCard, button, chip, label

SUBJECTS = ("WORD", "EXCEL", "POWERPOINT")


KIND_ICON = {"slides": "🖼", "video": "🎬", "scorm": "🧩"}


def kind_name(loai: str) -> str:
    return {"slides": tr("Slide"), "video": tr("Video"), "scorm": tr("Tương tác (SCORM)")}.get(loai, loai)


def lesson_name(lesson) -> str:
    return pick(lesson.ten, lesson.ten_en)


def practice_exam(lesson):
    """Bài luyện theo chương của bài giảng (None nếu bài không gắn chương hoặc chương chưa có câu)."""
    if not lesson.chuong:
        return None
    custom, _ = load_custom_exams()
    return chuong.practice_exam(merge_exams(ALL_EXAMS, custom), lesson.mon, lesson.chuong)


def practice_button(shell, lesson, text: str | None = None):
    """Nút "Thực hành ngay": học xong mở luôn bài luyện chương đó (None nếu không có câu luyện)."""
    if practice_exam(lesson) is None:
        return None
    return button(text or ("▶ " + tr("Thực hành ngay")),
                  lambda: _start_practice(shell, lesson), "accent")


def _start_practice(shell, lesson):
    exam = practice_exam(lesson)
    if exam is not None:
        shell.start_exam(exam, "chapter")


def open_viewer(shell, lesson, start: int = 1) -> QWidget:
    """Trình xem phù hợp với loại bài giảng."""
    if lesson.loai == "video":
        return VideoViewer(shell, lesson)
    if lesson.loai == "scorm":
        return ScormViewer(shell, lesson)
    return LessonViewer(shell, lesson, start)


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
            actions = [button(tr("Thêm bài giảng…"), self.import_lesson, "primary")]
        lay.addWidget(T.page_header(tr("Tài liệu học"),
                                    tr("Xem slide, video ngắn, bài tương tác theo từng chương rồi thực hành ngay. "
                                       "Tài liệu chỉ xem trong app, không tải về."),
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
                msg += " " + tr("Bấm “Thêm bài giảng…” để thêm slide PPTX, video hoặc gói SCORM.")
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
        top.addWidget(chip(KIND_ICON.get(lesson.loai, "") + " " + kind_name(lesson.loai), T.MUTED, T.IVORY))
        top.addStretch()
        seen = self.progress.get(lesson.id, 0)
        if lesson.count and seen >= lesson.count:
            top.addWidget(chip("✓ " + tr("Đã học"), T.SUCCESS, T.SUCCESS_SOFT))
        card.lay.addLayout(top)
        if lesson.loai == "slides":
            try:
                pix = QPixmap()
                pix.loadFromData(lesson.image(1))
                thumb = QLabel()
                thumb.setPixmap(pix.scaledToWidth(360, Qt.SmoothTransformation))
                thumb.setStyleSheet(f"border: 1px solid {T.BORDER}; border-radius: 8px;")
                card.lay.addWidget(thumb)
            except Exception:  # gói ảnh hỏng: vẫn hiện thẻ
                pass
        else:
            poster = QLabel(KIND_ICON.get(lesson.loai, "📖"))
            poster.setAlignment(Qt.AlignCenter)
            poster.setMinimumHeight(150)
            poster.setStyleSheet(f"background: {T.SIDEBAR}; color: {T.IVORY}; border-radius: 8px; font-size: 54px;")
            card.lay.addWidget(poster)
        card.lay.addWidget(label(lesson_name(lesson), "h3", wrap=True))
        if ch:
            card.lay.addWidget(label(pick(ch.name_vi, ch.name_en), "caption", wrap=True))
        card.lay.addWidget(T.progress(min(seen, lesson.count), max(lesson.count, 1)))
        foot = QHBoxLayout()
        if lesson.loai == "slides":
            done_text = tr("{n}/{total} slide").format(n=min(seen, lesson.count), total=lesson.count)
        else:
            done_text = tr("Đã học") if seen else tr("Chưa học")
            if lesson.loai == "scorm":
                score = tai_lieu.scorm_score(tai_lieu.load_scorm(self.shell.account.username, lesson.id))
                if score:
                    done_text += " · " + tr("Điểm {s}").format(s=score)
        foot.addWidget(label(done_text, "caption"))
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
        self.has_practice = practice_exam(lesson) is not None
        if self.has_practice:
            top.addSpacing(8)
            top.addWidget(practice_button(shell, lesson))
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
        self.next_btn = button(tr("Slide sau") + "  ›", self.next, "primary")
        foot.addWidget(label(tr("Dùng phím ← → để chuyển slide."), "caption"))
        foot.addStretch()
        foot.addWidget(self.prev_btn)
        foot.addWidget(self.next_btn)
        lay.addLayout(foot)

        self.index = 0
        self.show_slide(max(1, min(start, lesson.count)))

    def next(self):
        if self.index >= self.lesson.count and self.has_practice:
            _start_practice(self.shell, self.lesson)
        else:
            self.show_slide(self.index + 1)

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
        last = n >= self.lesson.count
        self.next_btn.setText(("▶ " + tr("Thực hành ngay")) if last and self.has_practice
                              else tr("Slide sau") + "  ›")
        self.next_btn.setEnabled(not last or self.has_practice)
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


# ====================================================================== khung chung cho video / SCORM


def _viewer_top(shell, lesson, lay: QVBoxLayout) -> QHBoxLayout:
    top = QHBoxLayout()
    top.addWidget(button("‹  " + tr("Tài liệu học"), lambda: shell.go("lessons", code=lesson.mon), "ghost"))
    top.addSpacing(8)
    top.addWidget(T.badge(lesson.mon, 28))
    top.addWidget(label(lesson_name(lesson), "h2"), 1)
    ch = chuong.chapter(lesson.mon, lesson.chuong)
    if ch:
        top.addWidget(chip(tr("Chương {n}").format(n=ch.number), T.FOREST, PRIMARY_SOFT))
    pb = practice_button(shell, lesson)
    if pb:
        top.addSpacing(8)
        top.addWidget(pb)
    lay.addLayout(top)
    return top


def _fmt_time(ms: int) -> str:
    s = max(0, ms // 1000)
    return f"{s // 60}:{s % 60:02d}"


# ====================================================================== video


class VideoViewer(QWidget):
    """Phát video ngay trong app: giải mã vào bộ nhớ, không có nút lưu / tải."""

    def __init__(self, shell, lesson):
        super().__init__()
        self.shell, self.lesson = shell, lesson
        self.setObjectName("Page")
        self.setAttribute(Qt.WA_StyledBackground, True)
        self.setContextMenuPolicy(Qt.NoContextMenu)
        self.setFocusPolicy(Qt.StrongFocus)
        lay = QVBoxLayout(self)
        lay.setContentsMargins(28, 20, 28, 20)
        lay.setSpacing(12)
        _viewer_top(shell, lesson, lay)
        self.player = None
        self.done = False
        try:
            from PySide6.QtMultimedia import QAudioOutput, QMediaPlayer
            from PySide6.QtMultimediaWidgets import QVideoWidget
        except Exception as exc:  # máy thiếu thành phần phát video của Qt
            lay.addWidget(_banner(tr("Máy chưa phát được video: {err}").format(err=exc), "bad"))
            lay.addStretch()
            return
        frame = QFrame()
        frame.setObjectName("VideoFrame")
        frame.setStyleSheet(f"QFrame#VideoFrame {{ background: {T.SIDEBAR}; border-radius: 14px; }}")
        fl = QVBoxLayout(frame)
        fl.setContentsMargins(10, 10, 10, 10)
        self.video = QVideoWidget()
        self.video.setContextMenuPolicy(Qt.NoContextMenu)
        self.video.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)
        fl.addWidget(self.video)
        lay.addWidget(frame, 1)
        if pick(lesson.mo_ta, lesson.mo_ta_en):
            lay.addWidget(label(pick(lesson.mo_ta, lesson.mo_ta_en), "muted", wrap=True))

        ctl = QHBoxLayout()
        self.play_btn = button("▶ " + tr("Phát"), self.toggle, "primary")
        ctl.addWidget(self.play_btn)
        self.pos_lbl = label("0:00 / 0:00", "caption")
        self.slider = QSlider(Qt.Horizontal)
        self.slider.sliderMoved.connect(lambda v: self.player.setPosition(v))
        ctl.addWidget(self.slider, 1)
        ctl.addWidget(self.pos_lbl)
        self.speed = QComboBox()
        for rate in (0.75, 1.0, 1.25, 1.5):
            self.speed.addItem(f"{rate:g}×", rate)
        self.speed.setCurrentIndex(1)
        self.speed.currentIndexChanged.connect(lambda *_: self.player.setPlaybackRate(self.speed.currentData()))
        ctl.addWidget(self.speed)
        lay.addLayout(ctl)
        lay.addWidget(label(tr("Phím cách: phát / dừng · ← →: lùi / tới 5 giây."), "caption"))

        self.end_banner = _banner("🎉  " + tr("Xem xong rồi! Bấm “Thực hành ngay” để luyện các câu của chương này."),
                                  "ok")
        self.end_banner.setVisible(False)
        lay.addWidget(self.end_banner)

        try:
            data = lesson.video()
        except Exception as exc:  # file video hỏng
            lay.addWidget(_banner(tr("Không đọc được video: {err}").format(err=exc), "bad"))
            return
        self._bytes = QByteArray(data)
        self._buffer = QBuffer(self._bytes, self)
        self._buffer.open(QIODevice.ReadOnly)
        self.audio = QAudioOutput(self)
        self.player = QMediaPlayer(self)
        self.player.setAudioOutput(self.audio)
        self.player.setVideoOutput(self.video)
        self.player.durationChanged.connect(self._duration)
        self.player.positionChanged.connect(self._position)
        self.player.playbackStateChanged.connect(self._state)
        self.player.mediaStatusChanged.connect(self._status)
        ext = json.loads((lesson.folder / tai_lieu.META).read_text(encoding="utf-8")).get("dinh_dang", ".mp4")
        self.player.setSourceDevice(self._buffer, QUrl("bai_giang" + ext))

    def toggle(self):
        if self.player is None:
            return
        from PySide6.QtMultimedia import QMediaPlayer
        if self.player.playbackState() == QMediaPlayer.PlayingState:
            self.player.pause()
        else:
            self.player.play()

    def _duration(self, ms):
        self.slider.setRange(0, ms)
        self._position(self.player.position())

    def _position(self, ms):
        if not self.slider.isSliderDown():
            self.slider.setValue(ms)
        self.pos_lbl.setText(f"{_fmt_time(ms)} / {_fmt_time(self.player.duration())}")
        if not self.done and self.player.duration() > 0 and ms >= self.player.duration() * 0.9:
            self._finish()

    def _state(self, state):
        from PySide6.QtMultimedia import QMediaPlayer
        playing = state == QMediaPlayer.PlayingState
        self.play_btn.setText(("⏸ " + tr("Dừng")) if playing else ("▶ " + tr("Phát")))

    def _status(self, status):
        from PySide6.QtMultimedia import QMediaPlayer
        if status == QMediaPlayer.EndOfMedia:
            self._finish()
        elif status == QMediaPlayer.InvalidMedia:
            self.pos_lbl.setText(tr("Không phát được video này."))

    def _finish(self):
        """Xem gần hết (≥ 90%) là tính đã học."""
        if self.done:
            return
        self.done = True
        tai_lieu.save_progress(self.shell.account.username, self.lesson.id, 1)
        self.end_banner.setVisible(practice_exam(self.lesson) is not None)
        T.Confetti(self.video, count=50, seconds=2.0)

    def keyPressEvent(self, e):
        if self.player is not None and e.key() == Qt.Key_Space:
            self.toggle()
        elif self.player is not None and e.key() in (Qt.Key_Left, Qt.Key_Right):
            step = 5000 if e.key() == Qt.Key_Right else -5000
            self.player.setPosition(max(0, self.player.position() + step))
        else:
            super().keyPressEvent(e)

    def showEvent(self, e):
        super().showEvent(e)
        self.setFocus()

    def hideEvent(self, e):
        if self.player is not None:
            self.player.pause()
        super().hideEvent(e)


# ====================================================================== SCORM


SCORM_PREFIX = "MOS_SCORM:"
_PROFILE = None

# Giả lập LMS (SCORM 1.2: window.API, SCORM 2004: window.API_1484_11) ngay trong trang;
# dữ liệu gửi về app qua console.log để lưu tiến độ / điểm.
_SCORM_JS = r"""
(function () {
  if (window.API && window.API_1484_11) return;
  var saved = __DATA__, user = __USER__, data = {}, err = "0", k;
  for (k in saved) data[k] = saved[k];
  var resume = !!(data["cmi.suspend_data"] || data["cmi.core.lesson_location"] || data["cmi.location"]);
  var defaults = {
    "cmi.core.student_id": user, "cmi.core.student_name": user, "cmi.core.lesson_status": "not attempted",
    "cmi.core.entry": resume ? "resume" : "ab-initio", "cmi.core.credit": "credit", "cmi.core.lesson_mode": "normal",
    "cmi.core.total_time": "0000:00:00", "cmi.launch_data": "", "cmi.suspend_data": "",
    "cmi.core.lesson_location": "", "cmi.core.score.raw": "",
    "cmi.learner_id": user, "cmi.learner_name": user, "cmi.completion_status": "unknown",
    "cmi.success_status": "unknown", "cmi.entry": resume ? "resume" : "ab-initio", "cmi.mode": "normal",
    "cmi.credit": "credit", "cmi.location": "", "cmi.total_time": "PT0S"
  };
  function send(kind) { try { console.log("__PREFIX__" + kind + ":" + JSON.stringify(data)); } catch (e) {} }
  function get(key) {
    err = "0";
    if (key in data) return String(data[key]);
    if (key in defaults) return defaults[key];
    if (/\._count$/.test(key)) return "0";
    if (/\._children$/.test(key)) return "";
    return "";
  }
  function set(key, value) { err = "0"; data[key] = String(value); send("set"); return "true"; }
  function ok() { err = "0"; return "true"; }
  function commit() { send("commit"); return "true"; }
  function finish() { send("finish"); return "true"; }
  function lastErr() { return err; }
  function empty() { return ""; }
  window.API = { LMSInitialize: ok, LMSFinish: finish, LMSGetValue: get, LMSSetValue: set, LMSCommit: commit,
                 LMSGetLastError: lastErr, LMSGetErrorString: empty, LMSGetDiagnostic: empty };
  window.API_1484_11 = { Initialize: ok, Terminate: finish, GetValue: get, SetValue: set, Commit: commit,
                         GetLastError: lastErr, GetErrorString: empty, GetDiagnostic: empty };
})();
"""


def scorm_script(saved: dict, user: str) -> str:
    return (_SCORM_JS.replace("__DATA__", json.dumps(saved, ensure_ascii=False))
            .replace("__USER__", json.dumps(user or "hoc_vien")).replace("__PREFIX__", SCORM_PREFIX))


def parse_scorm_message(message: str) -> dict | None:
    """Tin nhắn console "MOS_SCORM:<loại>:<json>" → dict cmi (None nếu không phải của SCORM)."""
    if not message.startswith(SCORM_PREFIX):
        return None
    _, _, rest = message[len(SCORM_PREFIX):].partition(":")
    try:
        data = json.loads(rest)
    except ValueError:
        return None
    return data if isinstance(data, dict) else None


class ScormViewer(QWidget):
    """Chạy gói SCORM (HTML tương tác) trong trình duyệt nhúng; không cho tải / lưu, ghi nhận hoàn thành + điểm."""

    def __init__(self, shell, lesson):
        super().__init__()
        self.shell, self.lesson = shell, lesson
        self.setObjectName("Page")
        self.setAttribute(Qt.WA_StyledBackground, True)
        lay = QVBoxLayout(self)
        lay.setContentsMargins(28, 20, 28, 20)
        lay.setSpacing(12)
        top = _viewer_top(shell, lesson, lay)
        self.user = shell.account.username
        self.cmi = tai_lieu.load_scorm(self.user, lesson.id)
        self.status = label("", "muted")
        top.insertWidget(top.count() - (2 if practice_exam(lesson) else 0), self.status)
        self._show_status()
        try:
            from PySide6.QtWebEngineCore import (QWebEnginePage, QWebEngineProfile, QWebEngineScript,
                                                 QWebEngineSettings)
            from PySide6.QtWebEngineWidgets import QWebEngineView
        except Exception as exc:  # máy thiếu QtWebEngine
            lay.addWidget(_banner(tr("Máy chưa chạy được bài tương tác: {err}").format(err=exc), "bad"))
            lay.addStretch()
            return

        viewer = self

        class Page(QWebEnginePage):
            def javaScriptConsoleMessage(self, level, message, line, source):  # noqa: N802 (tên hàm của Qt)
                data = parse_scorm_message(message)
                if data is not None:
                    viewer.on_scorm(message[len(SCORM_PREFIX):].split(":", 1)[0], data)

        global _PROFILE
        if _PROFILE is None:     # một profile ẩn danh dùng chung: không lưu bộ nhớ đệm ra đĩa, chặn tải về
            _PROFILE = QWebEngineProfile(QApplication.instance())
            _PROFILE.downloadRequested.connect(lambda item: item.cancel())
        self.profile = _PROFILE
        for old in self.profile.scripts().find("mos-scorm-api"):
            self.profile.scripts().remove(old)
        script = QWebEngineScript()
        script.setName("mos-scorm-api")
        script.setSourceCode(scorm_script(self.cmi, self.user))
        script.setInjectionPoint(QWebEngineScript.DocumentCreation)
        script.setWorldId(QWebEngineScript.MainWorld)
        script.setRunsOnSubFrames(True)
        self.profile.scripts().insert(script)
        self.view = QWebEngineView(self)
        self.page = Page(self.profile, self.view)
        st = self.page.settings()
        st.setAttribute(QWebEngineSettings.LocalContentCanAccessFileUrls, True)
        st.setAttribute(QWebEngineSettings.LocalContentCanAccessRemoteUrls, True)
        self.view.setPage(self.page)
        self.view.setContextMenuPolicy(Qt.NoContextMenu)
        lay.addWidget(self.view, 1)
        self.view.setUrl(QUrl.fromLocalFile(str(lesson.launch_path.resolve())))

    def on_scorm(self, kind: str, data: dict):
        was_done = tai_lieu.scorm_done(self.cmi)
        self.cmi = data
        if kind in ("commit", "finish") or (tai_lieu.scorm_done(data) and not was_done):
            tai_lieu.save_scorm(self.user, self.lesson.id, data)
        self._show_status()
        if tai_lieu.scorm_done(data) and not was_done:
            T.Confetti(self, count=60, seconds=2.2)

    def _show_status(self):
        parts = [("✓ " + tr("Đã hoàn thành")) if tai_lieu.scorm_done(self.cmi) else tr("Chưa hoàn thành")]
        score = tai_lieu.scorm_score(self.cmi)
        if score:
            parts.append(tr("Điểm {s}").format(s=score))
        self.status.setText(" · ".join(parts))

    def hideEvent(self, e):
        if self.cmi:
            tai_lieu.save_scorm(self.user, self.lesson.id, self.cmi)
        super().hideEvent(e)


# ====================================================================== nhập bài giảng


class ImportLessonDialog(QDialog):
    def __init__(self, parent):
        super().__init__(parent)
        self.setWindowTitle(tr("Thêm bài giảng"))
        self.setMinimumWidth(560)
        self.path: Path | None = None
        lay = QVBoxLayout(self)
        lay.setContentsMargins(26, 22, 26, 22)
        lay.setSpacing(10)
        lay.addWidget(label(tr("Thêm bài giảng"), "h2"))
        lay.addWidget(label(tr("Chọn một trong ba loại:\n"
                               "• Slide PowerPoint (.pptx): chuyển thành ảnh (máy có PowerPoint thì giống hệt bản gốc).\n"
                               "• Video ngắn (.mp4, .webm…): phát trong app.\n"
                               "• Bài tương tác SCORM (.zip xuất từ iSpring, Articulate, Adobe Captivate, "
                               "H5P…): chạy trong app, ghi nhận hoàn thành và điểm.\n"
                               "Học viên chỉ xem trong app, không có nút tải về. Gắn chương để có nút "
                               "“Thực hành ngay”."), "muted", wrap=True))
        form = QFormLayout()
        form.setSpacing(10)
        pick_row = QHBoxLayout()
        self.file_lbl = label(tr("Chưa chọn file"), "muted")
        pick_row.addWidget(self.file_lbl, 1)
        pick_row.addWidget(button(tr("Chọn file…"), self.choose, size="sm"))
        form.addRow(label(tr("File bài giảng"), "caption"), pick_row)
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
        self.desc = QLineEdit()
        self.desc.setPlaceholderText(tr("Mô tả ngắn (tuỳ chọn)"))
        form.addRow(label(tr("Mô tả"), "caption"), self.desc)
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
        video = " ".join("*" + e for e in tai_lieu.VIDEO_EXT)
        filters = ";;".join([tr("Bài giảng") + f" (*.pptx *.zip {video})", "PowerPoint (*.pptx)",
                             "Video (" + video + ")", "SCORM (*.zip)"])
        path, _ = QFileDialog.getOpenFileName(self, tr("Chọn bài giảng"), "", filters)
        if path:
            self.path = Path(path)
            self.file_lbl.setText(self.path.name)
            if not self.name.text().strip():
                self.name.setText(self.path.stem.replace("_", " "))

    def do_import(self):
        if self.path is None:
            self.err.setText(tr("Hãy chọn file bài giảng."))
            return
        if not self.name.text().strip():
            self.err.setText(tr("Hãy đặt tên bài."))
            return
        QApplication.setOverrideCursor(Qt.WaitCursor)
        try:
            tai_lieu.import_file(self.path, self.name.text().strip(), self.subject.currentData(),
                                 self.chapter.currentData(), ten_en=self.name_en.text().strip(),
                                 mo_ta=self.desc.text().strip())
        except Exception as exc:  # file hỏng / không đọc được
            QApplication.restoreOverrideCursor()
            self.err.setText(tr("Không nhập được: {err}").format(err=exc))
            return
        QApplication.restoreOverrideCursor()
        self.accept()
