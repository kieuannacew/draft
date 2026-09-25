"""Yếu tố "vui học": điểm kinh nghiệm (XP), cấp độ, chuỗi ngày học, huy hiệu, sao, mẹo mỗi ngày.

Mọi thứ tính lại từ lịch sử làm bài (history.json) nên không cần lưu thêm dữ liệu.
"""
from __future__ import annotations

from dataclasses import dataclass
from datetime import date, datetime, timedelta

from .core import PASS_SCORE

XP_PER_LEVEL = 400
LEVELS = [  # (tên Việt, tên Anh) theo cấp 1, 2, 3…
    ("Mầm non", "Seedling"), ("Chồi xanh", "Sprout"), ("Cây non", "Sapling"), ("Cây vững", "Young Tree"),
    ("Tán rộng", "Canopy"), ("Cổ thụ", "Old Oak"), ("Rừng già", "Ancient Forest"),
]


def attempt_xp(entry: dict) -> int:
    """XP của một lần nộp bài: điểm/10, +50 nếu đạt, +100 nếu tuyệt đối; thi thử x1.2."""
    xp = int(entry.get("score", 0)) // 10
    if entry.get("passed"):
        xp += 50
    if entry.get("score", 0) >= 1000:
        xp += 100
    if entry.get("mode") == "testing":
        xp = round(xp * 1.2)
    return xp


def total_xp(history: list[dict]) -> int:
    return sum(attempt_xp(h) for h in history)


@dataclass
class Level:
    number: int
    name_vi: str
    name_en: str
    xp_in_level: int
    xp_for_next: int

    @property
    def ratio(self) -> float:
        return self.xp_in_level / self.xp_for_next if self.xp_for_next else 1.0


def level(xp: int) -> Level:
    n = xp // XP_PER_LEVEL
    vi, en = LEVELS[min(n, len(LEVELS) - 1)]
    return Level(n + 1, vi, en, xp % XP_PER_LEVEL, XP_PER_LEVEL)


def _day(entry: dict) -> date | None:
    try:
        return datetime.strptime(entry["time"][:10], "%Y-%m-%d").date()
    except (KeyError, ValueError):
        return None


def streak(history: list[dict], today: date | None = None) -> int:
    """Số ngày liên tiếp có làm bài, tính tới hôm nay (hoặc hôm qua nếu hôm nay chưa làm)."""
    today = today or date.today()
    days = {d for d in (_day(h) for h in history) if d}
    if not days:
        return 0
    cur = today if today in days else today - timedelta(days=1)
    n = 0
    while cur in days:
        n += 1
        cur -= timedelta(days=1)
    return n


def stars(best: int | None) -> int:
    """0–3 sao theo điểm cao nhất của một đề."""
    if best is None:
        return 0
    return 3 if best >= 900 else 2 if best >= PASS_SCORE else 1 if best >= 400 else 0


def _code(entry: dict) -> str:
    if entry.get("code"):
        return entry["code"]
    name = entry.get("exam", "").casefold()
    for code, key in (("POWERPOINT", "powerpoint"), ("EXCEL", "excel"), ("WORD", "word")):
        if key in name:
            return code
    return ""


@dataclass
class Badge:
    key: str
    icon: str
    name_vi: str
    name_en: str
    desc_vi: str
    desc_en: str
    earned: bool


BADGES = [
    ("first", "🌱", "Khởi đầu", "First step", "Nộp bài lần đầu tiên", "Submit your first attempt"),
    ("pass", "🏅", "Vượt ngưỡng", "Passed", "Đạt từ 700 điểm", "Score 700 or more"),
    ("perfect", "💯", "Hoàn hảo", "Perfect", "Đạt 1000 điểm tuyệt đối", "Score a perfect 1000"),
    ("streak3", "🔥", "Chăm chỉ", "On fire", "Học 3 ngày liên tiếp", "Study 3 days in a row"),
    ("streak7", "🌳", "Bền bỉ", "Unstoppable", "Học 7 ngày liên tiếp", "Study 7 days in a row"),
    ("trio", "🎯", "Đa năng", "All-rounder", "Làm bài cả Word, Excel và PowerPoint",
     "Try Word, Excel and PowerPoint"),
    ("ten", "📚", "Mọt sách", "Bookworm", "Nộp bài 10 lần", "Submit 10 attempts"),
    ("master5", "🏆", "Chinh phục", "Conqueror", "Đạt 5 đề khác nhau", "Pass 5 different exams"),
    ("speed", "⚡", "Tốc độ", "Speedster", "Thi thử đạt trong dưới 25 phút",
     "Pass a timed test in under 25 minutes"),
]


def _earned(key: str, history: list[dict], today: date | None) -> bool:
    if key == "first":
        return bool(history)
    if key == "pass":
        return any(h.get("passed") for h in history)
    if key == "perfect":
        return any(h.get("score", 0) >= 1000 for h in history)
    if key in ("streak3", "streak7"):
        return _best_streak(history) >= (3 if key == "streak3" else 7)
    if key == "trio":
        return {"WORD", "EXCEL", "POWERPOINT"} <= {_code(h) for h in history}
    if key == "ten":
        return len(history) >= 10
    if key == "master5":
        return len({h.get("exam") for h in history if h.get("passed")}) >= 5
    if key == "speed":
        return any(h.get("passed") and h.get("mode") == "testing" and h.get("duration")
                   and h["duration"] < 25 * 60 for h in history)
    return False


def _best_streak(history: list[dict]) -> int:
    days = sorted({d for d in (_day(h) for h in history) if d})
    best = run = 0
    prev = None
    for d in days:
        run = run + 1 if prev and d - prev == timedelta(days=1) else 1
        best = max(best, run)
        prev = d
    return best


def badges(history: list[dict], today: date | None = None) -> list[Badge]:
    return [Badge(k, icon, nv, ne, dv, de, _earned(k, history, today)) for k, icon, nv, ne, dv, de in BADGES]


def new_badges(before: list[dict], after: list[dict]) -> list[Badge]:
    """Huy hiệu vừa mở khóa sau lần nộp bài mới nhất."""
    old = {b.key for b in badges(before) if b.earned}
    return [b for b in badges(after) if b.earned and b.key not in old]


TIPS = [
    ("Luôn bấm Ctrl + S sau mỗi dự án – app chỉ chấm phần đã lưu.",
     "Always press Ctrl + S after each project – only saved work is graded."),
    ("Đọc kỹ tên mục (section) trong đề: thao tác đúng nhưng sai chỗ vẫn bị tính sai.",
     "Read the section name carefully: the right action in the wrong place is still wrong."),
    ("Thi thật không có nút Undo cho cả dự án – dùng Ctrl + Z ngay khi lỡ tay.",
     "The real exam can't reset a project for you – use Ctrl + Z right after a mistake."),
    ("Đề nói “Accept all default settings” nghĩa là chỉ bấm OK, đừng đổi tùy chọn khác.",
     "“Accept all default settings” means just click OK – don't change other options."),
    ("Bí thì bấm Đánh dấu rồi làm câu khác, quay lại sau – đừng mất quá 3 phút cho một câu.",
     "Stuck? Mark the task, move on and come back – don't spend over 3 minutes on one task."),
    ("Mỗi ngày làm một đề ngắn giúp nhớ lâu hơn làm dồn nhiều đề một lúc.",
     "One short exam a day sticks better than cramming many at once."),
    ("Di chuột lên các ô màu để xem tên màu (vd Blue, Accent 1, Lighter 80%) trước khi chọn.",
     "Hover over color swatches to see their names (e.g. Blue, Accent 1, Lighter 80%) before choosing."),
    ("Học cả tên lệnh tiếng Anh: đề thi và giao diện Office trong phòng thi đều là tiếng Anh.",
     "Learn the English command names: the exam and Office in the test room are in English."),
]


def tip_of_day(today: date | None = None) -> tuple[str, str]:
    today = today or date.today()
    return TIPS[today.toordinal() % len(TIPS)]
