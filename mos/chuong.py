"""Chương (nhóm kỹ năng) theo khung đề thi MOS, dùng cho chế độ "Luyện theo chương" và tài liệu học.

Word theo MO-110 (Word 365, 6 nhóm + phần Expert), Excel theo MO-200, PowerPoint theo MO-300.
"""
from __future__ import annotations

import random
from dataclasses import dataclass, replace

from .core import Exam, Project


@dataclass(frozen=True)
class Chapter:
    code: str       # WORD / EXCEL / POWERPOINT
    number: int
    name_vi: str
    name_en: str
    icon: str


_RAW = {
    "WORD": [
        (1, "📁", "Quản lý tài liệu", "Manage documents"),
        (2, "✍", "Chèn và định dạng chữ, đoạn, section", "Insert and format text, paragraphs and sections"),
        (3, "▦", "Bảng và danh sách", "Manage tables and lists"),
        (4, "🔗", "Tham chiếu", "Create and manage references"),
        (5, "🖼", "Đồ họa (ảnh, hình, SmartArt, biểu đồ)", "Insert and format graphic elements"),
        (6, "💬", "Cộng tác (comment, theo dõi thay đổi)", "Manage document collaboration"),
        (7, "⭐", "Nâng cao (Expert)", "Advanced (Expert)"),
    ],
    "EXCEL": [
        (1, "📁", "Quản lý trang tính và sổ tính", "Manage worksheets and workbooks"),
        (2, "▦", "Quản lý ô và vùng dữ liệu", "Manage data cells and ranges"),
        (3, "📋", "Bảng và dữ liệu bảng", "Manage tables and table data"),
        (4, "ƒ", "Công thức và hàm", "Perform operations by using formulas and functions"),
        (5, "📊", "Biểu đồ", "Manage charts"),
    ],
    "POWERPOINT": [
        (1, "📁", "Quản lý bản trình bày", "Manage presentations"),
        (2, "🗂", "Quản lý slide", "Manage slides"),
        (3, "✍", "Chèn và định dạng chữ, hình, ảnh", "Insert and format text, shapes and images"),
        (4, "📊", "Bảng, biểu đồ, SmartArt, 3D, media", "Insert tables, charts, SmartArt, 3D models and media"),
        (5, "✨", "Chuyển trang và hiệu ứng động", "Apply transitions and animations"),
    ],
}

CHAPTERS: dict[str, list[Chapter]] = {code: [Chapter(code, n, vi, en, icon) for n, icon, vi, en in rows]
                                      for code, rows in _RAW.items()}


def chapters(code: str) -> list[Chapter]:
    return CHAPTERS.get(code, [])


def chapter(code: str, number: int | None) -> Chapter | None:
    return next((c for c in chapters(code) if c.number == number), None)


def count_tasks(exams: list[Exam], code: str) -> dict[int, int]:
    """{số chương: số nhiệm vụ} của một môn."""
    out: dict[int, int] = {}
    for exam in exams:
        if exam.code != code:
            continue
        for p in exam.projects:
            for t in p.tasks:
                if t.chapter:
                    out[t.chapter] = out.get(t.chapter, 0) + 1
    return out


def practice_exam(exams: list[Exam], code: str, number: int, max_projects: int = 6,
                  rng: random.Random | None = None) -> Exam | None:
    """Tạo bài luyện gồm các nhiệm vụ thuộc chương `number` của môn `code`.

    Mỗi dự án chỉ giữ lại nhiệm vụ thuộc chương; lấy tối đa `max_projects` dự án có nhiều nhiệm vụ nhất
    (chọn ngẫu nhiên trong các dự án ngang nhau để mỗi lần luyện một khác). Tên file trùng nhau được đổi để không ghi đè trong thư mục làm bài.
    """
    ch = chapter(code, number)
    if ch is None:
        return None
    picked: list[Project] = []
    seen_titles: set[str] = set()
    for exam in exams:
        if exam.code != code:
            continue
        for p in exam.projects:
            tasks = [t for t in p.tasks if t.chapter == number and t.title not in seen_titles]
            if tasks:
                seen_titles.update(t.title for t in tasks)
                picked.append(replace(p, tasks=tasks))
    if not picked:
        return None
    (rng or random).shuffle(picked)
    picked.sort(key=lambda p: -len(p.tasks))
    picked = picked[:max_projects]
    used, projects = set(), []
    for i, p in enumerate(picked, start=1):
        stem, dot, ext = p.filename.rpartition(".")
        name, n = p.filename, 1
        while name.casefold() in used:
            n += 1
            name = f"{stem}_{n}{dot}{ext}"
        used.add(name.casefold())
        short = p.name.split("–")[-1].strip()
        short_en = (p.name_en or p.name).split("–")[-1].strip()
        projects.append(replace(p, filename=name, name=f"Dự án {i} – {short}",
                                name_en=f"Project {i} – {short_en}"))
    return Exam(code=code, name=f"Chương {number}: {ch.name_vi}", name_en=f"Chapter {number}: {ch.name_en}",
                projects=projects, minutes=0)
