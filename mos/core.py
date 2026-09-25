"""Phần lõi: định nghĩa Đề thi / Dự án / Nhiệm vụ, chấm điểm và lưu lịch sử.

Cách hoạt động giống GMetrix / bài thi MOS thật:
  1. Mỗi đề thi (Exam) gồm nhiều dự án (Project).
  2. Mỗi dự án có 1 file Office khởi đầu (do hàm `build` tạo ra) và
     nhiều nhiệm vụ (Task) người học phải làm trên file đó.
  3. Người học mở file bằng Word/Excel/PowerPoint thật, làm bài, lưu lại.
  4. Khi nộp bài, mỗi Task có hàm `check(path)` đọc file đã lưu để xem
     nhiệm vụ đã được làm đúng chưa.
"""
from __future__ import annotations

import json
import os
import random
import shutil
import sys
import traceback
from dataclasses import dataclass, field
from datetime import datetime
from pathlib import Path
from typing import Callable

PASS_SCORE = 700
MAX_SCORE = 1000

APP_DIR = Path.home() / "MOS_Practice"     # thư mục làm bài trên từng máy


def app_dir() -> Path:
    """Thư mục chứa main.py (hoặc file .exe khi đã đóng gói)."""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).parent
    return Path(__file__).resolve().parents[1]


def _data_dir() -> Path:
    """Nơi lưu tài khoản, kết quả, đề tự soạn.

    Mặc định là APP_DIR. Phòng máy dùng chung dữ liệu: tạo file cau_hinh.json
    cạnh main.py / .exe với nội dung {"thu_muc_du_lieu": "\\\\MAYCHU\\MOS"}.
    """
    try:
        cfg = json.loads((app_dir() / "cau_hinh.json").read_text(encoding="utf-8-sig"))
        if cfg.get("thu_muc_du_lieu"):
            return Path(cfg["thu_muc_du_lieu"])
    except (OSError, ValueError):
        pass
    return APP_DIR


DATA_DIR = _data_dir()
HISTORY_FILE = DATA_DIR / "history.json"


@dataclass
class Task:
    title: str                      # Yêu cầu hiển thị cho người học
    hint: str                       # Hướng dẫn thao tác (hiện ở chế độ luyện tập / khi xem kết quả)
    check: Callable[[Path], bool | None]   # Hàm chấm: True nếu đúng; None = không chấm tự động được
    title_en: str = ""              # bản tiếng Anh (để trống = dùng bản tiếng Việt)
    hint_en: str = ""
    chapter: int | None = None      # chương / nhóm kỹ năng (mos/chuong.py)


@dataclass
class Project:
    name: str
    filename: str                   # Tên file làm bài, vd "DoanhSo.xlsx"
    intro: str                      # Mô tả tình huống
    build: Callable[[Path], None]   # Hàm tạo file khởi đầu
    tasks: list[Task]
    name_en: str = ""
    intro_en: str = ""


@dataclass
class Exam:
    code: str                       # vd "EXCEL"
    name: str                       # vd "Microsoft Excel (MO-200)"
    projects: list[Project]
    minutes: int = 50
    standalone: bool = False        # đề riêng (bộ đề nhập), không gộp vào bài thi có sẵn
    name_en: str = ""


@dataclass
class TaskResult:
    project: str
    task: str
    hint: str
    correct: bool | None            # None = tự kiểm tra (không tính điểm)
    error: str = ""
    task_en: str = ""
    hint_en: str = ""
    project_en: str = ""


@dataclass
class Session:
    """Một lượt làm bài: thư mục làm việc + trạng thái đánh dấu."""
    exam: Exam
    mode: str                       # "training", "chapter" (luyện theo chương) hoặc "testing"
    workdir: Path
    marked: set = field(default_factory=set)   # {(project_idx, task_idx)}
    done: set = field(default_factory=set)
    user: str | None = None                    # tên đăng nhập của người làm bài

    def file_of(self, project_idx: int) -> Path:
        return self.workdir / self.exam.projects[project_idx].filename


def new_session(exam: Exam, mode: str, base: Path | None = None, user: str | None = None) -> Session:
    """Tạo thư mục làm bài mới và sinh toàn bộ file khởi đầu."""
    base = base or APP_DIR / "work"
    stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
    workdir = base / f"{exam.code}_{stamp}_{random.randint(100, 999)}"
    workdir.mkdir(parents=True, exist_ok=True)
    for project in exam.projects:
        project.build(workdir / project.filename)
    return Session(exam=exam, mode=mode, workdir=workdir, user=user)


def reset_project(session: Session, project_idx: int) -> None:
    """Làm lại một dự án: tạo lại file khởi đầu (ghi đè bài đang làm)."""
    path = session.file_of(project_idx)
    if path.exists():
        backup = path.with_name(path.stem + "_cu" + path.suffix)
        shutil.copyfile(path, backup)
    session.exam.projects[project_idx].build(path)


def run_check(task: Task, path: Path) -> tuple[bool | None, str]:
    """Chạy hàm chấm một cách an toàn: lỗi đọc file = sai, kèm thông báo.
    Trả về None nếu nhiệm vụ không chấm tự động được (học viên tự kiểm tra)."""
    try:
        ok = task.check(path)
        return (None if ok is None else bool(ok)), ""
    except Exception as exc:  # file hỏng, đang bị khóa, thiếu sheet...
        return False, f"{type(exc).__name__}: {exc}"


def check_project(session: Session, project_idx: int) -> list[TaskResult]:
    project = session.exam.projects[project_idx]
    path = session.file_of(project_idx)
    results = []
    for task in project.tasks:
        ok, err = run_check(task, path)
        results.append(TaskResult(project.name, task.title, task.hint, ok, err,
                                  task.title_en, task.hint_en, project.name_en))
    return results


def grade(session: Session) -> dict:
    results: list[TaskResult] = []
    for i in range(len(session.exam.projects)):
        results.extend(check_project(session, i))
    total = sum(r.correct is not None for r in results)
    correct = sum(r.correct is True for r in results)
    score = round(MAX_SCORE * correct / total) if total else 0
    return {
        "exam": session.exam.name,
        "exam_en": session.exam.name_en,
        "code": session.exam.code,
        "user": session.user,
        "mode": session.mode,
        "time": datetime.now().strftime("%Y-%m-%d %H:%M"),
        "correct": correct,
        "total": total,
        "score": score,
        "passed": score >= PASS_SCORE,
        "results": results,
    }


def file_is_open(path: Path) -> bool:
    """Office tạo file khóa '~$...' cạnh file đang mở."""
    return any(p.name.startswith("~$") and p.suffix == path.suffix
               for p in path.parent.iterdir())


def open_in_office(path: Path) -> None:
    if os.name == "nt":
        os.startfile(path)  # type: ignore[attr-defined]
    elif shutil.which("open"):
        os.system(f'open "{path}"')
    else:
        os.system(f'xdg-open "{path}" >/dev/null 2>&1 &')


def load_history(user: str | None = None) -> list[dict]:
    """Toàn bộ lịch sử, hoặc chỉ của tài khoản `user`."""
    try:
        history = json.loads(HISTORY_FILE.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return []
    return history if user is None else [h for h in history if h.get("user") == user]


def save_history(report: dict) -> None:
    history = load_history()
    entry = {k: v for k, v in report.items() if k != "results"}
    entry["wrong"] = [f"{r.project}: {r.task}" for r in report["results"] if r.correct is False]
    history.append(entry)
    HISTORY_FILE.parent.mkdir(parents=True, exist_ok=True)
    HISTORY_FILE.write_text(json.dumps(history, ensure_ascii=False, indent=2), encoding="utf-8")


def format_exception() -> str:
    return traceback.format_exc(limit=2)
