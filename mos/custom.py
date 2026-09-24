"""Nạp đề tự soạn từ thư mục `de_thi/`.

Mỗi đề là một thư mục con gồm:
    de.json          – khai báo dự án, nhiệm vụ, luật chấm
    <file gốc>       – file .docx/.xlsx/.pptx soạn sẵn bằng Office
    dap_an/          – (tuỳ chọn) file đã làm đúng, để kiểm tra đề bằng kiem_tra_de.py
"""
from __future__ import annotations

import json
import re
import shutil
import sys
from dataclasses import replace
from pathlib import Path

from . import rules
from .core import APP_DIR, Exam, Project, Task

MON = {"WORD", "EXCEL", "POWERPOINT"}


def app_dir() -> Path:
    """Thư mục chứa main.py (hoặc file .exe khi đã đóng gói)."""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).parent
    return Path(__file__).resolve().parents[1]


def search_dirs() -> list[Path]:
    return [app_dir() / "de_thi", APP_DIR / "de_thi"]


def _copier(src: Path):
    def build(dest: Path) -> None:
        shutil.copyfile(src, dest)
    return build


def load_exam(folder: Path) -> tuple[Exam | None, list[str]]:
    """Đọc một thư mục đề. Trả về (Exam, danh sách lỗi)."""
    errors: list[str] = []
    try:
        data = json.loads((folder / "de.json").read_text(encoding="utf-8-sig"))
    except FileNotFoundError:
        return None, [f"{folder.name}: thiếu file de.json"]
    except ValueError as exc:
        return None, [f"{folder.name}/de.json: sai cú pháp JSON – {exc}"]

    mon = str(data.get("mon", "")).upper()
    if mon not in MON:
        errors.append(f"{folder.name}: \"mon\" phải là WORD, EXCEL hoặc POWERPOINT")
    projects = []
    for p_i, p in enumerate(data.get("du_an", []), start=1):
        where = f"{folder.name} › dự án {p_i}"
        src = folder / str(p.get("file", ""))
        if not p.get("file") or not src.is_file():
            errors.append(f"{where}: không thấy file gốc '{p.get('file')}'")
        tasks = []
        for t_i, t in enumerate(p.get("nhiem_vu", []), start=1):
            specs = t.get("cham")
            if not t.get("yeu_cau") or not specs:
                errors.append(f"{where} › nhiệm vụ {t_i}: cần có \"yeu_cau\" và \"cham\"")
                continue
            bad = [e for e in (rules.validate(s) if isinstance(s, dict) else "luật phải là một đối tượng {...}"
                               for s in rules._as_list(specs)) if e]
            if bad:
                errors.extend(f"{where} › nhiệm vụ {t_i}: {e}" for e in bad)
                continue
            tasks.append(Task(t["yeu_cau"], t.get("goi_y", "(Không có gợi ý)"), rules.make_check(specs)))
        projects.append(Project(name=p.get("ten", f"Dự án {p_i}"), filename=src.name,
                                intro=p.get("mo_ta", ""), build=_copier(src), tasks=tasks))
    if not projects:
        errors.append(f"{folder.name}: chưa có dự án nào trong \"du_an\"")
    if errors:
        return None, errors
    return Exam(code=mon, name=data.get("ten", folder.name), projects=projects,
                minutes=int(data.get("thoi_gian", 50))), []


def load_custom_exams() -> tuple[list[Exam], list[str]]:
    exams, errors = [], []
    for base in search_dirs():
        if not base.is_dir():
            continue
        for folder in sorted(p for p in base.iterdir() if p.is_dir()):
            exam, errs = load_exam(folder)
            errors.extend(errs)
            if exam:
                exams.append(exam)
    return exams, errors


def merge_exams(builtin: list[Exam], custom: list[Exam]) -> list[Exam]:
    """Gộp dự án của đề tự soạn vào đề có sẵn cùng môn (Word/Excel/PowerPoint).

    Dự án được đánh số lại liên tục ("Dự án 3 – …") và tên file trùng được
    đổi (vd DoanhSo_2.xlsx) để không ghi đè lên nhau trong thư mục làm bài.
    """
    merged = []
    for exam in builtin:
        projects = list(exam.projects) + [p for c in custom if c.code == exam.code for p in c.projects]
        used: set[str] = set()
        renamed = []
        for i, p in enumerate(projects, start=1):
            short = re.sub(r"^Dự án\s*\d+\s*[–-]\s*", "", p.name).strip() or p.name
            stem, dot, ext = p.filename.rpartition(".")
            filename, n = p.filename, 1
            while filename.casefold() in used:
                n += 1
                filename = f"{stem}_{n}{dot}{ext}"
            used.add(filename.casefold())
            renamed.append(replace(p, name=f"Dự án {i} – {short}", filename=filename))
        merged.append(replace(exam, projects=renamed))
    return merged
