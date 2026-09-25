"""Nạp đề tự soạn từ thư mục `de_thi/`.

Mỗi đề là một thư mục con gồm:
    de.json          – khai báo dự án, nhiệm vụ, luật chấm
    <file gốc>       – file .docx/.xlsx/.pptx soạn sẵn bằng Office
    dap_an/          – (tuỳ chọn) file đã làm đúng, để kiểm tra đề bằng kiem_tra_de.py

Khóa tùy chọn trong de.json:
    "rieng": true    – đề đứng riêng trên Trang chủ (không gộp vào bài thi có sẵn)
    "file_phu": [..] – file / thư mục phụ (ảnh, dữ liệu trộn thư…) chép cùng file làm bài
"""
from __future__ import annotations

import json
import re
import shutil
from dataclasses import replace
from pathlib import Path

from . import rules
from .core import DATA_DIR, Exam, Project, Task, app_dir

MON = {"WORD", "EXCEL", "POWERPOINT"}


def search_dirs() -> list[Path]:
    return [app_dir() / "de_thi", DATA_DIR / "de_thi"]


def _copier(src: Path, extras: tuple[Path, ...] = ()):
    def build(dest: Path) -> None:
        shutil.copyfile(src, dest)
        for extra in extras:
            target = dest.parent / extra.name
            if extra.is_dir():
                if not target.exists():
                    shutil.copytree(extra, target)
            elif extra.is_file() and not target.exists():
                shutil.copyfile(extra, target)
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
    extras = tuple(folder / str(x) for x in data.get("file_phu", []))
    for extra in extras:
        if not extra.exists():
            errors.append(f"{folder.name}: không thấy file phụ '{extra.name}'")
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
                                intro=p.get("mo_ta", ""), build=_copier(src, extras), tasks=tasks))
    if not projects:
        errors.append(f"{folder.name}: chưa có dự án nào trong \"du_an\"")
    if errors:
        return None, errors
    return Exam(code=mon, name=data.get("ten", folder.name), projects=projects,
                minutes=int(data.get("thoi_gian", 50)), standalone=bool(data.get("rieng"))), []


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
    Đề có "rieng": true (vd bộ đề nhập) đứng riêng, xếp sau các bài thi có sẵn.
    """
    merged = []
    for exam in builtin:
        projects = list(exam.projects) + [p for c in custom if c.code == exam.code and not c.standalone
                                          for p in c.projects]
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
    return merged + sorted((c for c in custom if c.standalone), key=lambda e: e.name)


# ====================================================================== dùng cho màn hình Soạn đề


def editable_dir() -> Path:
    """Nơi lưu đề soạn trong app (dùng chung khi cấu hình thư mục dữ liệu mạng)."""
    return DATA_DIR / "de_thi"


def list_exam_folders() -> list[Path]:
    seen, out = set(), []
    for base in search_dirs():
        if base.is_dir():
            for folder in sorted(p for p in base.iterdir() if (p / "de.json").is_file()):
                if folder.resolve() not in seen:
                    seen.add(folder.resolve())
                    out.append(folder)
    return out


def read_exam_json(folder: Path) -> dict:
    return json.loads((folder / "de.json").read_text(encoding="utf-8-sig"))


def save_exam_json(folder: Path, data: dict) -> None:
    folder.mkdir(parents=True, exist_ok=True)
    tmp = folder / "de.json.tmp"
    tmp.write_text(json.dumps(data, ensure_ascii=False, indent=2) + "\n", encoding="utf-8")
    tmp.replace(folder / "de.json")


def new_exam_folder(title: str) -> Path:
    """Tạo tên thư mục không dấu, không trùng, từ tên đề."""
    import unicodedata
    ascii_name = unicodedata.normalize("NFKD", title.replace("đ", "d").replace("Đ", "D"))
    ascii_name = "".join(c for c in ascii_name if not unicodedata.combining(c))
    slug = re.sub(r"[^A-Za-z0-9]+", "_", ascii_name).strip("_") or "De"
    folder, n = editable_dir() / slug, 1
    while folder.exists():
        n += 1
        folder = editable_dir() / f"{slug}_{n}"
    return folder


def check_exam(folder: Path, answers: dict[str, Path] | None = None) -> tuple[list[str], list[dict]]:
    """Chấm thử đề: file gốc (phải SAI) và file đáp án (phải ĐÚNG).

    answers: {tên file gốc: đường dẫn file đáp án}; mặc định lấy trong dap_an/.
    Trả về (lỗi khai báo, danh sách dòng kết quả).
    """
    from .core import run_check
    exam, errors = load_exam(folder)
    if errors:
        return errors, []
    rows = []
    for project in exam.projects:
        answer = (answers or {}).get(project.filename) or folder / "dap_an" / project.filename
        for i, task in enumerate(project.tasks, start=1):
            start_ok, _ = run_check(task, folder / project.filename)
            row = {"project": project.name, "file": project.filename, "index": i, "task": task.title,
                   "start_ok": start_ok, "answer_ok": None, "error": ""}
            if answer.is_file():
                row["answer_ok"], row["error"] = run_check(task, answer)
            rows.append(row)
    return [], rows
