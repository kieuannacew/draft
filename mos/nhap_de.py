"""Nhập bộ đề có sẵn (vd bộ đề MOS Word 365 "Đề thực tế" sinh tự động) vào app.

Bộ đề nguồn có dạng:
    <Bộ đề>/De_01/de_01.json          danh sách project: file, theme, tasks[id, text, steps]
    <Bộ đề>/De_01/Files/*.docx        file thực hành + file phụ (ảnh .png, dữ liệu .xlsx)
    <Bộ đề>/3D Models/                (tùy chọn) mô hình 3D cho câu "3D Models"

Mỗi De_xx được chuyển thành một đề tự soạn `de_thi/<tên>/de.json` đứng riêng trên
Trang chủ ("rieng": true). Câu Word được chấm tự động theo mã dạng (mos/word_auto.py).

    import_path(đường_dẫn)  – thư mục bộ đề, một thư mục De_xx, hoặc file .zip
    auto_import()           – tự nhập mọi bộ đề đặt trong thư mục nhap_de/ cạnh app
"""
from __future__ import annotations

import hashlib
import json
import re
import shutil
import tempfile
import zipfile
from pathlib import Path

from . import word_auto
from .core import app_dir
from .custom import editable_dir, read_exam_json, save_exam_json

MON_BY_EXT = {".docx": "WORD", ".xlsx": "EXCEL", ".pptx": "POWERPOINT"}
MON_LABEL = {"WORD": "Word", "EXCEL": "Excel", "POWERPOINT": "PowerPoint"}
AUTO_DIR_NAME = "nhap_de"


class ImportError_(Exception):
    """Lỗi khi nhập bộ đề (thông báo tiếng Việt)."""


def _set_json(folder: Path) -> Path | None:
    """File de_xx.json dạng danh sách project, nếu thư mục là một đề của bộ đề."""
    for f in sorted(folder.glob("de_*.json")):
        try:
            data = json.loads(f.read_text(encoding="utf-8-sig"))
        except (OSError, ValueError):
            continue
        if isinstance(data, list) and data and isinstance(data[0], dict) and "tasks" in data[0]:
            return f
    return None


def find_sets(root: Path) -> list[Path]:
    """Các thư mục đề (De_xx) nằm trong `root` (tìm sâu 3 cấp)."""
    root = Path(root)
    found = []
    for depth in range(4):
        pattern = "/".join(["*"] * depth) if depth else "."
        for folder in sorted(root.glob(pattern)) if depth else [root]:
            if folder.is_dir() and folder not in found and _set_json(folder):
                found.append(folder)
    return found


def _slug(text: str) -> str:
    import unicodedata
    text = unicodedata.normalize("NFKD", text.replace("đ", "d").replace("Đ", "D"))
    text = "".join(c for c in text if not unicodedata.combining(c))
    return re.sub(r"[^A-Za-z0-9]+", "_", text).strip("_") or "De"


def _collection(folder: Path) -> str:
    """Tên bộ đề lấy từ thư mục cha, vd MOS_Word365_DeThucTe -> Word365_DeThucTe."""
    name = folder.parent.name
    return re.sub(r"^MOS_", "", name) if name and name not in (".", "/") else ""


def _files_dir(folder: Path) -> Path:
    for name in ("Files", "files", "File"):
        if (folder / name).is_dir():
            return folder / name
    return folder


def _hint(steps) -> str:
    if isinstance(steps, str):
        return steps
    steps = [str(s).strip() for s in steps or [] if str(s).strip()]
    if len(steps) == 1:
        return steps[0]
    return "  ".join(f"{i}) {s}" for i, s in enumerate(steps, start=1))


def _signature(folder: Path) -> str:
    """Dấu vết nội dung thư mục nguồn để biết bộ đề đã đổi hay chưa."""
    items = []
    for f in sorted(folder.rglob("*")):
        if f.is_file():
            st = f.stat()
            items.append(f"{f.relative_to(folder)}:{st.st_size}:{int(st.st_mtime)}")
    return hashlib.md5("|".join(items).encode("utf-8")).hexdigest()[:16] if items else ""


def translation_memory() -> dict:
    """Bộ nhớ dịch (mos/dich_de.json): đề bài Anh → Việt, gợi ý Việt → Anh."""
    try:
        data = json.loads((Path(__file__).with_name("dich_de.json")).read_text(encoding="utf-8"))
    except (OSError, ValueError):
        data = {}
    return {"de_bai": data.get("de_bai", {}), "goi_y": data.get("goi_y", {})}


def convert_set(folder: Path, dest_base: Path | None = None) -> tuple[Path, list[str]]:
    """Chuyển một đề (De_xx) sang định dạng của app. Trả về (thư mục đề mới, cảnh báo)."""
    folder = Path(folder)
    jf = _set_json(folder)
    if jf is None:
        raise ImportError_(f"{folder.name}: không thấy file de_xx.json của bộ đề")
    entries = json.loads(jf.read_text(encoding="utf-8-sig"))
    files = _files_dir(folder)
    warnings: list[str] = []

    exts = {Path(e.get("file", "")).suffix.lower() for e in entries}
    mon = next((MON_BY_EXT[x] for x in exts if x in MON_BY_EXT), None)
    if mon is None:
        raise ImportError_(f"{folder.name}: không nhận ra loại file thực hành (.docx / .xlsx / .pptx)")

    collection = _collection(folder)
    m = re.search(r"(\d+)", folder.name)
    number = f"Đề {int(m.group(1)):02d}" if m else folder.name
    title = f"{MON_LABEL[mon]} – {number}" + (f" ({collection})" if collection else "")
    title_en = title.replace("Đề ", "Test ")
    source_id = f"{collection}/{folder.name}"

    base = Path(dest_base) if dest_base else editable_dir()
    dest = base / _slug(f"{collection}_{folder.name}" if collection else folder.name)
    if dest.exists():
        try:
            same = read_exam_json(dest).get("nguon") == source_id
        except (OSError, ValueError):
            same = False
        if not same:
            n = 2
            while (base / f"{dest.name}_{n}").exists():
                n += 1
            dest = base / f"{dest.name}_{n}"
        else:
            shutil.rmtree(dest)
    dest.mkdir(parents=True)

    memory = translation_memory()
    projects, main_files = [], set()
    for e in entries:
        src = files / str(e.get("file", ""))
        if not src.is_file():
            warnings.append(f"{folder.name}: thiếu file thực hành {src.name}")
            continue
        shutil.copyfile(src, dest / src.name)
        main_files.add(src.name)
        tasks = []
        for t in e.get("tasks", []):
            text = str(t.get("text", "")).strip()
            if not text:
                continue
            dang = str(t.get("id", ""))
            if mon == "WORD" and dang in word_auto.GRADERS:
                spec = {"luat": "word_mau_de", "dang": dang, "de_bai": text}
                goc = word_auto.snapshot(src, dang, text)
                if goc:
                    spec["goc"] = goc
            else:
                spec = {"luat": "tu_kiem_tra", "ghi_chu": dang}
                warnings.append(f"{folder.name} › {src.name}: câu dạng “{dang}” chưa chấm tự động được "
                                "(học viên tự kiểm tra)")
            hint = _hint(t.get("steps"))
            tasks.append({"yeu_cau": memory["de_bai"].get(text, text), "yeu_cau_en": text,
                          "goi_y": hint, "goi_y_en": memory["goi_y"].get(hint, ""), "cham": spec})
        theme = str(e.get("theme") or Path(src).stem)
        expert = "Expert" in src.stem or e.get("expert")
        no = e.get("no", len(projects) + 1)
        projects.append({
            "ten": f"Dự án {no} – {theme}" + (" (Expert)" if expert else ""),
            "ten_en": f"Project {no} – {theme}" + (" (Expert)" if expert else ""),
            "file": src.name,
            "mo_ta": f"Mở file {src.name} và làm lần lượt các yêu cầu bên dưới.",
            "mo_ta_en": f"Open {src.name} and complete the tasks below in order.",
            "nhiem_vu": tasks,
        })

    extras = []
    for f in sorted(files.iterdir()) if files.is_dir() else []:
        if f.is_file() and f.name not in main_files and not f.name.startswith(("~$", ".")):
            shutil.copyfile(f, dest / f.name)
            extras.append(f.name)
    all_text = " ".join(t["yeu_cau_en"] for p in projects for t in p["nhiem_vu"])
    models = folder.parent / "3D Models"
    if "3D Models" in all_text and models.is_dir():
        shutil.copytree(models, dest / "3D Models")
        extras.append("3D Models")
    if extras:
        for p in projects:
            p["mo_ta"] += " File phụ (ảnh, dữ liệu, 3D Models) nằm cùng thư mục bài làm."
            p["mo_ta_en"] += " Extra files (pictures, data, 3D Models) are in the same work folder."

    data = {"mon": mon, "ten": title, "ten_en": title_en, "thoi_gian": 50, "rieng": True, "nguon": source_id,
            "dau_vet": _signature(folder), "du_an": projects}
    if extras:
        data["file_phu"] = extras
    save_exam_json(dest, data)
    return dest, warnings


def import_path(path: Path, dest_base: Path | None = None) -> tuple[list[Path], list[str]]:
    """Nhập mọi đề tìm thấy trong thư mục hoặc file .zip. Trả về (các thư mục đề, cảnh báo)."""
    path = Path(path)
    if path.is_file() and path.suffix.lower() == ".zip":
        with tempfile.TemporaryDirectory() as tmp:
            try:
                with zipfile.ZipFile(path) as z:
                    z.extractall(tmp)
            except zipfile.BadZipFile as exc:
                raise ImportError_(f"{path.name}: file ZIP bị hỏng") from exc
            return import_path(Path(tmp), dest_base)
    if path.is_file() and path.name.startswith("de_") and path.suffix == ".json":
        path = path.parent
    if not path.is_dir():
        raise ImportError_(f"Không tìm thấy {path}")
    sets = find_sets(path)
    if not sets:
        raise ImportError_(f"{path.name}: không thấy đề nào (cần thư mục De_xx có file de_xx.json)")
    done, warnings = [], []
    for folder in sets:
        dest, warn = convert_set(folder, dest_base)
        done.append(dest)
        warnings.extend(warn)
    return done, warnings


def auto_import(src_dir: Path | None = None, dest_base: Path | None = None) -> tuple[list[Path], list[str]]:
    """Tự nhập bộ đề đặt trong thư mục nhap_de/ cạnh app (chỉ nhập lại khi nội dung đổi)."""
    src_dir = Path(src_dir) if src_dir else app_dir() / AUTO_DIR_NAME
    if not src_dir.is_dir():
        return [], []
    base = Path(dest_base) if dest_base else editable_dir()
    known = {}
    if base.is_dir():
        for f in base.glob("*/de.json"):
            try:
                data = read_exam_json(f.parent)
            except (OSError, ValueError):
                continue
            if data.get("nguon"):
                known[data["nguon"]] = data.get("dau_vet")
    done, warnings = [], []
    for item in sorted(src_dir.iterdir()):
        try:
            if item.suffix.lower() == ".zip":
                d, w = import_path(item, base)
                done += d
                warnings += w
                item.rename(item.with_name(item.name + ".da_nhap"))
                continue
            for folder in find_sets(item) if item.is_dir() else []:
                source_id = f"{_collection(folder)}/{folder.name}"
                if known.get(source_id) == _signature(folder):
                    continue
                dest, w = convert_set(folder, base)
                done.append(dest)
                warnings += w
        except (ImportError_, OSError, ValueError) as exc:
            warnings.append(str(exc))
    return done, warnings
