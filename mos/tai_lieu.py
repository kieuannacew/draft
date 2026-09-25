"""Tài liệu học (bài giảng PowerPoint) xem ngay trong app, không cho tải về.

Mỗi bài giảng là một thư mục `tai_lieu/<id>/`:
    bai.json      tên (Việt/Anh), môn, chương, chữ + ghi chú của từng slide (để tra cứu)
    slides.mosl   ảnh các slide, đóng gói và mã hóa (không mở được bằng trình xem ảnh / PowerPoint)

Nhập bài từ file .pptx:
  1. Trên Windows có PowerPoint: PowerPoint xuất từng slide thành ảnh PNG (đẹp y như bản gốc).
  2. Không có PowerPoint: app tự dựng ảnh slide từ chữ, hình và ảnh trong file (đơn giản hơn).
File .pptx gốc KHÔNG được chép vào app. Học viên chỉ xem ảnh trong trình xem, không có nút lưu / xuất.
(Không chặn được chụp màn hình – điều này đúng với mọi phần mềm.)
"""
from __future__ import annotations

import hashlib
import json
import os
import re
import shutil
import subprocess
import tempfile
import unicodedata
import zipfile
from dataclasses import dataclass, field
from pathlib import Path

from .core import DATA_DIR, app_dir

PACK = "slides.mosl"
META = "bai.json"
SLIDE_WIDTH = 1600
_MAGIC = b"MOSL1"
_QT_APP = None


def search_dirs() -> list[Path]:
    return [app_dir() / "tai_lieu", DATA_DIR / "tai_lieu"]


def editable_dir() -> Path:
    return DATA_DIR / "tai_lieu"


# ====================================================================== đóng gói ảnh


def _key(lesson_id: str) -> bytes:
    return hashlib.sha256(("mos-bai-giang:" + lesson_id).encode("utf-8")).digest()


def _xor(data: bytes, key: bytes) -> bytes:
    k = (key * (len(data) // len(key) + 1))[:len(data)]
    return bytes(a ^ b for a, b in zip(data, k))


def pack_slides(folder: Path, lesson_id: str, images: list[bytes]) -> None:
    key = _key(lesson_id)
    with zipfile.ZipFile(folder / PACK, "w", zipfile.ZIP_STORED) as z:
        z.writestr("id", _MAGIC + lesson_id.encode("utf-8"))
        for i, img in enumerate(images, start=1):
            z.writestr(f"{i:04d}", _xor(img, key))


def read_slide(folder: Path, lesson_id: str, number: int) -> bytes:
    """Ảnh PNG của slide thứ `number` (1, 2, …) – chỉ giải mã trong bộ nhớ."""
    with zipfile.ZipFile(folder / PACK) as z:
        return _xor(z.read(f"{number:04d}"), _key(lesson_id))


# ====================================================================== dữ liệu bài giảng


@dataclass
class Lesson:
    id: str
    folder: Path
    ten: str
    ten_en: str = ""
    mon: str = "WORD"
    chuong: int | None = None
    mo_ta: str = ""
    mo_ta_en: str = ""
    slides: list[dict] = field(default_factory=list)   # [{"text": ..., "notes": ...}]

    @property
    def count(self) -> int:
        return len(self.slides)

    def image(self, number: int) -> bytes:
        return read_slide(self.folder, self.id, number)


def load_lesson(folder: Path) -> Lesson | None:
    try:
        data = json.loads((folder / META).read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return None
    if not (folder / PACK).is_file():
        return None
    return Lesson(id=data.get("id", folder.name), folder=folder, ten=data.get("ten", folder.name),
                  ten_en=data.get("ten_en", ""), mon=str(data.get("mon", "WORD")).upper(),
                  chuong=data.get("chuong"), mo_ta=data.get("mo_ta", ""), mo_ta_en=data.get("mo_ta_en", ""),
                  slides=list(data.get("slides", [])))


def list_lessons() -> list[Lesson]:
    out, seen = [], set()
    for base in search_dirs():
        if not base.is_dir():
            continue
        for folder in sorted(p for p in base.iterdir() if p.is_dir()):
            lesson = load_lesson(folder)
            if lesson and lesson.id not in seen:
                seen.add(lesson.id)
                out.append(lesson)
    order = {"WORD": 0, "EXCEL": 1, "POWERPOINT": 2}
    return sorted(out, key=lambda x: (order.get(x.mon, 9), x.chuong or 99, x.ten.casefold()))


def delete_lesson(lesson: Lesson) -> None:
    shutil.rmtree(lesson.folder, ignore_errors=True)


def _slug(text: str) -> str:
    text = unicodedata.normalize("NFKD", text.replace("đ", "d").replace("Đ", "D"))
    text = "".join(c for c in text if not unicodedata.combining(c))
    return re.sub(r"[^A-Za-z0-9]+", "_", text).strip("_").lower() or "bai"


# ====================================================================== đọc chữ trong pptx


def _shape_text(shape) -> list[str]:
    out = []
    if getattr(shape, "has_text_frame", False) and shape.has_text_frame:
        out += [p.text for p in shape.text_frame.paragraphs if p.text.strip()]
    if getattr(shape, "has_table", False) and shape.has_table:
        for row in shape.table.rows:
            out.append("  |  ".join(c.text for c in row.cells))
    if shape.shape_type == 6:           # nhóm
        for s in shape.shapes:
            out += _shape_text(s)
    return out


def extract_text(pptx: Path) -> list[dict]:
    from pptx import Presentation
    prs = Presentation(str(pptx))
    slides = []
    for slide in prs.slides:
        lines = []
        for shape in slide.shapes:
            lines += _shape_text(shape)
        notes = ""
        if slide.has_notes_slide:
            notes = slide.notes_slide.notes_text_frame.text.strip()
        slides.append({"text": "\n".join(lines), "notes": notes})
    return slides


# ====================================================================== ảnh slide


def export_with_powerpoint(pptx: Path, out_dir: Path, width: int = SLIDE_WIDTH) -> list[bytes]:
    """Nhờ PowerPoint (Windows) xuất mỗi slide thành PNG. Lỗi / không có PowerPoint → []."""
    if os.name != "nt":
        return []
    script = (
        "$ErrorActionPreference='Stop';"
        "$pp = New-Object -ComObject PowerPoint.Application;"
        f"$p = $pp.Presentations.Open('{pptx}', $true, $false, $false);"
        f"$h = [int]({width} * $p.PageSetup.SlideHeight / $p.PageSetup.SlideWidth);"
        f"$p.Export('{out_dir}', 'PNG', {width}, $h);"
        "$p.Close();"
        "if ($pp.Presentations.Count -eq 0) { $pp.Quit() }"
    )
    try:
        subprocess.run(["powershell", "-NoProfile", "-NonInteractive", "-Command", script],
                       check=True, capture_output=True, timeout=300,
                       creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0))
    except (OSError, subprocess.SubprocessError):
        return []
    files = [f for f in out_dir.glob("*") if f.suffix.lower() == ".png"]
    files.sort(key=lambda f: int(re.findall(r"\d+", f.stem)[-1]) if re.findall(r"\d+", f.stem) else 0)
    return [f.read_bytes() for f in files]


def _rgb(fill, default):
    from PySide6.QtGui import QColor
    try:
        if fill.type == 1:           # MSO_FILL.SOLID
            return QColor("#" + str(fill.fore_color.rgb))
    except (AttributeError, TypeError, ValueError, KeyError):
        pass
    return default


def _font_color(run, default):
    from PySide6.QtGui import QColor
    try:
        if run.font.color and run.font.color.type is not None and run.font.color.rgb is not None:
            return QColor("#" + str(run.font.color.rgb))
    except (AttributeError, TypeError, ValueError):
        pass
    return default


def render_basic(pptx: Path, width: int = SLIDE_WIDTH) -> list[bytes]:
    """Tự dựng ảnh slide (không cần PowerPoint): nền, hình (chữ nhật / bo góc / tròn), ảnh, chữ, bảng."""
    from pptx import Presentation
    from pptx.enum.shapes import MSO_SHAPE
    from pptx.enum.text import MSO_ANCHOR, PP_ALIGN
    from pptx.util import Pt
    from PySide6.QtCore import QBuffer, QByteArray, QIODevice, QRectF, Qt
    from PySide6.QtGui import QColor, QFont, QFontDatabase, QFontMetricsF, QGuiApplication, QImage, QPainter, QPen

    global _QT_APP
    if QGuiApplication.instance() is None:
        _QT_APP = QGuiApplication([])   # cần để vẽ chữ khi chạy ngoài giao diện
    families = set(QFontDatabase.families())
    prs = Presentation(str(pptx))
    sw, sh = prs.slide_width or 12192000, prs.slide_height or 6858000
    scale = width / sw
    height = int(sh * scale)
    images = []

    def qfont(run, pt, bold, italic):
        f = QFont()
        name = run.font.name if run is not None else None
        if name and name in families:
            f.setFamily(name)
        f.setPixelSize(max(9, int(Pt(pt) * scale)))
        f.setBold(bool(bold))
        f.setItalic(bool(italic))
        return f

    def draw_text(p: QPainter, shape, rect: QRectF, is_title: bool):
        tf = shape.text_frame
        ml = (tf.margin_left if tf.margin_left is not None else 91440) * scale
        mr = (tf.margin_right if tf.margin_right is not None else 91440) * scale
        mt = (tf.margin_top if tf.margin_top is not None else 45720) * scale
        mb = (tf.margin_bottom if tf.margin_bottom is not None else 45720) * scale
        inner = rect.adjusted(ml, mt, -mr, -mb)
        blocks = []
        for para in tf.paragraphs:
            run = next((r for r in para.runs if r.text.strip()), para.runs[0] if para.runs else None)
            size = next((r.font.size for r in para.runs if r.font.size), None)
            pt = size.pt if size else (36 if is_title else 18)
            bold = is_title or any(r.font.bold for r in para.runs)
            italic = any(r.font.italic for r in para.runs)
            font = qfont(run, pt, bold, italic)
            color = _font_color(run, QColor("#1E2A1D")) if run is not None else QColor("#1E2A1D")
            align = {PP_ALIGN.CENTER: Qt.AlignHCenter, PP_ALIGN.RIGHT: Qt.AlignRight}.get(para.alignment, Qt.AlignLeft)
            text = para.text
            if para.level and not text.startswith("•"):
                text = "    " * para.level + text
            p.setFont(font)
            h = p.boundingRect(QRectF(0, 0, inner.width(), 10000), Qt.TextWordWrap | align, text or " ").height()
            if h <= 0:      # một số font đậm cỡ nhỏ trả về 0: ước lượng theo số dòng
                fm = QFontMetricsF(font)
                lines = max(1, int(fm.horizontalAdvance(text) // max(inner.width(), 1)) + 1)
                h = fm.lineSpacing() * lines
            blocks.append((text, font, color, align, h))
        total = sum(b[4] for b in blocks) + 4 * max(0, len(blocks) - 1)
        anchor = tf.vertical_anchor if tf.vertical_anchor is not None else (
            MSO_ANCHOR.MIDDLE if getattr(shape, "auto_shape_type", None) not in (None,) and not
            getattr(shape, "is_placeholder", False) and shape.shape_type == 1 else MSO_ANCHOR.TOP)
        y = inner.y()
        if anchor == MSO_ANCHOR.MIDDLE:
            y = inner.y() + (inner.height() - total) / 2
        elif anchor == MSO_ANCHOR.BOTTOM:
            y = inner.bottom() - total
        for text, font, color, align, h in blocks:
            p.setFont(font)
            p.setPen(color)
            p.drawText(QRectF(inner.x(), y, inner.width(), h), Qt.TextWordWrap | align, text)
            y += h + 4

    def draw_shape(p: QPainter, shape):
        if shape.left is None or shape.width is None:
            if shape.shape_type == 6:
                for s in shape.shapes:
                    draw_shape(p, s)
            return
        rect = QRectF(shape.left * scale, shape.top * scale, shape.width * scale, shape.height * scale)
        if shape.shape_type == 6:
            for s in shape.shapes:
                draw_shape(p, s)
            return
        if shape.shape_type == 13 and hasattr(shape, "image"):          # ảnh
            img = QImage.fromData(shape.image.blob)
            if not img.isNull():
                p.drawImage(rect, img)
            return
        try:
            fill = _rgb(shape.fill, None)
        except (AttributeError, TypeError):
            fill = None
        if fill is not None:
            p.setPen(Qt.NoPen)
            p.setBrush(fill)
            kind = getattr(shape, "auto_shape_type", None) if shape.shape_type == 1 else None
            if kind == MSO_SHAPE.OVAL:
                p.drawEllipse(rect)
            elif kind == MSO_SHAPE.ROUNDED_RECTANGLE:
                r = min(rect.width(), rect.height()) * 0.12
                p.drawRoundedRect(rect, r, r)
            else:
                p.drawRect(rect)
            p.setBrush(Qt.NoBrush)
        if getattr(shape, "has_table", False) and shape.has_table:
            rows = list(shape.table.rows)
            cols = len(rows[0].cells) if rows else 0
            if rows and cols:
                rh, cw = rect.height() / len(rows), rect.width() / cols
                f = QFont()
                f.setPixelSize(max(12, int(rh * 0.35)))
                p.setFont(f)
                for r, row in enumerate(rows):
                    for c, cell in enumerate(row.cells):
                        cr = QRectF(rect.x() + c * cw, rect.y() + r * rh, cw, rh)
                        if r == 0:
                            p.fillRect(cr, QColor("#ECEFE3"))
                        p.setPen(QPen(QColor("#9CA38A"), 1))
                        p.drawRect(cr)
                        p.setPen(QColor("#1E2A1D"))
                        p.drawText(cr.adjusted(8, 4, -8, -4), Qt.AlignVCenter | Qt.TextWordWrap, cell.text)
            return
        if getattr(shape, "has_text_frame", False) and shape.has_text_frame and shape.text_frame.text.strip():
            is_title = bool(shape.is_placeholder and shape.placeholder_format.type in (1, 3))
            draw_text(p, shape, rect, is_title)

    for slide in prs.slides:
        img = QImage(width, height, QImage.Format_RGB32)
        img.fill(QColor("white"))
        p = QPainter(img)
        p.setRenderHint(QPainter.Antialiasing)
        p.setRenderHint(QPainter.TextAntialiasing)
        try:
            color = _rgb(slide.background.fill, None)
            if color is not None:
                p.fillRect(img.rect(), color)
        except (AttributeError, TypeError):
            pass
        for shape in slide.shapes:
            try:
                draw_shape(p, shape)
            except Exception:  # hình lạ: bỏ qua, không làm hỏng cả slide
                continue
        p.end()
        buf = QByteArray()
        dev = QBuffer(buf)
        dev.open(QIODevice.WriteOnly)
        img.save(dev, "PNG")
        images.append(bytes(buf.data()))
    return images


# ====================================================================== nhập bài giảng


def import_pptx(pptx: Path, ten: str, mon: str, chuong: int | None = None, ten_en: str = "",
                mo_ta: str = "", dest_base: Path | None = None, use_powerpoint: bool = True) -> Lesson:
    """Nhập một file .pptx thành bài giảng (ảnh slide đóng gói + chữ để tra cứu)."""
    pptx = Path(pptx)
    slides = extract_text(pptx)
    images: list[bytes] = []
    if use_powerpoint:
        with tempfile.TemporaryDirectory() as tmp:
            images = export_with_powerpoint(pptx.resolve(), Path(tmp))
    if len(images) != len(slides):
        images = render_basic(pptx)
    base = Path(dest_base) if dest_base else editable_dir()
    slug = _slug(ten)
    folder, n = base / slug, 1
    while folder.exists():
        n += 1
        folder = base / f"{slug}_{n}"
    folder.mkdir(parents=True)
    lesson_id = f"{slug}-{hashlib.sha1(pptx.read_bytes()).hexdigest()[:10]}"
    pack_slides(folder, lesson_id, images)
    data = {"id": lesson_id, "ten": ten, "ten_en": ten_en, "mon": mon.upper(), "chuong": chuong,
            "mo_ta": mo_ta, "nguon": pptx.name, "slides": slides}
    (folder / META).write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
    return load_lesson(folder)


# ====================================================================== tiến độ học


def _progress_file() -> Path:
    return DATA_DIR / "tien_do_hoc.json"


def load_progress(user: str | None) -> dict[str, int]:
    """{id bài giảng: slide xa nhất đã xem} của một tài khoản."""
    try:
        data = json.loads(_progress_file().read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return {}
    return data.get(user or "", {})


def save_progress(user: str | None, lesson_id: str, slide: int) -> None:
    try:
        data = json.loads(_progress_file().read_text(encoding="utf-8"))
    except (OSError, ValueError):
        data = {}
    mine = data.setdefault(user or "", {})
    if slide > mine.get(lesson_id, 0):
        mine[lesson_id] = slide
        _progress_file().parent.mkdir(parents=True, exist_ok=True)
        _progress_file().write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
