"""Tiện ích đọc XML bên trong file Office.

File .xlsx/.docx/.pptx thực chất là file ZIP chứa các file XML. Một số thứ
(biểu đồ Excel, watermark, footnote, section PowerPoint...) thư viện Python
không đọc được, nên ta mở thẳng ZIP và tìm trong XML.
"""
from __future__ import annotations

import re
import zipfile
from pathlib import Path


def parts(path: Path, pattern: str) -> dict[str, str]:
    """Trả về {tên_part: nội_dung_xml} cho các part khớp regex `pattern`."""
    rx = re.compile(pattern)
    with zipfile.ZipFile(path) as z:
        return {n: z.read(n).decode("utf-8", "replace") for n in z.namelist() if rx.fullmatch(n)}


def part(path: Path, name: str) -> str:
    with zipfile.ZipFile(path) as z:
        if name not in z.namelist():
            return ""
        return z.read(name).decode("utf-8", "replace")


def any_part_contains(path: Path, pattern: str, regex: str, flags: int = 0) -> bool:
    rx = re.compile(regex, flags)
    return any(rx.search(xml) for xml in parts(path, pattern).values())


def norm(text: str) -> str:
    """Chuẩn hóa để so sánh chuỗi: bỏ khoảng trắng thừa, chữ thường."""
    return " ".join(text.split()).casefold()
