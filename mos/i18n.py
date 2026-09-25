"""Đa ngôn ngữ Việt / Anh.

Chữ trên giao diện viết bằng tiếng Việt trong code và bọc bằng `tr("...")`. Khi đang ở
chế độ tiếng Anh, `tr` tra bản dịch trong mos/i18n_en.json (khóa = câu tiếng Việt);
câu có biến dùng `tr("Đã làm {n}/{total}").format(n=..., total=...)`.

Nội dung đề (tên đề, yêu cầu, gợi ý) có sẵn hai bản trong dữ liệu; chọn bằng `pick(vi, en)`.
Ngôn ngữ đang dùng được lưu trong ~/MOS_Practice/cai_dat.json (riêng từng máy).
"""
from __future__ import annotations

import json
from pathlib import Path

from .core import APP_DIR

LANGS = ("vi", "en")
LANG_NAMES = {"vi": "Tiếng Việt", "en": "English"}
SETTINGS_FILE = APP_DIR / "cai_dat.json"

_lang: str | None = None
_en: dict | None = None


def _settings() -> dict:
    try:
        data = json.loads(SETTINGS_FILE.read_text(encoding="utf-8"))
        return data if isinstance(data, dict) else {}
    except (OSError, ValueError):
        return {}


def get_lang() -> str:
    global _lang
    if _lang is None:
        _lang = _settings().get("ngon_ngu", "vi")
        if _lang not in LANGS:
            _lang = "vi"
    return _lang


def set_lang(code: str, save: bool = True) -> None:
    """Đổi ngôn ngữ (và lưu lại cho lần mở sau)."""
    global _lang
    if code not in LANGS:
        raise ValueError(f"ngôn ngữ không hỗ trợ: {code}")
    _lang = code
    if save:
        data = _settings()
        data["ngon_ngu"] = code
        try:
            SETTINGS_FILE.parent.mkdir(parents=True, exist_ok=True)
            SETTINGS_FILE.write_text(json.dumps(data, ensure_ascii=False, indent=2), encoding="utf-8")
        except OSError:
            pass


def is_en() -> bool:
    return get_lang() == "en"


def english() -> dict:
    global _en
    if _en is None:
        try:
            _en = json.loads(Path(__file__).with_name("i18n_en.json").read_text(encoding="utf-8"))
        except (OSError, ValueError):
            _en = {}
    return _en


def tr(text: str) -> str:
    """Dịch chữ giao diện (viết bằng tiếng Việt) sang ngôn ngữ đang dùng."""
    if not is_en():
        return text
    return english().get(text, text)


ROLE_EN = {"Quản trị": "Admin", "Học viên": "Student"}


def role_name(vi: str) -> str:
    """Tên vai trò tài khoản theo ngôn ngữ."""
    return pick(vi, ROLE_EN.get(vi))


def pick(vi: str, en: str | None) -> str:
    """Chọn bản tiếng Việt / tiếng Anh của nội dung đề (thiếu bản nào thì dùng bản còn lại)."""
    if is_en():
        return en or vi
    return vi or (en or "")
