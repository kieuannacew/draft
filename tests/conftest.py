"""Cấu hình chung cho pytest: luôn chạy test bằng tiếng Việt (không phụ thuộc cài đặt trên máy)."""
import pytest

from mos import i18n


@pytest.fixture(autouse=True)
def _tieng_viet():
    i18n.set_lang("vi", save=False)
    yield
    i18n.set_lang("vi", save=False)
