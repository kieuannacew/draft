"""Kiểm tra tài khoản (mos/accounts.py) và form luật chấm (rules.parse_value...)."""
import json
from datetime import date, timedelta

import pytest

from mos import rules
from mos.accounts import ADMIN, STUDENT, AccountError, AccountStore


@pytest.fixture
def store(tmp_path):
    return AccountStore(tmp_path / "tai_khoan.json")


def test_first_run_creates_default_admin(store):
    admin = store.authenticate("admin", "admin")
    assert admin.is_admin and admin.doi_mat_khau


def test_password_is_hashed_not_stored(store, tmp_path):
    store.create("nguyenvana", "Nguyễn Văn A", "matkhau1")
    raw = (tmp_path / "tai_khoan.json").read_text(encoding="utf-8")
    assert "matkhau1" not in raw
    assert json.loads(raw)["nguyenvana"]["ho_ten"] == "Nguyễn Văn A"


def test_login_rules(store):
    store.create("hv01", "Học viên 1", "abcd", STUDENT, must_change=False)
    assert store.authenticate(" HV01 ", "abcd").username == "hv01"        # không phân biệt hoa/thường
    with pytest.raises(AccountError, match="Sai tên"):
        store.authenticate("hv01", "sai")
    with pytest.raises(AccountError, match="Sai tên"):
        store.authenticate("khongco", "abcd")
    store.update("hv01", khoa=True)
    with pytest.raises(AccountError, match="khóa"):
        store.authenticate("hv01", "abcd")
    store.update("hv01", khoa=False, han_dung=(date.today() - timedelta(days=1)).isoformat())
    with pytest.raises(AccountError, match="hết hạn"):
        store.authenticate("hv01", "abcd")
    store.update("hv01", han_dung=date.today().isoformat())   # hết hạn SAU ngày này
    assert store.authenticate("hv01", "abcd")


def test_persistence_and_password_change(store, tmp_path):
    store.create("hv02", "B", "1111")
    store.set_password("hv02", "2222")
    again = AccountStore(tmp_path / "tai_khoan.json")
    acc = again.authenticate("hv02", "2222")
    assert not acc.doi_mat_khau
    with pytest.raises(AccountError):
        again.authenticate("hv02", "1111")


@pytest.mark.parametrize("username,password,han,msg", [
    ("ab", "1234", None, "3–32"),
    ("có dấu", "1234", None, "3–32"),
    ("admin", "1234", None, "đã tồn tại"),
    ("hv03", "12", None, "ít nhất"),
    ("hv03", "1234", "31/12/2026", "YYYY-MM-DD"),
])
def test_create_validation(store, username, password, han, msg):
    with pytest.raises(AccountError, match=msg):
        store.create(username, "X", password, han_dung=han)


def test_last_admin_is_protected(store):
    with pytest.raises(AccountError):
        store.delete("admin")
    with pytest.raises(AccountError):
        store.update("admin", khoa=True)
    with pytest.raises(AccountError):
        store.update("admin", vai_tro=STUDENT)
    store.create("gv2", "Giáo viên 2", "1234", ADMIN)
    store.delete("admin")
    assert [a.username for a in store.list()] == ["gv2"]


# ---------------------------------------------------------------- form luật chấm

@pytest.mark.parametrize("kind,text,value", [
    ("list", "C{hang}; *;D{hang} ", ["C{hang}", "*", "D{hang}"]),
    ("number", "1,5", 1.5),
    ("number", "4", 4),
    ("auto", "cuoi", "cuoi"),
    ("auto", "3", 3),
    ("bool", "Có", True),
    ("bool", "Không", False),
    ("text", "  ", None),
])
def test_parse_value(kind, text, value):
    assert rules.parse_value(kind, text) == value


def test_parse_value_errors():
    with pytest.raises(ValueError):
        rules.parse_value("number", "abc")
    with pytest.raises(ValueError):
        rules.parse_value("bool", "maybe")


def test_every_rule_has_form_metadata():
    for name in rules.RULES:
        assert name in rules.RULE_TITLES
        for p in rules.rule_params(name):
            assert p["label"] != p["name"], (name, p["name"])
            value = rules.parse_value(p["kind"], rules.format_value(p["kind"], p["default"]))
            assert value == p["default"] or p["default"] is None, (name, p["name"])
    assert rules.rules_for("WORD")[0].startswith("word_") and rules.rules_for("WORD")[-1] == "xml_chua"
