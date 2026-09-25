"""Kiểm tra song ngữ Việt / Anh (mos/i18n.py) và phần "vui học" (mos/gamify.py)."""
import ast
import json
import re
from datetime import date
from pathlib import Path

import pytest

from mos import gamify, i18n
from mos.custom import load_exam
from mos.exams import ALL_EXAMS

ROOT = Path(__file__).resolve().parents[1]


def _tr_keys():
    keys = set()
    for f in (ROOT / "mos").rglob("*.py"):
        for node in ast.walk(ast.parse(f.read_text(encoding="utf-8"))):
            if isinstance(node, ast.Call) and getattr(node.func, "id", None) == "tr" and node.args:
                arg = node.args[0]
                if isinstance(arg, ast.Constant) and isinstance(arg.value, str):
                    keys.add(arg.value)
    return keys


def _placeholders(text):
    return sorted(re.findall(r"\{\w*\}", text))


def test_every_ui_text_has_english():
    en = i18n.english()
    missing = sorted(k for k in _tr_keys() if k not in en)
    assert not missing, missing
    for vi, text in en.items():
        assert _placeholders(vi) == _placeholders(text), vi
        assert vi.count("&&") == text.count("&&"), vi


def test_rule_titles_and_labels_translated():
    from mos import rules
    en = i18n.english()
    for title in rules.RULE_TITLES.values():
        assert title in en, title
    for label, *_ in list(rules.PARAM_UI.values()) + list(rules.PARAM_UI_RULE.values()):
        assert label in en, label


def test_switch_language():
    assert i18n.tr("Trang chủ") == "Trang chủ"
    assert i18n.pick("Việt", "English") == "Việt"
    i18n.set_lang("en", save=False)
    assert i18n.tr("Trang chủ") == "Home"
    assert i18n.tr("Đã làm {n}/{total}").format(n=1, total=2) == "Done 1/2"
    assert i18n.pick("Việt", "English") == "English"
    assert i18n.pick("Việt", "") == "Việt"                  # thiếu bản Anh → dùng bản Việt
    assert i18n.tr("câu chưa có trong từ điển") == "câu chưa có trong từ điển"
    with pytest.raises(ValueError):
        i18n.set_lang("fr", save=False)


def test_language_is_saved(tmp_path, monkeypatch):
    monkeypatch.setattr(i18n, "SETTINGS_FILE", tmp_path / "cai_dat.json")
    i18n.set_lang("en")
    assert json.loads((tmp_path / "cai_dat.json").read_text(encoding="utf-8"))["ngon_ngu"] == "en"
    monkeypatch.setattr(i18n, "_lang", None)
    assert i18n.get_lang() == "en"


def test_builtin_exams_are_bilingual():
    for exam in ALL_EXAMS:
        for p in exam.projects:
            assert p.name_en and p.intro_en
            for t in p.tasks:
                assert t.title_en and t.hint_en, t.title


def test_bundled_exams_are_bilingual():
    folders = sorted((ROOT / "de_thi").glob("*/de.json"))
    assert folders
    for f in folders:
        data = json.loads(f.read_text(encoding="utf-8"))
        assert data.get("ten_en"), f
        for p in data["du_an"]:
            assert p.get("ten_en") and p.get("mo_ta_en"), f
            for t in p["nhiem_vu"]:
                assert t.get("yeu_cau_en") and t.get("goi_y_en"), t["yeu_cau"]
                # chữ trong ngoặc (nội dung file Word) phải giữ nguyên khi dịch
                q = r"“[^”]*”"
                assert sorted(re.findall(q, t["yeu_cau"])) == sorted(re.findall(q, t["yeu_cau_en"])), t["yeu_cau"]
        exam, errors = load_exam(f.parent)
        assert not errors and exam.name_en


def test_translation_memory_consistent():
    mem = json.loads((ROOT / "mos" / "dich_de.json").read_text(encoding="utf-8"))
    assert len(mem["de_bai"]) > 300 and len(mem["goi_y"]) > 300
    for en, vi in mem["de_bai"].items():
        assert sorted(re.findall(r"“[^”]*”", en)) == sorted(re.findall(r"“[^”]*”", vi)), en
    for vi, en in mem["goi_y"].items():
        assert re.findall(r"\d+\)", vi) == re.findall(r"\d+\)", en), vi


# ====================================================================== vui học


def _h(day, score, mode="training", exam="Word – Đề 01", code="WORD", duration=1800):
    return {"time": f"{day} 10:00", "score": score, "passed": score >= 700, "mode": mode, "exam": exam,
            "code": code, "duration": duration}


def test_xp_and_level():
    assert gamify.attempt_xp(_h("2026-01-01", 650)) == 65
    assert gamify.attempt_xp(_h("2026-01-01", 800)) == 130
    assert gamify.attempt_xp(_h("2026-01-01", 1000)) == 250
    assert gamify.attempt_xp(_h("2026-01-01", 800, "testing")) == 156
    lv = gamify.level(0)
    assert lv.number == 1 and lv.xp_for_next == gamify.XP_PER_LEVEL
    lv = gamify.level(gamify.XP_PER_LEVEL * 2 + 100)
    assert lv.number == 3 and lv.xp_in_level == 100 and 0 < lv.ratio < 1
    assert gamify.level(10 ** 6).name_en == gamify.LEVELS[-1][1]


def test_streak():
    today = date(2026, 9, 25)
    hist = [_h("2026-09-23", 500), _h("2026-09-24", 500), _h("2026-09-25", 500)]
    assert gamify.streak(hist, today) == 3
    assert gamify.streak(hist[:2], today) == 2          # hôm nay chưa làm vẫn giữ chuỗi
    assert gamify.streak(hist[:1], today) == 0
    assert gamify.streak([], today) == 0


def test_stars():
    assert [gamify.stars(x) for x in (None, 300, 450, 700, 950)] == [0, 0, 1, 2, 3]


def test_badges():
    assert not any(b.earned for b in gamify.badges([]))
    hist = [_h("2026-09-01", 1000, "testing", duration=600), _h("2026-09-02", 700, exam="Excel", code="EXCEL"),
            _h("2026-09-03", 400, exam="PP", code="POWERPOINT")]
    got = {b.key for b in gamify.badges(hist) if b.earned}
    assert {"first", "pass", "perfect", "streak3", "trio", "speed"} <= got
    assert "ten" not in got and "streak7" not in got
    new = gamify.new_badges(hist[:1], hist)
    assert {b.key for b in new} == {"streak3", "trio"}


def test_tip_of_day_rotates():
    tips = {gamify.tip_of_day(date(2026, 1, d)) for d in range(1, 20)}
    assert len(tips) == len(gamify.TIPS)
    assert all(vi and en for vi, en in tips)
