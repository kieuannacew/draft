"""Kiểm tra bộ đề nhập (mos/nhap_de.py) và chấm tự động câu Word (mos/word_auto.py).

Với MỌI câu trong các đề đã nhập ở de_thi/Word365_*:
  * file gốc phải chấm SAI (trừ vài câu bộ sinh đề tạo sẵn đã đạt – xem ALREADY_DONE),
  * file sau khi áp lời giải mô phỏng (tests/word_solutions.py) phải chấm ĐÚNG.
"""
import json
import shutil
import zipfile
from pathlib import Path

import pytest

from mos import custom, nhap_de, word_auto
from mos.core import grade, new_session
from tests.word_solutions import SOL, solve

ROOT = Path(__file__).resolve().parents[1]
SETS = sorted((ROOT / "de_thi").glob("Word365_*"))

# (file, dạng): file gốc do bộ sinh đề tạo đã đạt sẵn yêu cầu (vd tiêu đề đã viết hoa chữ đầu)
ALREADY_DONE = {
    ("De03_P1_Resume.docx", "pic_size"), ("De03_P3_Robotics.docx", "keep_next"),
    ("De04_P3_Tea.docx", "change_case"), ("De05_P4_Resume.docx", "keep_next"),
    ("De05_P6_Pharma.docx", "biblio_style"), ("De08_P2_Festival.docx", "keep_next"),
    ("De09_P3_Rock.docx", "repeat_header"),
}


def _tasks():
    out = []
    for folder in SETS:
        data = json.loads((folder / "de.json").read_text(encoding="utf-8"))
        for p in data["du_an"]:
            for t in p["nhiem_vu"]:
                spec = t["cham"]
                if spec["luat"] == "word_mau_de":
                    out.append(pytest.param(folder / p["file"], spec, id=f"{p['file']}-{spec['dang']}"))
    return out


TASKS = _tasks()


def test_sets_are_bundled():
    assert len(SETS) == 10
    for folder in SETS:
        exam, errors = custom.load_exam(folder)
        assert not errors, errors
        assert exam.standalone and len(exam.projects) == 7


def test_every_task_is_auto_graded():
    assert len(TASKS) == 380
    missing = {p.values[1]["dang"] for p in TASKS} - set(word_auto.GRADERS)
    assert not missing


@pytest.mark.parametrize("src,spec", TASKS)
def test_original_file_is_wrong(src, spec):
    ok = word_auto.grade(src, spec["dang"], spec["de_bai"], spec.get("goc"))
    assert ok is ((src.name, spec["dang"]) in ALREADY_DONE)


@pytest.mark.parametrize("src,spec", TASKS)
def test_simulated_solution_is_right(src, spec, tmp_path):
    if spec["dang"] not in SOL:
        pytest.skip("chưa có lời giải mô phỏng")
    work = tmp_path / src.name
    shutil.copyfile(src, work)
    solve(work, spec["dang"], spec["de_bai"])
    assert word_auto.grade(work, spec["dang"], spec["de_bai"], spec.get("goc")) is True


def test_external_saves(tmp_path, monkeypatch):
    """Quick Part và style set lưu vào thư mục người dùng (APPDATA)."""
    monkeypatch.setenv("APPDATA", str(tmp_path))
    src = SETS[0] / "De01_P1_Space.docx"
    qp = 'Save the heading as a Quick Part named “School Summary” in the Quick Part gallery.'
    ss = 'Save the current style set as a new style set named “RecycleStyleSet” in the default location.'
    assert word_auto.grade(src, "quickpart", qp) is False
    assert word_auto.grade(src, "custom_styleset", ss) is False
    bb = tmp_path / "Microsoft" / "Document Building Blocks" / "1033" / "16"
    bb.mkdir(parents=True)
    with zipfile.ZipFile(bb / "Building Blocks.dotx", "w") as z:
        z.writestr("word/glossary/document.xml", '<w:docPart><w:name w:val="School Summary"/></w:docPart>')
    (tmp_path / "Microsoft" / "QuickStyles").mkdir(parents=True)
    (tmp_path / "Microsoft" / "QuickStyles" / "RecycleStyleSet.dotx").write_bytes(b"x")
    assert word_auto.grade(src, "quickpart", qp) is True
    assert word_auto.grade(src, "custom_styleset", ss) is True


def test_unknown_type_is_self_check(tmp_path):
    src = SETS[0] / "De01_P1_Space.docx"
    assert word_auto.grade(src, "khong_co_dang_nay", "…") is None


# ====================================================================== nhập bộ đề


def _fake_set(root: Path) -> Path:
    """Dựng bộ đề nguồn giả từ một đề đã nhập (de.json → de_01.json + Files/)."""
    src = SETS[0]
    data = json.loads((src / "de.json").read_text(encoding="utf-8"))
    folder = root / "MOS_Word365_Thu" / "De_01"
    files = folder / "Files"
    files.mkdir(parents=True)
    entries = []
    for i, p in enumerate(data["du_an"][:2], start=1):
        shutil.copyfile(src / p["file"], files / p["file"])
        entries.append({"no": i, "file": p["file"], "theme": "Chu de", "tasks": [
            {"id": t["cham"].get("dang", "x"), "text": t["yeu_cau"], "steps": ["Bước 1", "Bước 2"]}
            for t in p["nhiem_vu"]] + [{"id": "dang_moi_la", "text": "Làm gì đó", "steps": ["…"]}]})
    for extra in data.get("file_phu", []):
        if (src / extra).is_file():
            shutil.copyfile(src / extra, files / extra)
    (folder / "de_01.json").write_text(json.dumps(entries, ensure_ascii=False), encoding="utf-8")
    return folder.parent


def test_import_folder_and_zip(tmp_path):
    source = _fake_set(tmp_path / "src")
    dest = tmp_path / "de_thi"
    done, warnings = nhap_de.import_path(source, dest)
    assert [d.name for d in done] == ["Word365_Thu_De_01"]
    assert any("dang_moi_la" in w for w in warnings)
    exam, errors = custom.load_exam(done[0])
    assert not errors and exam.standalone and exam.name == "Word – Đề 01 (Word365_Thu)"
    assert exam.projects[0].tasks[0].hint == "1) Bước 1  2) Bước 2"

    # nhập lại cùng bộ đề: ghi đè, không tạo thư mục mới
    zip_path = tmp_path / "bo_de.zip"
    with zipfile.ZipFile(zip_path, "w") as z:
        for f in source.rglob("*"):
            z.write(f, f.relative_to(source.parent))
    done2, _ = nhap_de.import_path(zip_path, dest)
    assert [d.name for d in done2] == ["Word365_Thu_De_01"]
    assert len(list(dest.iterdir())) == 1


def test_imported_exam_grading_counts_self_check_separately(tmp_path):
    source = _fake_set(tmp_path / "src")
    done, _ = nhap_de.import_path(source, tmp_path / "de_thi")
    exam, _ = custom.load_exam(done[0])
    report = grade(new_session(exam, "testing", tmp_path / "work"))
    n_self = sum(r.correct is None for r in report["results"])
    assert n_self == 2                                   # câu "dang_moi_la" ở 2 project
    assert report["total"] == len(report["results"]) - n_self
    assert report["correct"] == 0


def test_auto_import_only_when_changed(tmp_path):
    drop = tmp_path / "nhap_de"
    _fake_set(drop)
    dest = tmp_path / "de_thi"
    done, _ = nhap_de.auto_import(drop, dest)
    assert len(done) == 1
    again, _ = nhap_de.auto_import(drop, dest)
    assert again == []


def test_standalone_exams_not_merged():
    from mos.exams import ALL_EXAMS
    exams, errors = custom.load_custom_exams()
    merged = custom.merge_exams(ALL_EXAMS, exams)
    word = next(e for e in merged if e.code == "WORD" and not e.standalone)
    assert not any("Project" in p.name for p in word.projects)
    assert sum(e.standalone for e in merged) >= 10


def test_extra_files_copied_to_workdir(tmp_path):
    folder = next(f for f in SETS if "file_phu" in json.loads((f / "de.json").read_text(encoding="utf-8")))
    exam, _ = custom.load_exam(folder)
    session = new_session(exam, "training", tmp_path)
    extras = json.loads((folder / "de.json").read_text(encoding="utf-8"))["file_phu"]
    for name in extras:
        assert (session.workdir / name).exists()
