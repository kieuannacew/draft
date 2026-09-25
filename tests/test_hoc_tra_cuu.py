"""Luyện theo chương, tài liệu học (bài giảng PPTX) và tra từ khóa."""
import os
import random

import pytest

from mos import chuong, tai_lieu, tu_khoa
from mos.custom import load_custom_exams, merge_exams
from mos.exams import ALL_EXAMS

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")


@pytest.fixture(scope="module")
def exams():
    custom, _ = load_custom_exams()
    return merge_exams(ALL_EXAMS, custom)


# ---------------------------------------------------------------- chương
def test_moi_mon_co_chuong():
    for code in ("WORD", "EXCEL", "POWERPOINT"):
        assert chuong.chapters(code)
        assert chuong.chapter(code, 1).number == 1
    assert chuong.chapter("WORD", 99) is None


def test_moi_cau_de_co_san_deu_co_chuong():
    for exam in ALL_EXAMS:
        for p in exam.projects:
            for t in p.tasks:
                assert chuong.chapter(exam.code, t.chapter), (exam.name, t.title)


def test_bai_luyen_theo_chuong(exams):
    counts = chuong.count_tasks(exams, "WORD")
    assert counts and all(n > 0 for n in counts.values())
    for number in counts:
        exam = chuong.practice_exam(exams, "WORD", number, rng=random.Random(1))
        assert exam is not None and exam.minutes == 0
        assert exam.name.startswith(f"Chương {number}:")
        names = [p.filename.casefold() for p in exam.projects]
        assert len(names) == len(set(names)), "tên file trùng sẽ ghi đè nhau"
        assert len(exam.projects) <= 6
        for p in exam.projects:
            assert p.tasks and all(t.chapter == number for t in p.tasks)
    assert chuong.practice_exam(exams, "WORD", 99) is None


# ---------------------------------------------------------------- tài liệu học
def _make_pptx(path):
    from pptx import Presentation
    from pptx.util import Inches
    prs = Presentation()
    s = prs.slides.add_slide(prs.slide_layouts[1])
    s.shapes.title.text = "Chèn hình mờ"
    s.placeholders[1].text = "Design → Watermark → Custom Watermark"
    s.notes_slide.notes_text_frame.text = "Ghi chú cho giáo viên"
    s = prs.slides.add_slide(prs.slide_layouts[5])
    s.shapes.title.text = "Bảng"
    s.shapes.add_table(2, 2, Inches(1), Inches(2), Inches(4), Inches(1)).table.cell(0, 0).text = "Ô A1"
    prs.save(path)


def test_dong_goi_anh_khong_doc_truc_tiep(tmp_path):
    imgs = [b"\x89PNG anh 1", b"\x89PNG anh 2"]
    tai_lieu.pack_slides(tmp_path, "bai-1", imgs)
    raw = (tmp_path / tai_lieu.PACK).read_bytes()
    assert b"PNG anh" not in raw
    assert tai_lieu.read_slide(tmp_path, "bai-1", 2) == imgs[1]


def test_nhap_pptx_va_tien_do(tmp_path, monkeypatch):
    pytest.importorskip("PySide6")
    src = tmp_path / "bai.pptx"
    _make_pptx(src)
    base = tmp_path / "tai_lieu"
    lesson = tai_lieu.import_pptx(src, "Hình mờ", "word", chuong=2, dest_base=base, use_powerpoint=False)
    assert lesson.count == 2 and lesson.mon == "WORD" and lesson.chuong == 2
    assert "Watermark" in lesson.slides[0]["text"]
    assert lesson.slides[0]["notes"] == "Ghi chú cho giáo viên"
    assert "Ô A1" in lesson.slides[1]["text"]
    assert lesson.image(1).startswith(b"\x89PNG")
    assert not list(base.rglob("*.pptx")), "không được chép file pptx gốc"
    # tìm thấy trong danh sách, tra cứu được slide
    monkeypatch.setattr(tai_lieu, "search_dirs", lambda: [base])
    assert [x.id for x in tai_lieu.list_lessons()] == [lesson.id]
    hits = tu_khoa.search_slides("watermark", [lesson])
    assert hits and hits[0].number == 1
    # tiến độ chỉ tăng
    monkeypatch.setattr(tai_lieu, "DATA_DIR", tmp_path)
    tai_lieu.save_progress("hv", lesson.id, 2)
    tai_lieu.save_progress("hv", lesson.id, 1)
    assert tai_lieu.load_progress("hv") == {lesson.id: 2}
    assert tai_lieu.load_progress("khac") == {}
    tai_lieu.delete_lesson(lesson)
    assert not lesson.folder.exists()


def test_bai_giang_mau_doc_duoc():
    lessons = [x for x in tai_lieu.list_lessons() if x.folder.parent.name == "tai_lieu"]
    assert lessons
    for x in lessons:
        assert x.count > 0 and x.image(1).startswith(b"\x89PNG")


# ---------------------------------------------------------------- tra từ khóa
def test_fold_bo_dau():
    assert tu_khoa.fold("Mục Lục") == tu_khoa.fold("muc luc")
    assert tu_khoa.fold("Đường viền") == "duong vien"


@pytest.mark.parametrize("q", ["watermark", "hình mờ", "hinh mo", "Freeze Panes", "mục lục"])
def test_tra_thuat_ngu(q):
    assert tu_khoa.search_terms(q)


def test_tra_theo_mon():
    terms = tu_khoa.search_terms("freeze", "EXCEL")
    assert terms and all(t.subject in ("EXCEL", "ALL") for t in terms)
    assert not tu_khoa.search_terms("freeze", "POWERPOINT")


def test_tra_cau_trong_de(exams):
    hits = tu_khoa.search_tasks("watermark", exams)
    assert hits
    assert all(h.exam.code == "WORD" for h in hits)
    assert tu_khoa.search_tasks("zzzkhongcozzz", exams) == []
