"""Kiểm tra luật chấm (mos/rules.py) và bộ nạp đề tự soạn (mos/custom.py)."""
import json
import shutil
from pathlib import Path

import pytest

import kiem_tra_de
from mos import rules
from mos.core import new_session
from mos.custom import load_exam
from mos.exams import ALL_EXAMS
from tests.test_exams import SOLUTIONS

ROOT = Path(__file__).resolve().parents[1]

# (file, luật) – luật phải SAI trên file gốc và ĐÚNG trên file đã giải
CASES = [
    ("DoanhSo.xlsx", {"luat": "excel_ten_sheet", "ten": "DoanhSo"}),
    ("DoanhSo.xlsx", {"luat": "excel_khong_co_sheet", "ten": "Sheet1"}),
    ("DoanhSo.xlsx", {"luat": "excel_cong_thuc", "sheet": "DoanhSo", "o": "E2:E11", "chua": ["C{hang}*D{hang}"]}),
    ("DoanhSo.xlsx", {"luat": "excel_co_bang", "sheet": "DoanhSo", "vung": "A1:E11"}),
    ("DoanhSo.xlsx", {"luat": "excel_co_dinh", "sheet": "DoanhSo", "o": "A2"}),
    ("DoanhSo.xlsx", {"luat": "excel_dinh_dang_so", "sheet": "DoanhSo", "o": "D2:E11", "chua": "#,##0"}),
    ("DoanhSo.xlsx", {"luat": "excel_gia_tri", "sheet": "DoanhSo", "o": "G1", "bang": "tổng doanh thu"}),
    ("NhanVien.xlsx", {"luat": "excel_dinh_dang_dieu_kien", "vung": "C2:C11"}),
    ("NhanVien.xlsx", {"luat": "excel_ten_vung", "ten": "LuongNV", "vung": "C2:C11"}),
    ("NhanVien.xlsx", {"luat": "excel_bieu_do", "loai": "cot"}),
    ("NhanVien.xlsx", {"luat": "excel_huong_trang", "huong": "ngang"}),
    ("NhanVien.xlsx", {"luat": "excel_sap_xep", "cot": "C", "tu": 2, "den": 11, "chieu": "giam"}),
    ("BaoCao.docx", {"luat": "word_kieu_doan", "doan": ["Giới thiệu", "Kế hoạch năm tới"], "kieu": "Heading 1"}),
    ("BaoCao.docx", {"luat": "word_khong_co_chu", "chu": "ABC"}),
    ("BaoCao.docx", {"luat": "word_co_chu", "chu": "công ty XYZ", "so_lan": 3}),
    ("BaoCao.docx", {"luat": "word_muc_luc"}),
    ("BaoCao.docx", {"luat": "word_so_trang", "vi_tri": "chan_trang"}),
    ("BaoCao.docx", {"luat": "word_watermark", "chu": "BẢO MẬT"}),
    ("BaoCao.docx", {"luat": "word_huong_trang", "huong": "ngang"}),
    ("ThuMoi.docx", {"luat": "word_can_le", "doan": "Thư mời hội thảo", "can": "giua"}),
    ("ThuMoi.docx", {"luat": "word_dau_dong", "doan": ["Khai mạc", "Báo cáo chuyên đề"]}),
    ("ThuMoi.docx", {"luat": "word_bang", "so_cot": 3, "so_hang": 4, "o_dau": "Thời gian"}),
    ("ThuMoi.docx", {"luat": "word_footnote", "chu": "đối tác đính kèm"}),
    ("ThuMoi.docx", {"luat": "word_thuoc_tinh", "truong": "title", "gia_tri": "Thư mời hội thảo"}),
    ("ThuMoi.docx", {"luat": "word_theo_doi_thay_doi"}),
    ("GioiThieu.pptx", {"luat": "ppt_ghi_chu", "slide": 1, "chu": "chào mừng"}),
    ("GioiThieu.pptx", {"luat": "ppt_chuyen_trang"}),
    ("GioiThieu.pptx", {"luat": "ppt_an_slide", "tieu_de": "Giá bán"}),
    ("GioiThieu.pptx", {"luat": "ppt_hieu_ung", "slide": 1}),
    ("GioiThieu.pptx", {"luat": "ppt_kich_thuoc", "ti_le": "4:3"}),
    ("GioiThieu.pptx", {"luat": "ppt_co_slide", "tieu_de": "Kế hoạch", "bo_cuc": "Title and Content", "vi_tri": "cuoi"}),
    ("BaoCaoQuy.pptx", {"luat": "ppt_bang", "tieu_de": "Doanh thu theo khu vực", "so_cot": 3, "so_hang": 4}),
    ("BaoCaoQuy.pptx", {"luat": "ppt_smartart", "tieu_de": "Quy trình bán hàng"}),
    ("BaoCaoQuy.pptx", {"luat": "ppt_khong_co_slide", "tieu_de": "Xóa slide này"}),
    ("BaoCaoQuy.pptx", {"luat": "ppt_so_slide", "so": 4}),
    ("BaoCaoQuy.pptx", {"luat": "ppt_section", "ten": "Mở đầu"}),
    ("BaoCaoQuy.pptx", {"luat": "ppt_so_trang"}),
]


@pytest.fixture(scope="module")
def files(tmp_path_factory):
    """Sinh file gốc + file đã giải cho mọi dự án có sẵn."""
    fresh, solved = {}, {}
    for exam in ALL_EXAMS:
        s1 = new_session(exam, "testing", tmp_path_factory.mktemp("fresh"))
        s2 = new_session(exam, "testing", tmp_path_factory.mktemp("solved"))
        for i, p in enumerate(exam.projects):
            SOLUTIONS[p.filename](s2.file_of(i))
            fresh[p.filename], solved[p.filename] = s1.file_of(i), s2.file_of(i)
    return fresh, solved


@pytest.mark.parametrize("filename,spec", CASES, ids=[c[1]["luat"] for c in CASES])
def test_rule(files, filename, spec):
    fresh, solved = files
    assert rules.validate(spec) is None
    check = rules.make_check(spec)
    try:
        before = check(fresh[filename])
    except KeyError:   # vd sheet "DoanhSo" chưa tồn tại trong file gốc
        before = False
    assert before is False
    assert check(solved[filename]) is True


def test_every_rule_documented():
    for name, fn in rules.RULES.items():
        assert fn.__doc__, name
    assert "excel_cong_thuc" in rules.describe()


def test_sample_exam_passes_checker():
    assert kiem_tra_de.main([str(ROOT / "de_thi" / "Excel_Mau")]) == 0


def _write(folder: Path, data: dict, with_file=True):
    folder.mkdir()
    (folder / "de.json").write_text(json.dumps(data, ensure_ascii=False), encoding="utf-8")
    if with_file:
        shutil.copy(ROOT / "de_thi" / "Excel_Mau" / "BangDiem.xlsx", folder / "BangDiem.xlsx")


def _exam(cham):
    return {"mon": "EXCEL", "ten": "T", "du_an": [
        {"ten": "P", "file": "BangDiem.xlsx", "nhiem_vu": [{"yeu_cau": "x", "cham": cham}]}]}


def test_loader_ok(tmp_path):
    _write(tmp_path / "ok", _exam({"luat": "excel_co_dinh"}))
    exam, errors = load_exam(tmp_path / "ok")
    assert not errors and exam.code == "EXCEL" and len(exam.projects[0].tasks) == 1
    dest = tmp_path / "copy.xlsx"
    exam.projects[0].build(dest)
    assert dest.read_bytes() == (tmp_path / "ok" / "BangDiem.xlsx").read_bytes()


@pytest.mark.parametrize("cham,msg", [
    ({"luat": "khong_ton_tai"}, "không tồn tại"),
    ({"luat": "excel_ten_sheet"}, "excel_ten_sheet"),                    # thiếu tham số
    ({"luat": "excel_co_dinh", "sai_ten": 1}, "excel_co_dinh"),         # tham số lạ
    ("excel_co_dinh", "đối tượng"),
])
def test_loader_reports_bad_rules(tmp_path, cham, msg):
    _write(tmp_path / "bad", _exam(cham))
    exam, errors = load_exam(tmp_path / "bad")
    assert exam is None and any(msg in e for e in errors), errors


def test_loader_reports_missing_file_and_json(tmp_path):
    _write(tmp_path / "nofile", _exam({"luat": "excel_co_dinh"}), with_file=False)
    assert any("không thấy file gốc" in e for e in load_exam(tmp_path / "nofile")[1])
    (tmp_path / "badjson").mkdir()
    (tmp_path / "badjson" / "de.json").write_text("{ sai", encoding="utf-8")
    assert any("JSON" in e for e in load_exam(tmp_path / "badjson")[1])


def test_custom_projects_merge_into_same_subject(tmp_path):
    from mos.custom import merge_exams
    _write(tmp_path / "a", _exam({"luat": "excel_co_dinh"}))
    data = _exam({"luat": "excel_co_dinh"})
    data["du_an"][0]["ten"] = "Dự án 1 – Bảng điểm"
    _write(tmp_path / "b", data)
    custom = [load_exam(tmp_path / "a")[0], load_exam(tmp_path / "b")[0]]
    merged = merge_exams(ALL_EXAMS, custom)

    assert [e.code for e in merged] == [e.code for e in ALL_EXAMS]   # không thêm thẻ riêng
    excel = next(e for e in merged if e.code == "EXCEL")
    assert [p.name for p in excel.projects] == [
        "Dự án 1 – Doanh số", "Dự án 2 – Nhân viên", "Dự án 3 – P", "Dự án 4 – Bảng điểm"]
    assert [p.filename for p in excel.projects][2:] == ["BangDiem.xlsx", "BangDiem_2.xlsx"]
    word = next(e for e in merged if e.code == "WORD")
    assert len(word.projects) == 2

    session = new_session(excel, "testing", tmp_path / "work")
    assert all(session.file_of(i).is_file() for i in range(4))
