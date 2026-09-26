"""Nhập học viên từ CSV / Excel và phân quyền giáo viên."""
import pytest
from openpyxl import Workbook

from mos import importer
from mos.accounts import ADMIN, STUDENT, TEACHER, AccountStore


@pytest.fixture
def store(tmp_path):
    s = AccountStore(tmp_path / "tk.json")
    s.create("gv1", "Cô Lan", "1234", TEACHER)
    s.create("gv2", "Thầy Minh", "1234", TEACHER)
    s.create("binhtt", "Trần Thị Bình", "1234")          # đã tồn tại
    return s


def test_csv_with_vietnamese_headers_and_semicolons(store, tmp_path):
    f = tmp_path / "ds.csv"
    f.write_text("Họ và tên;Lớp;Hạn dùng\nNguyễn Văn An;10A1;31/12/2026\nNguyễn Văn An;10A2;\n\n",
                 encoding="utf-8-sig")
    entries = importer.plan(f, store)
    assert [e.username for e in entries] == ["nguyenvanan", "nguyenvanan2"]
    assert entries[0].lop == "10A1" and entries[0].han_dung == "2026-12-31"
    assert all(len(e.password) >= 6 and not e.error for e in entries)
    created, errors = importer.apply(entries, store, giao_vien="gv1")
    assert len(created) == 2 and not errors
    acc = store.authenticate("nguyenvanan", entries[0].password)
    assert acc.vai_tro == STUDENT and acc.giao_vien == "gv1" and acc.lop == "10A1" and acc.doi_mat_khau


def test_xlsx_with_errors(store, tmp_path):
    wb = Workbook()
    ws = wb.active
    ws.append(["Danh sách lớp 10A1"])                    # dòng tiêu đề phụ phía trên
    ws.append(["STT", "Tên đăng nhập", "Họ tên", "Mật khẩu"])
    ws.append([1, "BinhTT", "Trần Thị Bình", ""])         # trùng tài khoản có sẵn
    ws.append([2, "sai ten", "X", ""])                    # tên đăng nhập sai
    ws.append([3, "hv03", "Lê C", "12"])                  # mật khẩu ngắn
    ws.append([4, "hv04", "Phạm D", "abcd"])
    ws.append([5, "hv04", "Phạm E", "abcd"])              # trùng trong file
    f = tmp_path / "ds.xlsx"
    wb.save(f)
    entries = importer.plan(f, store)
    assert [bool(e.error) for e in entries] == [True, True, True, False, True]
    assert entries[0].row == 3
    created, errors = importer.apply(entries, store, giao_vien=None)
    assert [e.username for e in created] == ["hv04"] and len(errors) == 4


def test_template_roundtrip(store, tmp_path):
    for name in ("mau.xlsx", "mau.csv"):
        importer.write_template(tmp_path / name)
        entries = importer.plan(tmp_path / name, store)
        assert [e.username for e in entries] == ["nguyenvanan", "binhtt"]
        assert entries[1].error                          # binhtt đã tồn tại


def test_bad_files(store, tmp_path):
    (tmp_path / "a.txt").write_text("abc,def\n1,2\n", encoding="utf-8")
    with pytest.raises(importer.ImportFileError, match="tiêu đề"):
        importer.plan(tmp_path / "a.txt", store)
    (tmp_path / "a.pdf").write_bytes(b"%PDF")
    with pytest.raises(importer.ImportFileError, match="csv"):
        importer.plan(tmp_path / "a.pdf", store)


def test_teacher_permissions(store):
    gv1, gv2, admin = store.get("gv1"), store.get("gv2"), store.get("admin")
    store.create("hv_a", "A", "1234", giao_vien="gv1")
    store.create("hv_b", "B", "1234", giao_vien="gv2")
    visible = {a.username for a in store.visible_to(gv1)}
    assert visible == {"gv1", "hv_a"}
    assert store.can_manage(gv1, store.get("hv_a")) and not store.can_manage(gv1, store.get("hv_b"))
    assert not store.can_manage(gv1, gv2) and not store.can_manage(gv1, admin)
    assert store.can_manage(admin, gv2)
    assert store.assignable_roles(gv1) == [STUDENT]
    assert set(store.assignable_roles(admin)) == {ADMIN, TEACHER, STUDENT}
    assert len(store.visible_to(admin)) == len(store.accounts)
    student = store.get("hv_a")
    assert store.visible_to(student) == [student] and not student.is_staff and gv1.is_staff


def test_teacher_field_cleared_for_non_students(store):
    store.create("hv_c", "C", "1234", giao_vien="gv1")
    store.update("hv_c", vai_tro=TEACHER)
    assert store.get("hv_c").giao_vien is None
