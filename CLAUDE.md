# CLAUDE.md – Luyện thi MOS

App desktop luyện thi Microsoft Office Specialist (Word MO-100, Excel MO-200, PowerPoint MO-300) kiểu GMetrix:
học viên làm bài trên Office thật, app đọc file đã lưu để chấm (thang 1000, đạt ≥ 700).
Có tài khoản (quản trị / học viên) và form soạn đề trong app.

## Cách làm việc với người dùng
- Người dùng không phải lập trình viên: trả lời **bằng tiếng Việt, ngắn gọn, dễ hiểu**.
- Hạn chế chụp màn hình để tự kiểm tra; chỉ chụp khi đổi giao diện và cần cho người dùng xem.
- Máy người dùng: **Windows + Python 3.14 (lệnh `py`) + Microsoft Office**. Môi trường cloud là Linux, không có
  Office → không thử được bước làm bài thật; nói rõ điều này khi báo kết quả.
- Làm xong: chạy test, commit, push. Người dùng tải bằng **Code → Download ZIP** rồi chạy `run.bat`.

## Cấu trúc
```
main.py                 điểm khởi động (gọi mos.qt.app.main)
mos/core.py             Exam / Project / Task, chấm điểm, lịch sử, DATA_DIR (cau_hinh.json → thư mục dữ liệu chung)
mos/accounts.py         tài khoản: PBKDF2, vai trò quan_tri / giao_vien / hoc_vien, lop, giao_vien (chủ quản);
                        visible_to / can_manage / assignable_roles = phân quyền; admin/admin lần đầu
mos/importer.py         nhập học viên từ CSV/XLSX: plan() xem trước → apply(); write_template()
mos/rules.py            46 luật chấm (@rule) + RULE_TITLES + PARAM_UI (dựng form soạn đề)
mos/custom.py           đề tự soạn de_thi/<tên>/de.json; merge_exams gộp vào bài thi cùng môn; check_exam
mos/ooxml.py            đọc XML trong file .docx/.xlsx/.pptx (ZIP)
mos/exams/*.py          đề có sẵn: build(path) tạo file gốc + chk_*(path) -> bool
mos/qt/theme.py         hệ thống thiết kế: màu, QSS, Card, chip, badge, table(), ScoreRing, hộp thoại
mos/qt/app.py           đăng nhập, Shell (sidebar + trang), Home, History, Result, ExamBar (thanh làm bài)
mos/qt/admin.py         trang Tài khoản / Đề thi / Kết quả, ImportDialog, ExamEditor, RuleDialog, CheckReportDialog
mos/qt/winlayout.py     Windows (ctypes): AppBar gắn thanh đề ở đáy, OfficeFitter thu cửa sổ Office vừa phần trống
kiem_tra_de.py          kiểm tra đề bằng dòng lệnh
de_thi/Excel_Mau/       đề mẫu (de.json + file gốc + dap_an/)
tests/                  pytest
```
Dữ liệu người dùng (không nằm trong repo): `~/MOS_Practice/` → `tai_khoan.json`, `history.json`, `de_thi/`, `work/`.

## Lệnh
```bash
pip install -r requirements.txt pytest pyflakes
python -m pytest -q                 # phải pass hết
python -m pyflakes mos main.py kiem_tra_de.py
QT_QPA_PLATFORM=offscreen python ...   # chạy/chụp giao diện Qt không cần màn hình: widget.grab().save(...)
```
Trên Linux cloud có thể cần: `apt-get install -y libegl1 libgl1 libxkbcommon0 libfontconfig1`.

## Quy ước
- Chữ trên giao diện, docstring, thông báo lỗi: **tiếng Việt**. Tên biến/hàm theo code hiện có.
- Giao diện chỉ dùng thành phần trong `mos/qt/theme.py` (button(kind=primary|ghost|danger), Card, chip, table,
  info/warn/error/confirm). Không dùng Tkinter.
- Qt: ký tự `&` trong chữ của nút phải viết `&&`.
- **Thêm luật chấm**: hàm `@rule` trong `rules.py` (có docstring) + tên trong `RULE_TITLES` + nhãn tham số mới trong
  `PARAM_UI`; thêm ca kiểm tra vào `tests/test_custom.py` (file gốc phải SAI, file đã giải phải ĐÚNG).
- **Thêm đề có sẵn**: `build_*` + `chk_*` trong `mos/exams/`, lời giải mô phỏng trong `tests/test_exams.py` (`SOLUTIONS`).
- Hàm chấm phải chịu được cách Office thật lưu file (tiền tố XML khác nhau, shared formula, style tên tiếng Anh…);
  khi không chắc, chấm lỏng phần cốt lõi thay vì so khớp nguyên văn.
- Không sửa `de_thi/Excel_Mau` khi chạy thử (dùng thư mục tạm / HOME tạm).
- Phân quyền: trang quản lý chỉ cho `is_staff`; giáo viên chỉ thấy/sửa học viên có `giao_vien` = mình
  (luôn dùng `store.visible_to` / `store.can_manage`, không tự lọc).
- Code Windows-only (`winlayout.py`, `core.unblock`) không chạy được trên cloud: kiểm bằng API giả lập
  (`tests/test_winlayout.py`) và nói rõ với người dùng là chưa thử trên Windows thật.
