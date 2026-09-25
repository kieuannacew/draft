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
mos/accounts.py         tài khoản: PBKDF2, vai trò quan_tri / hoc_vien, khóa, hạn dùng; admin/admin lần đầu
mos/rules.py            46 luật chấm (@rule) + RULE_TITLES + PARAM_UI (dựng form soạn đề)
mos/custom.py           đề tự soạn de_thi/<tên>/de.json; merge_exams gộp vào bài thi cùng môn ("rieng": đề riêng;
                        "file_phu": file phụ chép cùng file làm bài); check_exam
mos/nhap_de.py          nhập bộ đề ngoài (De_xx/de_xx.json + Files/, thư mục hoặc .zip) → de_thi/<tên>/de.json
mos/word_auto.py        chấm tự động ~180 dạng câu Word của bộ đề nhập (GRADERS theo mã dạng, SNAPS lưu info file gốc)
mos/i18n.py             song ngữ: tr("chữ Việt") tra mos/i18n_en.json; pick(vi, en) cho nội dung đề; lưu ngôn ngữ
                        trong ~/MOS_Practice/cai_dat.json
mos/dich_de.json        bộ nhớ dịch cho bộ đề nhập (đề bài Anh→Việt, gợi ý Việt→Anh)
mos/gamify.py           XP, cấp độ, chuỗi ngày, huy hiệu, sao, mẹo mỗi ngày (tính từ history)
mos/ooxml.py            đọc XML trong file .docx/.xlsx/.pptx (ZIP)
mos/exams/*.py          đề có sẵn: build(path) tạo file gốc + chk_*(path) -> bool
mos/qt/theme.py         hệ thống thiết kế Forest Canopy: màu, QSS, Card, chip, badge, table(), ScoreRing,
                        LevelBadge, stars, badge_tile, lang_switch, Confetti, hộp thoại
mos/qt/app.py           đăng nhập, Shell (sidebar + trang), Home, History, Result, ExamBar (thanh làm bài)
mos/qt/admin.py         trang Tài khoản / Đề thi / Kết quả, ExamEditor, RuleDialog, CheckReportDialog
kiem_tra_de.py          kiểm tra đề bằng dòng lệnh
nhap_de.py / .bat       nhập bộ đề bằng dòng lệnh / kéo thả
de_thi/Excel_Mau/       đề mẫu (de.json + file gốc + dap_an/)
de_thi/Word365_DeThucTe_De_01..10/  bộ đề Word 365 đã nhập (luật word_mau_de)
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
- **Song ngữ**: mọi chữ hiện trên giao diện bọc `tr("…")` (có biến: `tr("Đã làm {n}").format(n=…)`, không dùng
  f-string bên trong tr) và thêm bản Anh vào `mos/i18n_en.json` (test_i18n_gamify kiểm tra đủ khóa). Nội dung đề có
  cặp trường `yeu_cau`/`yeu_cau_en`, `goi_y`/`goi_y_en`, `ten`/`ten_en`, `mo_ta`/`mo_ta_en`; đề có sẵn dùng
  `Task(title, hint, check, title_en, hint_en)`. Chữ trong ngoặc “…” (nội dung file Office) giữ nguyên khi dịch.
- Màu chỉ lấy từ hằng số trong theme.py (FOREST, SAGE, OLIVE, IVORY, GOLD…); không tự đặt mã màu mới rải rác.
- Giao diện chỉ dùng thành phần trong `mos/qt/theme.py` (button(kind=primary|ghost|danger), Card, chip, table,
  info/warn/error/confirm). Không dùng Tkinter.
- Qt: ký tự `&` trong chữ của nút phải viết `&&`.
- **Thêm luật chấm**: hàm `@rule` trong `rules.py` (có docstring) + tên trong `RULE_TITLES` + nhãn tham số mới trong
  `PARAM_UI`; thêm ca kiểm tra vào `tests/test_custom.py` (file gốc phải SAI, file đã giải phải ĐÚNG).
- **Thêm đề có sẵn**: `build_*` + `chk_*` trong `mos/exams/`, lời giải mô phỏng trong `tests/test_exams.py` (`SOLUTIONS`).
- Hàm chấm phải chịu được cách Office thật lưu file (tiền tố XML khác nhau, shared formula, style tên tiếng Anh…);
  khi không chắc, chấm lỏng phần cốt lõi thay vì so khớp nguyên văn.
- Không sửa `de_thi/Excel_Mau` khi chạy thử (dùng thư mục tạm / HOME tạm).
- **Dạng câu Word mới cho bộ đề nhập**: `@grader("ma_dang")` trong `word_auto.py` (đọc thông số từ đề bài; cần
  thông tin file gốc thì thêm `@snap`), lời giải mô phỏng trong `tests/word_solutions.py`. `tests/test_word_auto.py`
  chạy MỌI câu của de_thi/Word365_*: file gốc phải SAI (trừ `ALREADY_DONE`), lời giải phải ĐÚNG.
- Hàm chấm trả `None` = không chấm tự động (hiện "Tự kiểm tra", không tính điểm).
