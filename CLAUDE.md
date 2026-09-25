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
mos/chuong.py           chương theo khung MOS (CHAPTERS), count_tasks, practice_exam (chế độ "chapter")
mos/tai_lieu.py         bài giảng, "loai" slides/video/scorm; import_file theo đuôi: import_pptx (PowerPoint COM →
                        PNG, không có thì render_basic bằng Qt; ảnh mã hóa slides.mosl), import_video (video.mosv mã
                        hóa XOR), import_scorm (.zip/thư mục có imsmanifest.xml → scorm/, read_manifest lấy launch);
                        tiến độ tien_do_hoc.json, dữ liệu cmi.* scorm_hoc_vien.json
mos/lop_hoc.py          Lớp học trực tuyến (Supabase qua urllib): load_server/save_server (cau_hinh.json "may_chu" cạnh
                        app, hoặc ~/MOS_Practice/may_chu.json), Cloud (auth + PostgREST + RPC, tự refresh token,
                        CloudError/OfflineError thông báo tiếng Việt), hàng đợi cho_dong_bo.json (queue_result/
                        queue_progress/flush/flush_in_background), summarize, export_excel
mos/tu_khoa.py          thuật ngữ MOS song ngữ (_GLOSSARY), fold() bỏ dấu, search_terms/search_tasks/search_slides
mos/ooxml.py            đọc XML trong file .docx/.xlsx/.pptx (ZIP)
mos/exams/*.py          đề có sẵn: build(path) tạo file gốc + chk_*(path) -> bool
mos/qt/theme.py         hệ thống thiết kế Forest Canopy: màu, QSS, Card, chip, badge, table(), ScoreRing,
                        LevelBadge, stars, badge_tile, lang_switch, Confetti, hộp thoại
mos/qt/app.py           đăng nhập, Shell (sidebar + trang), Home, History, Result, ExamBar (thanh làm bài)
mos/qt/admin.py         trang Tài khoản / Đề thi / Kết quả, ExamEditor, RuleDialog, CheckReportDialog
mos/qt/hoc.py           LessonsPage, open_viewer → LessonViewer (slide) / VideoViewer (QtMultimedia, phát từ QBuffer) /
                        ScormViewer (QtWebEngine, profile ẩn danh chặn tải; API SCORM 1.2+2004 giả lập bằng JS,
                        gửi cmi về qua console.log "MOS_SCORM:"); practice_button = "Thực hành ngay"; ImportLessonDialog
mos/qt/tra_cuu.py       SearchPage (Tra từ khóa)
mos/qt/lop.py           ClassPage (giáo viên: lớp, mã lớp, bảng điểm; học sinh: vào lớp), OnlineAccountsPage (quản trị),
                        RegisterDialog, ServerDialog; err_text/busy/act cho thao tác mạng
may_chu/supabase_lop_hoc.sql  bảng ho_so/lop/thanh_vien/ket_qua/tien_do + RLS + RPC (vao_lop, tao_lop, lop_cua_toi,
                        luu_tien_do, dat_vai_tro, dat_khoa, dat_lai_mat_khau, xoa_tai_khoan…); chạy lại được
HUONG_DAN_MAY_CHU.md    hướng dẫn người dùng tạo Supabase
kiem_tra_de.py          kiểm tra đề bằng dòng lệnh
nhap_de.py / .bat       nhập bộ đề bằng dòng lệnh / kéo thả
de_thi/Excel_Mau/       đề mẫu (de.json + file gốc + dap_an/)
de_thi/Word365_DeThucTe_De_01..10/  bộ đề Word 365 đã nhập (luật word_mau_de)
tai_lieu/<bài>/          bài giảng đi kèm app (bai.json + slides.mosl); bài nhập thêm ở DATA_DIR/tai_lieu
tests/                  pytest
```
Dữ liệu người dùng (không nằm trong repo): `~/MOS_Practice/` → `tai_khoan.json`, `history.json`, `de_thi/`, `work/`,
`tai_lieu/`, `tien_do_hoc.json`, `scorm_hoc_vien.json`, `cho_dong_bo.json` (hàng đợi gửi lên máy chủ lớp học).
Chế độ Lớp học: tài khoản thật trên Supabase; `tai_khoan.json` giữ bản sao (Account.nguon="may_chu") để đăng nhập
lúc mất mạng. Vai trò: quan_tri / giao_vien (chỉ dùng khi có máy chủ) / hoc_vien.

## Lệnh
```bash
pip install -r requirements.txt pytest pyflakes
python -m pytest -q                 # phải pass hết
python -m pyflakes mos main.py kiem_tra_de.py
QT_QPA_PLATFORM=offscreen python ...   # chạy/chụp giao diện Qt không cần màn hình: widget.grab().save(...)
```
Trên Linux cloud có thể cần: `apt-get install -y libegl1 libgl1 libxkbcommon0 libfontconfig1 libpulse0 libnss3`
(QtWebEngine chạy offscreen dưới root cần `QTWEBENGINE_CHROMIUM_FLAGS=--no-sandbox`).

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
- Nhiệm vụ có chương: `Task(..., chapter=N)` (đề có sẵn) hoặc `"chuong": N` trong de.json; bộ đề nhập lấy chương
  từ domain. Chương mới / đổi tên chương sửa trong `mos/chuong.py`.
- Thuật ngữ tra cứu mới: thêm dòng vào `_GLOSSARY` trong `mos/tu_khoa.py` (Việt + Anh + đường dẫn menu).
- **Lớp học trực tuyến**: đổi cấu trúc máy chủ thì sửa `may_chu/supabase_lop_hoc.sql` (phải chạy lại được nhiều lần,
  quyền qua RLS/RPC, không tin client) + `tests/test_lop_hoc.py`. Test tích hợp chạy trên Postgres + PostgREST thật
  với auth giả (`tests/supabase_gia.py`), cần `MOS_TEST_PG` + `POSTGREST_BIN`; thiếu thì tự bỏ qua. Trên cloud:
  `apt-get install -y postgresql-16 && service postgresql start`, đặt mật khẩu postgres, tải PostgREST (github
  releases, bản linux-static-x64). Không bao giờ đưa khóa service_role vào app.
- Hàm chấm trả `None` = không chấm tự động (hiện "Tự kiểm tra", không tính điểm).
