# Luyện thi MOS (kiểu GMetrix)

Phần mềm luyện thi **Microsoft Office Specialist** cho Word (MO-100), Excel (MO-200), PowerPoint (MO-300).
Giống GMetrix: học viên **làm bài trực tiếp trên Word/Excel/PowerPoint thật**, phần mềm tự đọc file để chấm điểm
theo thang 1000 (đạt từ 700). Có **tài khoản** cho giáo viên / học viên và **soạn đề ngay trong app**.
Giao diện dùng **Qt (PySide6)**: sắc nét trên màn hình độ phân giải cao, một cửa sổ với thanh menu bên trái.

## Cách chạy (Windows)

1. Cài [Python 3.10+](https://www.python.org/downloads/) (khi cài nhớ tick **Add Python to PATH**).
2. Tải repo này về, giải nén.
3. Nhấp đúp **`run.bat`** (lần đầu sẽ tự cài thư viện).

Muốn có file `.exe` để chép sang máy khác: nhấp đúp **`build_exe.bat`** → file nằm ở `dist\LuyenThiMOS.exe`.

## Tài khoản

| Vai trò | Được làm gì |
|---|---|
| **Quản trị** (giáo viên) | Làm bài, **tạo / khóa / xóa tài khoản**, đặt lại mật khẩu, **soạn đề**, **xem kết quả của mọi học viên** |
| **Học viên** | Làm bài, xem lịch sử của chính mình, đổi mật khẩu |

**Lần đầu chạy:** đăng nhập bằng `admin` / `admin`. App sẽ bắt đổi mật khẩu ngay lần đăng nhập này.

Cấp tài khoản cho học viên: vào mục **Tài khoản** ở thanh menu bên trái (chỉ tài khoản quản trị mới thấy), rồi chọn:
- **Thêm tài khoản**: tạo từng người. Mật khẩu được tạo ngẫu nhiên, và có thể đặt **hạn dùng** (ví dụ `2026-12-31`).
- **Tạo cho cả lớp**: dán danh sách, mỗi dòng dạng `tên_đăng_nhập, Họ tên`. App tạo tài khoản hàng loạt và cho
  **lưu danh sách tên đăng nhập / mật khẩu ra file CSV** để phát cho học viên.
- **Đặt lại mật khẩu**, **Khóa / Mở khóa**, **Xóa**.

Mục **Kết quả học viên** hiện mọi lượt làm bài. Có thể lọc theo học viên và **xuất CSV** để mở bằng Excel.

> Mật khẩu được lưu dạng mã băm (PBKDF2), không lưu mật khẩu thật. Dữ liệu nằm trong
> `C:\Users\<tên>\MOS_Practice\` (`tai_khoan.json`, `history.json`, `de_thi\`).
>
> **Phòng máy dùng chung dữ liệu:** tạo file `cau_hinh.json` cạnh `main.py` (hoặc cạnh `LuyenThiMOS.exe`):
> `{"thu_muc_du_lieu": "\\\\MAYCHU\\MOS"}`. Khi đó mọi máy dùng chung tài khoản, đề và kết quả trong thư mục mạng đó.

## Giao diện, song ngữ Việt / Anh và "vui học"

- Giao diện theo chủ đề màu **Forest Canopy** (xanh rừng, xanh ô liu, nền ngà; font Be Vietnam Pro đóng gói kèm app, hiển thị tiếng Việt chuẩn).
- Nút **VI | EN** ở màn đăng nhập, thanh bên và thanh làm bài: đổi toàn bộ giao diện **và cả đề, gợi ý / đáp án**
  sang tiếng Anh hoặc tiếng Việt (đang làm bài cũng đổi được – tiện đối chiếu với đề tiếng Anh như thi thật).
  App nhớ ngôn ngữ đã chọn cho lần mở sau.
- **Vui học**: mỗi lần nộp bài được **XP** (điểm/10, +50 nếu đạt, +100 nếu 1000 điểm, thi thử ×1.2), lên **cấp**
  (Mầm non → Rừng già), **chuỗi ngày học** 🔥, 9 **huy hiệu** để mở khóa, **sao** ★★★ cho từng đề, **mẹo mỗi ngày**,
  pháo giấy khi đạt bài.
- Đề tự soạn: trong form Soạn đề có thêm ô **bản tiếng Anh** cho yêu cầu và gợi ý (tuỳ chọn).

## Nhập bộ đề có sẵn (tự động)

App đã có sẵn **10 đề Word 365 "Đề thực tế"** (mỗi đề 7 project, 38 câu, tiếng Anh như thi thật) trong
`de_thi\Word365_DeThucTe_De_01 … _10`. Mỗi đề là một bài thi riêng trên Trang chủ, **chấm tự động**.

Muốn thêm bộ đề khác cùng dạng (thư mục `De_xx` có `de_xx.json` + thư mục `Files`), chọn một trong ba cách:
- Trong app: **Đề thi → Nhập bộ đề từ thư mục…** hoặc **Nhập từ file ZIP…**
- Kéo thả thư mục bộ đề (hoặc file `.zip`) vào **`nhap_de.bat`**.
- Chép bộ đề vào thư mục **`nhap_de`** cạnh `run.bat`: app tự nhập khi tài khoản quản trị mở Trang chủ.

Ảnh / file dữ liệu trộn thư / thư mục `3D Models` được chép cùng file làm bài. Câu nào chưa chấm tự động được
sẽ hiện **"Tự kiểm tra"** và không tính vào điểm. Các câu "Save a copy … in your Documents folder" được chấm bằng
cách tìm file trong thư mục **Documents** của máy (hoặc trong thư mục bài làm).

## Soạn đề ngay trong app (không cần lập trình)

Vào mục **Đề thi** ở thanh menu bên trái, bấm **Soạn đề mới**, rồi làm lần lượt:

1. Đặt **tên đề** và chọn **môn** (Word / Excel / PowerPoint).
2. **① Dự án**: bấm *Thêm*, rồi chọn **file gốc** (file chưa làm, soạn sẵn bằng Office).
3. **② Nhiệm vụ**: bấm *Thêm*, ghi **yêu cầu** (học viên nhìn thấy) và **gợi ý cách làm**.
4. **③ Luật chấm**: bấm *Thêm luật*, chọn điều cần kiểm tra trong danh sách (ví dụ *Cố định hàng/cột (Freeze Panes)*),
   rồi điền các ô app hiện ra. Có 46 luật cho Word, Excel và PowerPoint.
5. Chọn **file đáp án** (file đã làm đúng hết), rồi bấm **Lưu & kiểm tra**. App sẽ báo:
   - file gốc có bị chấm **sai** không (đúng như mong đợi),
   - file đáp án có được chấm **đúng** không.

Đề đã lưu được **gộp vào bài thi của môn tương ứng**. Ví dụ đề Excel mới sẽ thành Dự án 3, 4… của bài Excel.

Chi tiết từng luật chấm, cùng cách soạn đề bằng file `de.json` cho người muốn làm tay:
[HUONG_DAN_SOAN_DE.md](HUONG_DAN_SOAN_DE.md).

## Chọn bài thi, luyện theo chương

Trang chủ chỉ có **3 môn: Word, Excel, PowerPoint**. Bấm vào một môn → chọn chế độ:
- **Luyện tập** / **Thi thử**: chọn một đề trong môn đó (các đề nằm bên trong môn, không bày hết ra ngoài).
- **Luyện theo chương** 📚: chọn một chương theo khung đề MOS (vd Word: 1 Quản lý tài liệu, 2 Chữ – đoạn – section,
  3 Bảng & danh sách, 4 Tham chiếu, 5 Đồ họa, 6 Cộng tác, 7 Nâng cao), rồi chọn **đề riêng của chương** hoặc
  **Trộn ngẫu nhiên** (gom câu của chương từ mọi đề, tối đa 6 dự án, mỗi lần một khác). Không giới hạn giờ, có gợi ý.
- Word có sẵn **14 đề theo chương** (211 câu, `de_thi\Word365_TheoChuong_Chuong_<c>_De_<k>`) sinh bằng bộ công cụ
  sinh đề Mít tin học, phủ **đủ 183 dạng câu**, tất cả chấm tự động, song ngữ Việt/Anh.
- Đề tự soạn: trong form Soạn đề, mỗi nhiệm vụ có ô **Chương** để xếp vào chương.

## Tài liệu học: slide, video ngắn, bài tương tác SCORM → thực hành ngay

Mục **📖 Tài liệu học** ở thanh bên. Quản trị bấm **Thêm bài giảng…**, chọn file, môn và chương:

| Loại | File | Học viên thấy |
|---|---|---|
| 🖼 Slide | `.pptx` | từng slide, ảnh thu nhỏ, phím ← →, ghi chú giáo viên (máy có PowerPoint: giống hệt bản gốc) |
| 🎬 Video ngắn | `.mp4`, `.webm`, `.mov`… | trình phát trong app: phát/dừng, tua, tốc độ 0.75–1.5× |
| 🧩 Tương tác | gói **SCORM** `.zip` (iSpring, Articulate, Captivate, H5P…) | bài tương tác chạy trong app; app ghi **hoàn thành** và **điểm** |

- Bài có gắn chương có nút **▶ Thực hành ngay**: mở luôn bài luyện các câu của chương đó trên Office thật
  (hết slide cuối / xem xong video cũng hiện nút này).
- Không có nút tải về: file gốc không chép vào app, ảnh slide và video được mã hóa, chỉ giải mã trong bộ nhớ khi xem.
  (Gói SCORM là trang web nên được giải nén trong thư mục bài; không chặn được chụp màn hình.)
- Có sẵn 2 bài mẫu Word chương 1: slide *Quản lý tài liệu* và trắc nghiệm SCORM *Trắc nghiệm nhanh* (6 câu).
- Bài đi kèm app nằm trong thư mục `tai_lieu\`; bài quản trị thêm nằm trong thư mục dữ liệu chung.
- Gợi ý làm video: PowerPoint → *File → Export → Create a Video* (có ghi âm lời giảng), hoặc quay màn hình bằng
  Clipchamp / OBS; mỗi video 2–5 phút cho một thao tác. Bài SCORM: soạn bằng iSpring Suite (add-in PowerPoint),
  Articulate, Captivate hoặc H5P rồi xuất **SCORM 1.2 / 2004 dạng .zip**.

## Lớp học trực tuyến (giáo viên quản lý học sinh ở mọi nơi)

Học sinh tải app về laptop, **đăng ký bằng mã lớp**; điểm và tiến độ học tự gửi lên máy chủ (Supabase, miễn phí),
giáo viên xem ở mục **🏫 Lớp học**: bảng điểm cả lớp, chi tiết từng học sinh, xuất Excel, đặt lại mật khẩu.
Mất mạng vẫn học được, điểm gửi sau. Vai trò: **Quản trị** (người đăng ký đầu tiên) · **Giáo viên** · **Học viên**.

👉 Cách tạo máy chủ, từng bước có hình dung: **[HUONG_DAN_MAY_CHU.md](HUONG_DAN_MAY_CHU.md)**.
Không cấu hình máy chủ thì app chạy ngoại tuyến như cũ (tài khoản trên từng máy).

## Tra từ khóa

Mục **🔎 Tra từ khóa**: gõ tên lệnh tiếng Anh hoặc tiếng Việt (không cần dấu, vd `hinh mo`, `freeze panes`) → xem
**thuật ngữ** (giải thích + đường dẫn menu), **slide bài giảng** có từ đó và **câu luyện trong đề** kèm cách làm;
bấm **▶ Luyện câu này** để làm riêng câu đó.

## Cách làm bài

1. Đăng nhập, chọn bài thi và chế độ:
   - **Luyện tập**: không giới hạn giờ, có nút *Gợi ý*, nút *Kiểm tra dự án* để chấm ngay.
   - **Thi thử**: 50 phút, không gợi ý, hết giờ tự nộp.
2. Phần mềm tự mở file bài làm bằng Office. Thanh yêu cầu nằm ở **cạnh dưới màn hình**, luôn nổi trên cùng.
3. Làm các nhiệm vụ, bấm **Ctrl+S để lưu**, chuyển sang *Dự án sau*, làm lần lượt đến hết rồi bấm **Nộp bài**.
4. Xem điểm và từng nhiệm vụ đúng/sai. Bấm vào một dòng để xem cách làm.

## Chương trình được viết như thế nào

```
main.py              ← điểm khởi động
mos/
  core.py            ← Đề thi / Dự án / Nhiệm vụ, chấm điểm, lịch sử, thư mục dữ liệu
  accounts.py        ← tài khoản: đăng nhập, vai trò, khóa, hạn dùng, mật khẩu băm
  rules.py           ← 46 luật chấm + thông tin để dựng form soạn đề
  custom.py          ← đọc / lưu đề tự soạn (de_thi/), gộp vào bài thi, kiểm tra đề
  ooxml.py           ← đọc XML bên trong file Office (.docx/.xlsx/.pptx là file ZIP)
  chuong.py          ← chương theo khung MOS, tạo bài luyện theo chương
  tai_lieu.py        ← bài giảng (slide PPTX / video / SCORM): nhập, mã hóa, tiến độ, điểm SCORM
  tu_khoa.py         ← thuật ngữ MOS và tìm kiếm (thuật ngữ, câu trong đề, slide)
  lop_hoc.py         ← Lớp học trực tuyến: kết nối Supabase, hàng đợi gửi điểm, báo cáo lớp, xuất Excel
  qt/                ← giao diện (Qt / PySide6)
    theme.py         ← hệ thống thiết kế: màu, font, stylesheet, thẻ, nút, bảng, vòng điểm
    app.py           ← đăng nhập, khung ứng dụng (thanh bên), trang chủ, lịch sử, thanh làm bài, kết quả
    admin.py         ← Quản trị: tài khoản, đề thi + form soạn đề, kết quả học viên
    hoc.py           ← Tài liệu học: danh sách, trình xem slide / video / SCORM, thêm bài giảng
    tra_cuu.py       ← trang Tra từ khóa
    lop.py           ← trang Lớp học, Tài khoản máy chủ, Đăng ký bằng mã lớp, Cài đặt máy chủ
  exams/             ← đề có sẵn (word.py, excel.py, powerpoint.py)
de_thi/              ← đề mẫu tự soạn (de.json + file gốc + dap_an/)
may_chu/             ← supabase_lop_hoc.sql: cài đặt máy chủ Lớp học (chạy 1 lần trong Supabase)
tai_lieu/            ← bài giảng đi kèm app (bai.json + slides.mosl / video.mosv / scorm/)
kiem_tra_de.py       ← kiểm tra đề bằng dòng lệnh
tests/               ← kiểm tra tự động (python -m pytest -q)
```

Luồng chấm điểm:

```
Tạo file mẫu ──► Mở bằng Office ──► Học viên làm & lưu ──► Đọc file & chấm ──► Điểm /1000 ──► Lưu kết quả theo tài khoản
                   os.startfile                            openpyxl / python-docx / python-pptx / đọc XML
```

Mỗi nhiệm vụ gồm **yêu cầu**, **gợi ý** và **hàm chấm**. Ví dụ luật `excel_co_dinh` trong `mos/rules.py`:

```python
@rule
def excel_co_dinh(path, o="A2", sheet=None):
    """Freeze Panes tại ô `o` (A2 = cố định hàng đầu, B1 = cột đầu)."""
    return _ws(path, sheet)[1].freeze_panes == o
```

Với những thứ thư viện không đọc được (biểu đồ, watermark, footnote, section…), chương trình mở file ZIP và
tìm trực tiếp trong XML. Ví dụ biểu đồ cột Excel nằm ở `xl/charts/chart1.xml`, trong thẻ `<c:barDir val="col"/>`.

**Thêm luật chấm mới (cho người biết Python):** viết một hàm trong `mos/rules.py` và gắn `@rule`. Sau đó thêm tên
tiếng Việt vào `RULE_TITLES`, và thêm nhãn tham số vào `PARAM_UI` nếu là tham số mới. Luật sẽ tự hiện trong form soạn đề.
