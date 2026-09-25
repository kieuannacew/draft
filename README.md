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

- Giao diện theo chủ đề màu **Forest Canopy** (xanh rừng, xanh ô liu, nền ngà; chữ hiện đại không chân – Segoe UI).
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
  qt/                ← giao diện (Qt / PySide6)
    theme.py         ← hệ thống thiết kế: màu, font, stylesheet, thẻ, nút, bảng, vòng điểm
    app.py           ← đăng nhập, khung ứng dụng (thanh bên), trang chủ, lịch sử, thanh làm bài, kết quả
    admin.py         ← Quản trị: tài khoản, đề thi + form soạn đề, kết quả học viên
  exams/             ← đề có sẵn (word.py, excel.py, powerpoint.py)
de_thi/              ← đề mẫu tự soạn (de.json + file gốc + dap_an/)
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
