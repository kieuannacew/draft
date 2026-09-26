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

## Tài khoản và phân quyền

| Vai trò | Được làm gì |
|---|---|
| **Quản trị** | Toàn quyền: quản lý mọi tài khoản (kể cả giáo viên), soạn đề, xem kết quả của tất cả |
| **Giáo viên** | Tạo / nhập / khóa tài khoản **học viên của mình**, soạn đề, xem kết quả **học viên của mình** |
| **Học viên** | Làm bài, xem lịch sử của chính mình, đổi mật khẩu |

**Lần đầu chạy:** đăng nhập `admin` / `admin`. App sẽ bắt đổi mật khẩu ngay lần đăng nhập này.
Sau đó vào mục **Tài khoản**, bấm **+ Thêm tài khoản** và chọn vai trò **Giáo viên** để tạo tài khoản cho từng giáo viên.

Cấp tài khoản học viên (mục **Tài khoản** / **Học viên** ở thanh menu bên trái):
- **Nhập từ file (CSV / Excel):** chọn file danh sách. App cho **xem trước** từng dòng (sẽ tạo hay bị bỏ qua và
  vì sao), rồi mới tạo tài khoản. Chỉ cột **Họ tên** là bắt buộc; các cột tuỳ chọn là **Tên đăng nhập**,
  **Mật khẩu**, **Lớp**, **Hạn dùng**. Thiếu tên đăng nhập thì app tự tạo từ họ tên (Nguyễn Văn An → `nguyenvanan`);
  thiếu mật khẩu thì app tạo ngẫu nhiên. Nút **Tải file mẫu** cho sẵn một file Excel mẫu để điền.
- **Tạo cho cả lớp:** dán danh sách họ tên, mỗi dòng một người.
- **+ Thêm tài khoản:** tạo từng người.

Tài khoản giáo viên tạo ra tự động thuộc về giáo viên đó. Quản trị có thể chọn **giáo viên phụ trách** cho từng học viên.
Sau khi tạo, app hiện danh sách **tên đăng nhập / mật khẩu** để lưu ra file CSV phát cho học viên.

Mục **Kết quả học viên** cho xem điểm theo lớp hoặc theo từng học viên, và **xuất CSV** để mở bằng Excel.

> Mật khẩu được lưu dạng mã băm (PBKDF2), không lưu mật khẩu thật. Dữ liệu nằm trong
> `C:\Users\<tên>\MOS_Practice\` (`tai_khoan.json`, `history.json`, `de_thi\`).
>
> **Phòng máy dùng chung dữ liệu:** tạo file `cau_hinh.json` cạnh `main.py` (hoặc cạnh `LuyenThiMOS.exe`):
> `{"thu_muc_du_lieu": "\\\\MAYCHU\\MOS"}`. Khi đó mọi máy dùng chung tài khoản, đề và kết quả trong thư mục mạng đó.

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

## Nếu Word / Excel / PowerPoint mở file bị mờ

Khi app được giải nén từ file ZIP tải trên mạng, Office có thể mở file ở **chế độ bảo vệ (Protected View)**:
thanh công cụ mờ đi và không sửa được. App đã tự gỡ dấu này cho file bài làm. Nếu vẫn gặp:
- Bấm **Enable Editing / Bật chỉnh sửa** trên dải vàng ở đầu cửa sổ, hoặc
- Trước khi giải nén: chuột phải file ZIP, chọn **Properties**, tick **Unblock**, bấm OK, rồi giải nén lại.

## Cách làm bài

1. Đăng nhập, chọn bài thi và chế độ:
   - **Luyện tập**: không giới hạn giờ, có nút *Gợi ý*, nút *Kiểm tra dự án* để chấm ngay.
   - **Thi thử**: 50 phút, không gợi ý, hết giờ tự nộp.
2. Phần mềm tự mở file bài làm bằng Office. Thanh đề **gắn ở cạnh dưới màn hình** (giống Taskbar), và cửa sổ
   Word/Excel/PowerPoint được **tự thu nhỏ vừa phần trống phía trên**, không bị thanh đề che.
3. Làm các nhiệm vụ, bấm **Ctrl+S để lưu**, chuyển sang *Dự án sau*, làm lần lượt đến hết rồi bấm **Nộp bài**.
4. Xem điểm và từng nhiệm vụ đúng/sai. Bấm vào một dòng để xem cách làm.

## Chương trình được viết như thế nào

```
main.py              ← điểm khởi động
mos/
  core.py            ← Đề thi / Dự án / Nhiệm vụ, chấm điểm, lịch sử, thư mục dữ liệu
  accounts.py        ← tài khoản: đăng nhập, vai trò quản trị / giáo viên / học viên, lớp, khóa, hạn dùng
  importer.py        ← nhập danh sách học viên từ CSV / Excel
  rules.py           ← 46 luật chấm + thông tin để dựng form soạn đề
  custom.py          ← đọc / lưu đề tự soạn (de_thi/), gộp vào bài thi, kiểm tra đề
  ooxml.py           ← đọc XML bên trong file Office (.docx/.xlsx/.pptx là file ZIP)
  qt/                ← giao diện (Qt / PySide6)
    theme.py         ← hệ thống thiết kế: màu, font, stylesheet, thẻ, nút, bảng, vòng điểm
    app.py           ← đăng nhập, khung ứng dụng (thanh bên), trang chủ, lịch sử, thanh làm bài, kết quả
    admin.py         ← Quản trị: tài khoản, nhập học viên từ file, đề thi + form soạn đề, kết quả
    winlayout.py     ← Windows: gắn thanh đề ở đáy màn hình, thu cửa sổ Office vừa phần trống
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
