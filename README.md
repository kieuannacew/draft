# Luyện thi MOS (kiểu GMetrix)

Phần mềm luyện thi **Microsoft Office Specialist** cho Word (MO-100), Excel (MO-200), PowerPoint (MO-300).
Giống GMetrix: bạn **làm bài trực tiếp trên Word/Excel/PowerPoint thật**, phần mềm tự đọc file để chấm điểm
theo thang 1000 (đạt từ 700).

## Cách chạy (Windows)

1. Cài [Python 3.10+](https://www.python.org/downloads/) (khi cài nhớ tick **Add Python to PATH**).
2. Tải repo này về, giải nén.
3. Nhấp đúp **`run.bat`** (lần đầu sẽ tự cài thư viện).

Muốn có file `.exe` để chép sang máy khác: nhấp đúp **`build_exe.bat`** → file nằm ở `dist\LuyenThiMOS.exe`
(chép kèm thư mục `de_thi` đặt cạnh file `.exe` nếu có đề tự soạn).

## Đưa đề của bạn vào app (không cần lập trình)

1. Soạn file gốc bằng Word/Excel/PowerPoint.
2. Viết `de.json`: liệt kê nhiệm vụ và chọn **luật chấm** có sẵn (46 luật), ví dụ:
   `{"luat": "excel_co_dinh", "o": "A2"}`.
3. Đặt cả hai vào `de_thi\<tên đề>\`, rồi kéo thả thư mục đó vào `kiem_tra_de.bat` để kiểm tra.
4. Mở app: các dự án của bạn được gộp luôn vào bài thi của môn đó (Word/Excel/PowerPoint).

Hướng dẫn chi tiết và bảng luật: **[HUONG_DAN_SOAN_DE.md](HUONG_DAN_SOAN_DE.md)**. Đề mẫu: `de_thi/Excel_Mau`.

## Cách làm bài

1. Chọn bài thi và chế độ:
   - **Luyện tập**: không giới hạn giờ, có nút *Gợi ý*, nút *Kiểm tra dự án* để chấm ngay.
   - **Thi thử**: 50 phút, không gợi ý, hết giờ tự nộp.
2. Phần mềm tự mở file bài làm bằng Office. Thanh yêu cầu nằm ở **cạnh dưới màn hình**, luôn nổi trên cùng.
3. Làm các nhiệm vụ → **Ctrl+S để lưu** → *Dự án sau* → … → **Nộp bài**.
4. Xem điểm, nhiệm vụ đúng/sai, bấm vào từng dòng để xem cách làm.

File bài làm và lịch sử được lưu tại `C:\Users\<tên>\MOS_Practice\`.

## Chương trình được viết như thế nào

```
main.py              ← điểm khởi động
mos/
  core.py            ← Đề thi / Dự án / Nhiệm vụ, chấm điểm, lưu lịch sử
  ooxml.py           ← đọc XML bên trong file Office (.docx/.xlsx/.pptx là file ZIP)
  gui.py             ← giao diện Tkinter: màn hình chính, thanh làm bài, kết quả
  rules.py           ← thư viện luật chấm dùng cho đề tự soạn (de.json)
  custom.py          ← nạp đề tự soạn từ thư mục de_thi/
  exams/
    word.py          ← đề Word: tạo file mẫu + hàm chấm từng nhiệm vụ
    excel.py         ← đề Excel
    powerpoint.py    ← đề PowerPoint
de_thi/              ← đề tự soạn (mỗi đề một thư mục: de.json + file gốc + dap_an/)
kiem_tra_de.py       ← công cụ kiểm tra đề tự soạn
tests/               ← kiểm tra: file gốc = 0 điểm, file làm đúng = 1000 điểm; kiểm tra từng luật
```

Luồng hoạt động:

```
Tạo file mẫu (build) ──► Mở bằng Office ──► Người học làm & lưu ──► Đọc file & chấm (check) ──► Điểm /1000
   python-docx                os.startfile                                openpyxl / python-docx
   openpyxl / python-pptx                                                 python-pptx / đọc XML
```

Mỗi nhiệm vụ gồm 3 phần: **yêu cầu**, **gợi ý**, **hàm chấm**. Ví dụ trong `mos/exams/excel.py`:

```python
def chk_freeze(path):
    # Excel lưu "Freeze Top Row" thành freeze_panes = "A2"
    return _sheet(_wb(path), "DoanhSo").freeze_panes == "A2"

Task("Cố định (freeze) hàng tiêu đề để luôn hiển thị khi cuộn.",
     "View > Freeze Panes > Freeze Top Row.",
     chk_freeze)
```

Với những thứ thư viện không đọc được (biểu đồ, watermark, footnote, section…), chương trình mở file ZIP và
tìm trực tiếp trong XML, ví dụ biểu đồ cột Excel nằm ở `xl/charts/chart1.xml` với thẻ `<c:barDir val="col"/>`.

## Thêm đề mới bằng code (cho người biết Python)

Cách dễ nhất là soạn đề bằng `de.json` như ở trên. Nếu cần kiểu chấm đặc biệt:

1. Viết hàm `build_xxx(path)` tạo file mẫu.
2. Viết hàm `chk_xxx(path) -> bool` cho mỗi nhiệm vụ.
3. Thêm `Project(...)` vào `EXAM` trong file đề tương ứng.
4. Thêm lời giải mẫu vào `tests/test_exams.py` và chạy `python -m pytest -q`.

> Mẹo tìm cách chấm: tự làm nhiệm vụ trong Office, lưu file, đổi đuôi thành `.zip`, giải nén và xem XML
> thay đổi ở đâu.
