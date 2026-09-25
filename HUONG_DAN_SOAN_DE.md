# Hướng dẫn soạn đề cho app Luyện thi MOS

> **Cách dễ nhất:** soạn đề ngay trong app. Đăng nhập tài khoản quản trị, vào mục
> **Đề thi** ở thanh menu bên trái, rồi bấm **Soạn đề mới**. Bạn chỉ cần chọn file, gõ yêu cầu và chọn luật chấm từ danh sách;
> app tự tạo `de.json`. Xem các bước trong [README](README.md#soạn-đề-ngay-trong-app-không-cần-lập-trình).
>
> Phần dưới đây dành cho ai muốn **viết `de.json` bằng tay**, và là **bảng tra các luật chấm**.
> Trong form soạn đề, mỗi luật hiện bằng tên tiếng Việt, ví dụ `excel_co_dinh` là "Cố định hàng / cột (Freeze Panes)".

Bạn **không cần biết lập trình**. Một đề gồm 2 thứ:

1. **File gốc** (`.xlsx`, `.docx`, `.pptx`): bạn tự soạn bằng Excel/Word/PowerPoint. Đây là file người học sẽ mở ra để làm bài.
2. **File `de.json`**: ghi các nhiệm vụ, gợi ý, và **luật chấm** cho từng nhiệm vụ.

## Chấm bằng công cụ gì?

App chấm bằng **Python**, đọc thẳng file người học đã lưu. Không cần mở Office và không cần macro/VBA.

| Loại file | Thư viện đọc file |
|---|---|
| Excel `.xlsx` | `openpyxl`: đọc công thức, định dạng, Table, Freeze Panes, Named Range… |
| Word `.docx` | `python-docx`: đọc đoạn văn, style, bảng, hướng trang, thuộc tính… |
| PowerPoint `.pptx` | `python-pptx`: đọc slide, tiêu đề, ghi chú, bảng, kích thước… |
| Phần thư viện không đọc được | đọc XML bên trong file (`.docx/.xlsx/.pptx` thực chất là file ZIP): biểu đồ, watermark, footnote, section, transition… |

Mỗi nhiệm vụ gắn với một hoặc nhiều **luật chấm** có sẵn (bảng ở cuối trang). Ví dụ luật
`excel_co_dinh` kiểm tra xem người học đã **Freeze Panes** chưa.

## Các bước soạn một đề

### Bước 1: Tạo thư mục đề
Tạo một thư mục con trong `de_thi` cạnh app, hoặc trong `C:\Users\<tên>\MOS_Practice\de_thi`, ví dụ `Excel_De2`.
Trong mục **Đề thi**, nút **Mở thư mục** sẽ mở nhanh chỗ lưu đề.

```
de_thi\
  Excel_De2\
    de.json          ← khai báo nhiệm vụ
    DoanhThu.xlsx    ← file gốc (bạn tự soạn)
    dap_an\          ← (nên có) file đã làm đúng, để kiểm tra đề
      DoanhThu.xlsx
```

### Bước 2: Soạn file gốc bằng Office
Nhập dữ liệu **ở trạng thái chưa làm**. Ví dụ muốn người học đổi tên sheet thì để nguyên "Sheet1".

### Bước 3: Viết `de.json`
Mở bằng Notepad (hoặc VS Code) và **lưu với mã hóa UTF-8**. Cách dễ nhất là copy
`de_thi\Excel_Mau\de.json` rồi sửa lại.

```json
{
  "mon": "EXCEL",
  "ten": "Excel – Đề 2",
  "thoi_gian": 50,
  "du_an": [
    {
      "ten": "Dự án 1 – Doanh thu",
      "file": "DoanhThu.xlsx",
      "mo_ta": "Bạn là kế toán, cần hoàn thiện bảng doanh thu.",
      "nhiem_vu": [
        {
          "yeu_cau": "Ô D2:D20 tính Thành tiền = Số lượng × Đơn giá.",
          "goi_y": "Chọn D2, gõ =B2*C2, kéo xuống D20.",
          "cham": {"luat": "excel_cong_thuc", "o": "D2:D20", "chua": ["B{hang}*C{hang}"]}
        },
        {
          "yeu_cau": "Cố định hàng tiêu đề và in đậm A1:D1.",
          "goi_y": "View > Freeze Panes > Freeze Top Row; chọn A1:D1 > Ctrl+B.",
          "cham": [
            {"luat": "excel_co_dinh", "o": "A2"},
            {"luat": "excel_font", "o": "A1:D1", "dam": true}
          ]
        }
      ]
    }
  ]
}
```

| Trường | Ý nghĩa |
|---|---|
| `mon` | `WORD`, `EXCEL` hoặc `POWERPOINT`: các dự án sẽ được gộp vào bài thi của môn này |
| `ten` | Tên đề (không bắt buộc, chỉ hiện khi chạy `kiem_tra_de`) |
| `thoi_gian` | Không bắt buộc. Trong app, thời gian thi là thời gian của bài thi môn đó (50 phút) |
| `du_an` | Danh sách dự án; mỗi dự án có một file gốc |
| `file` | Tên file gốc nằm cùng thư mục với `de.json` |
| `nhiem_vu` | Danh sách nhiệm vụ: `yeu_cau`, `goi_y`, `cham` |
| `cham` | Một luật `{...}`, hoặc danh sách `[{...}, {...}]` (phải đúng **tất cả** thì nhiệm vụ mới được tính đúng) |

Lưu ý khi viết JSON:
- Chữ phải nằm trong ngoặc kép `"..."`.
- `true`/`false` viết thường.
- Không có dấu phẩy sau phần tử cuối cùng.

### Bước 4: Kiểm tra đề
1. Mở file gốc bằng Office, **tự làm đúng hết**, rồi lưu vào `dap_an\` với **cùng tên file**.
2. Kéo thả thư mục đề vào **`kiem_tra_de.bat`**, hoặc chạy lệnh `py kiem_tra_de.py de_thi\Excel_De2`.

Công cụ sẽ báo:
- **Lỗi khai báo**: sai tên luật, thiếu tham số, không thấy file gốc.
- **File gốc: ĐÚNG ⚠**: nhiệm vụ đã đúng sẵn khi chưa làm gì, nghĩa là luật chấm quá dễ.
- **Đáp án: SAI ✗**: làm đúng mà vẫn bị chấm sai. Cần sửa luật, hoặc báo lại cho người viết app.

Khi thấy dòng `KẾT QUẢ: Đề ổn ✓`, mở lại app. Các dự án trong đề của bạn sẽ được **gộp thẳng vào bài thi
của môn tương ứng**. Ví dụ đề có `"mon": "EXCEL"` sẽ thành Dự án 3, 4… của bài thi Excel, không hiện thành thẻ riêng.

> Đề cũng có thể đặt ở `C:\Users\<tên>\MOS_Practice\de_thi\`. Khi dùng bản `.exe`, đặt thư mục `de_thi`
> cạnh file `LuyenThiMOS.exe`.

## Mẹo viết luật chấm
- **Đoạn văn Word** được tìm theo **phần đầu nội dung** (không phân biệt hoa/thường): `"doan": "Giới thiệu"`.
- **Slide PowerPoint** chọn theo tiêu đề (`"tieu_de": "Giá bán"`) hoặc số thứ tự (`"slide": 2`). Nên dùng tiêu đề,
  vì người học có thể thêm hoặc xóa slide làm thay đổi số thứ tự.
- **Công thức Excel**:
  - Viết `{hang}` để khớp với từng hàng. Ví dụ `"C{hang}*D{hang}"` ở ô E5 sẽ thành `C5*D5`.
  - Khi so sánh, app bỏ khoảng trắng và không phân biệt hoa/thường.
  - Chấp nhận nhiều cách viết bằng cách chỉ kiểm tra phần cốt lõi, ví dụ `["SUM(", "E2:E11"]`.
- **Không tìm được luật phù hợp?** Dùng luật nâng cao `xml_chua`:
  1. Làm thao tác đó trong Office rồi lưu file.
  2. Đổi đuôi file thành `.zip`, giải nén, và tìm xem XML thay đổi ở đâu.
  3. Viết luật, ví dụ `{"luat": "xml_chua", "part_regex": "word/document\\.xml", "chua": "<w:pgBorders"}` (viền trang).
- **Người biết Python** có thể thêm luật mới vào `mos/rules.py`: viết hàm và gắn `@rule`, tên hàm chính là tên luật.

## Danh sách luật chấm

Xem nhanh bằng lệnh `py kiem_tra_de.py --luat`.

### Excel

| Luật | Tham số | Ý nghĩa |
|---|---|---|
| `excel_ten_sheet` | `ten` | Có trang tính tên `ten`. |
| `excel_khong_co_sheet` | `ten` | KHÔNG còn trang tính tên `ten` (vd đã đổi tên / xóa). |
| `excel_cong_thuc` | `o`, `chua`, `sheet` *(tuỳ chọn)* | Mọi ô trong vùng `o` là công thức chứa tất cả chuỗi trong `chua`. Dùng {hang} để thay bằng số hàng của từng ô, vd "C{hang}*D{hang}". |
| `excel_gia_tri` | `o`, `bang`, `sheet` *(tuỳ chọn)* | Ô `o` có giá trị bằng `bang` (so sánh không phân biệt hoa/thường). |
| `excel_dinh_dang_so` | `o`, `chua` *(tuỳ chọn)*, `sheet` *(tuỳ chọn)* | Vùng `o` có định dạng số khác General; nếu có `chua` thì mã định dạng phải chứa chuỗi đó (vd "%", "0.00"). |
| `excel_font` | `o`, `dam` *(tuỳ chọn)*, `nghieng` *(tuỳ chọn)*, `co` *(tuỳ chọn)*, `sheet` *(tuỳ chọn)* | Font của vùng `o`: dam (in đậm), nghieng, co (cỡ chữ). |
| `excel_co_bang` | `vung`, `sheet` *(tuỳ chọn)*, `ten` *(tuỳ chọn)* | Có Table (Format as Table) bắt đầu ở góc trên-trái và phủ hết `vung`. |
| `excel_co_dinh` | `o` *(mặc định 'A2')*, `sheet` *(tuỳ chọn)* | Freeze Panes tại ô `o` (A2 = cố định hàng đầu, B1 = cột đầu). |
| `excel_dinh_dang_dieu_kien` | `vung`, `sheet` *(tuỳ chọn)* | Có Conditional Formatting phủ vùng `vung`. |
| `excel_ten_vung` | `ten`, `vung` | Có Named Range `ten` trỏ tới `vung` (vd "C2:C11"). |
| `excel_bieu_do` | `loai` *(mặc định 'bat_ky')* | Có biểu đồ. loai: cot, thanh (ngang), duong, tron, bat_ky. |
| `excel_huong_trang` | `huong` *(mặc định 'ngang')*, `sheet` *(tuỳ chọn)* | Hướng trang in: ngang hoặc doc. |
| `excel_sap_xep` | `cot`, `tu`, `den`, `chieu` *(mặc định 'tang')*, `sheet` *(tuỳ chọn)* | Cột `cot` từ hàng `tu` đến `den` được sắp xếp tang / giam. |

### Word

| Luật | Tham số | Ý nghĩa |
|---|---|---|
| `word_kieu_doan` | `doan`, `kieu` | Các đoạn (bắt đầu bằng `doan`, có thể là danh sách) dùng style `kieu` (vd "Heading 1", "Title"). |
| `word_can_le` | `doan`, `can` *(mặc định 'giua')* | Căn lề đoạn: trai, giua, phai, deu. |
| `word_co_chu` | `chu`, `so_lan` *(tuỳ chọn)* | Tài liệu có chứa `chu` (nếu có so_lan thì phải xuất hiện đúng số lần). |
| `word_khong_co_chu` | `chu` | Tài liệu KHÔNG còn chứa `chu` (vd sau Replace All). |
| `word_dau_dong` | `doan` | Các đoạn trong `doan` là danh sách bullet/đánh số. |
| `word_gian_dong` | `doan`, `gia_tri` | Giãn dòng của đoạn = gia_tri (vd 1.5, 2). |
| `word_bang` | `so_cot`, `so_hang` *(tuỳ chọn)*, `o_dau` *(tuỳ chọn)* | Có bảng `so_cot` cột (và ít nhất `so_hang` hàng; ô đầu = `o_dau` nếu có). |
| `word_muc_luc` | — | Có mục lục tự động (Table of Contents). |
| `word_so_trang` | `vi_tri` *(mặc định 'bat_ky')* | Có số trang (trường PAGE) ở chan_trang, dau_trang hoặc bat_ky. |
| `word_dau_trang_chan_trang` | `chu`, `vi_tri` *(mặc định 'bat_ky')* | Đầu/chân trang có chứa `chu`. |
| `word_watermark` | `chu` | Có watermark chữ `chu`. |
| `word_huong_trang` | `huong` *(mặc định 'ngang')* | Hướng trang của section đầu: ngang hoặc doc. |
| `word_so_cot` | `so` | Có đoạn văn bản được chia `so` cột (Layout > Columns). |
| `word_footnote` | `chu` | Có footnote chứa `chu`. |
| `word_thuoc_tinh` | `truong`, `gia_tri` | Thuộc tính tài liệu: truong = title, author, subject, keywords, comments, category. |
| `word_theo_doi_thay_doi` | — | Đã bật Track Changes. |
| `word_hinh_anh` | `alt_text` *(tuỳ chọn)* | Có hình ảnh (nếu có alt_text thì mô tả thay thế phải chứa chuỗi này). |

### PowerPoint

| Luật | Tham số | Ý nghĩa |
|---|---|---|
| `ppt_co_slide` | `tieu_de`, `bo_cuc` *(tuỳ chọn)*, `vi_tri` *(tuỳ chọn)* | Có slide tiêu đề `tieu_de` (bố cục `bo_cuc`, ở vị trí `vi_tri`: số thứ tự hoặc "cuoi"). |
| `ppt_khong_co_slide` | `tieu_de` | Đã xóa slide tiêu đề `tieu_de`. |
| `ppt_so_slide` | `so` | Bài có đúng `so` slide. |
| `ppt_ghi_chu` | `chu`, `tieu_de` *(tuỳ chọn)*, `slide` *(tuỳ chọn)* | Slide có Notes chứa `chu`. |
| `ppt_chuyen_trang` | `tieu_de` *(tuỳ chọn)*, `slide` *(tuỳ chọn)* | Có Transition (không ghi tieu_de/slide = tất cả slide). |
| `ppt_hieu_ung` | `tieu_de` *(tuỳ chọn)*, `slide` *(tuỳ chọn)*, `nhom` *(mặc định 'entr')* | Slide có Animation. nhom: entr (xuất hiện), exit (biến mất), emph (nhấn mạnh), bat_ky. |
| `ppt_an_slide` | `tieu_de` *(tuỳ chọn)*, `slide` *(tuỳ chọn)* | Slide bị ẩn (Hide Slide). |
| `ppt_kich_thuoc` | `ti_le` *(mặc định '16:9')* | Kích thước slide: "16:9" hoặc "4:3". |
| `ppt_bang` | `so_cot`, `so_hang`, `tieu_de` *(tuỳ chọn)*, `slide` *(tuỳ chọn)* | Slide có bảng `so_cot` cột × `so_hang` hàng. |
| `ppt_smartart` | `tieu_de` *(tuỳ chọn)*, `slide` *(tuỳ chọn)* | Slide có SmartArt. |
| `ppt_bieu_do` | `tieu_de` *(tuỳ chọn)*, `slide` *(tuỳ chọn)* | Slide có biểu đồ (Chart). |
| `ppt_hinh_anh` | `tieu_de` *(tuỳ chọn)*, `slide` *(tuỳ chọn)* | Slide có hình ảnh. |
| `ppt_section` | `ten` | Có Section tên `ten`. |
| `ppt_so_trang` | `tru_slide_dau` *(mặc định True)* | Hiện Slide number trên mọi slide (mặc định bỏ qua slide tiêu đề). |
| `ppt_co_chu` | `chu`, `tieu_de` *(tuỳ chọn)*, `slide` *(tuỳ chọn)* | Slide (hoặc cả bài) có chứa `chu`. |

### Bộ đề nhập

- `word_mau_de(dang, de_bai, goc)` – chấm tự động theo mã dạng câu của bộ đề nhập (vd `pic_size`), tự sinh khi
  bấm **Nhập bộ đề**. Không cần viết tay.
- `tu_kiem_tra(ghi_chu)` – không chấm tự động: học viên tự đối chiếu, nhiệm vụ không tính điểm.

### Nâng cao

| Luật | Tham số | Ý nghĩa |
|---|---|---|
| `xml_chua` | `part_regex`, `chua`, `bo_qua_hoa_thuong` *(mặc định False)* | Luật nâng cao: có part XML (khớp regex `part_regex`) chứa regex `chua`. |
