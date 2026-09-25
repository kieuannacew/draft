"""Tra từ khóa: từ điển thuật ngữ MOS (Việt / Anh, kèm đường dẫn lệnh) + tìm trong đề và bài giảng.

Tìm không phân biệt hoa/thường và dấu tiếng Việt ("muc luc" khớp "mục lục").
"""
from __future__ import annotations

import re
import unicodedata
from dataclasses import dataclass

# (thuật ngữ, môn, đường dẫn lệnh, giải thích Việt, giải thích Anh, từ khóa phụ)
_GLOSSARY = [
    # ---------------- Word
    ("Watermark", "WORD", "Design > Watermark", "Chữ hoặc hình mờ nằm sau nội dung mỗi trang (vd BẢO MẬT, DRAFT).",
     "Faded text or picture behind the content of every page (e.g. CONFIDENTIAL).", "hình mờ chữ mờ"),
    ("Header / Footer", "WORD", "Insert > Header / Footer",
     "Vùng đầu / chân mỗi trang, thường chứa tiêu đề, số trang. Different First Page để trang 1 khác.",
     "Area at the top / bottom of each page. Different First Page makes page 1 different.", "đầu trang chân trang"),
    ("Page Number", "WORD", "Insert > Page Number", "Chèn số trang; Format Page Numbers để đổi kiểu số / bắt đầu từ số khác.",
     "Insert page numbers; Format Page Numbers changes the style or starting number.", "số trang"),
    ("Margins", "WORD", "Layout > Margins", "Lề trang. Custom Margins để nhập số đo trên, dưới, trái, phải.",
     "Page margins. Custom Margins lets you type top, bottom, left and right.", "lề trang can le"),
    ("Orientation", "WORD", "Layout > Orientation", "Hướng trang: Portrait (dọc) hoặc Landscape (ngang).",
     "Page orientation: Portrait or Landscape.", "hướng trang ngang dọc landscape portrait"),
    ("Paper Size", "WORD", "Layout > Size", "Khổ giấy (A4, A5, Letter, Legal…).", "Paper size (A4, A5, Letter, Legal…).",
     "khổ giấy size"),
    ("Columns", "WORD", "Layout > Columns", "Chia chữ thành nhiều cột; More Columns để chỉnh khoảng cách, đường kẻ giữa cột.",
     "Split text into columns; More Columns sets spacing and a line between.", "chia cột"),
    ("Section Break", "WORD", "Layout > Breaks > Section Breaks",
     "Ngắt section: Next Page (sang trang mới), Continuous (cùng trang). Mỗi section có thể có lề, hướng trang riêng.",
     "Section break: Next Page or Continuous. Each section can have its own margins and orientation.", "ngắt section"),
    ("Page Break", "WORD", "Insert > Page Break (Ctrl + Enter)", "Ngắt trang: đẩy phần sau sang trang mới.",
     "Starts the following content on a new page.", "ngắt trang"),
    ("Themes", "ALL", "Design > Themes", "Bộ màu, font, hiệu ứng áp dụng cho cả tài liệu.",
     "A set of colors, fonts and effects for the whole file.", "chủ đề theme colors fonts"),
    ("Style Set", "WORD", "Design > Document Formatting", "Bộ kiểu định dạng cho tiêu đề và nội dung của cả tài liệu.",
     "A set of formats for headings and body text in the whole document.", "bộ kiểu"),
    ("Styles", "WORD", "Home > Styles", "Kiểu định dạng có tên (Heading 1, Title, Quote…). Áp dụng để định dạng đồng bộ và tạo mục lục.",
     "Named formats (Heading 1, Title, Quote…) for consistent formatting and tables of contents.", "kiểu heading title"),
    ("Format Painter", "ALL", "Home > Format Painter", "Sao chép định dạng của đoạn / chữ này sang chỗ khác.",
     "Copies formatting from one place to another.", "sao chép định dạng chổi"),
    ("Clear All Formatting", "WORD", "Home > Clear All Formatting", "Xóa mọi định dạng, đưa chữ về kiểu Normal.",
     "Removes all formatting and returns text to Normal.", "xóa định dạng"),
    ("Line Spacing", "ALL", "Home > Line and Paragraph Spacing", "Giãn dòng: 1.0, 1.5, 2.0, Exactly (chính xác theo pt)…",
     "Space between lines: 1.0, 1.5, 2.0, Exactly…", "giãn dòng khoảng cách dòng"),
    ("Paragraph Spacing", "WORD", "Layout > Spacing Before / After", "Khoảng cách trước / sau đoạn văn.",
     "Space before / after a paragraph.", "khoảng cách đoạn"),
    ("Indent", "WORD", "Layout > Indent", "Thụt lề trái / phải; First line và Hanging trong hộp thoại Paragraph.",
     "Left / right indent; First line and Hanging in the Paragraph dialog.", "thụt lề"),
    ("Drop Cap", "WORD", "Insert > Drop Cap", "Chữ cái đầu đoạn phóng to, rơi xuống vài dòng.",
     "Large first letter that drops several lines.", "chữ cái lớn đầu đoạn"),
    ("Text Effects", "WORD", "Home > Text Effects and Typography", "Hiệu ứng chữ: bóng, viền, phát sáng, phản chiếu.",
     "Text effects: shadow, outline, glow, reflection.", "hiệu ứng chữ"),
    ("Change Case", "WORD", "Home > Change Case (Aa)", "Đổi chữ hoa / thường: UPPERCASE, lowercase, Capitalize Each Word.",
     "Change case: UPPERCASE, lowercase, Capitalize Each Word.", "chữ hoa chữ thường"),
    ("Find / Replace", "ALL", "Home > Find (Ctrl + F) / Replace (Ctrl + H)", "Tìm và thay thế chữ; More >> để tìm theo định dạng, dùng wildcards.",
     "Find and replace text; More >> for formatting and wildcards.", "tìm kiếm thay thế"),
    ("Symbol", "ALL", "Insert > Symbol > More Symbols", "Chèn ký hiệu đặc biệt (©, ®, ký hiệu Wingdings…) theo font và mã ký tự.",
     "Insert special characters by font and character code.", "ký hiệu ký tự đặc biệt"),
    ("Hyperlink", "ALL", "Insert > Link (Ctrl + K)", "Liên kết tới trang web, file hoặc vị trí trong tài liệu (Place in This Document).",
     "Link to a web page, file or place in the document.", "siêu liên kết link"),
    ("Bookmark", "WORD", "Insert > Bookmark", "Đánh dấu một vị trí có tên để liên kết / tham chiếu tới.",
     "A named location you can link or refer to.", "đánh dấu"),
    ("Table of Contents", "WORD", "References > Table of Contents", "Mục lục tự động dựa trên các style Heading.",
     "Automatic contents list built from Heading styles.", "mục lục toc"),
    ("Footnote / Endnote", "WORD", "References > Insert Footnote / Insert Endnote",
     "Chú thích cuối trang / cuối tài liệu; hộp thoại để đổi định dạng số, chuyển đổi qua lại.",
     "Notes at the page bottom / document end; the dialog changes numbering and converts.", "chú thích cuối trang"),
    ("Citation / Bibliography", "WORD", "References > Insert Citation / Bibliography",
     "Trích dẫn nguồn và danh mục tài liệu tham khảo; Style đổi kiểu APA, MLA…",
     "Cite sources and build a bibliography; Style switches APA, MLA…", "trích dẫn tài liệu tham khảo"),
    ("Caption", "ALL", "References > Insert Caption", "Chú thích đánh số cho ảnh / bảng (Figure 1, Table 1).",
     "Numbered label for pictures / tables (Figure 1, Table 1).", "chú thích ảnh bảng"),
    ("Cross-reference", "WORD", "References > Cross-reference", "Tham chiếu tới bảng, hình, tiêu đề (tự cập nhật số).",
     "Refers to a table, figure or heading (numbers update).", "tham chiếu chéo"),
    ("Index", "WORD", "References > Mark Entry / Insert Index", "Chỉ mục cuối tài liệu: đánh dấu mục rồi chèn Index.",
     "Mark entries then insert an index at the end.", "chỉ mục"),
    ("Convert Text to Table", "WORD", "Insert > Table > Convert Text to Table", "Biến dòng chữ phân cách bằng Tab / dấu phẩy thành bảng.",
     "Turn tab- or comma-separated text into a table.", "chuyển chữ thành bảng"),
    ("Merge / Split Cells", "ALL", "Table Layout > Merge Cells / Split Cells", "Gộp nhiều ô thành một / tách một ô thành nhiều ô.",
     "Combine cells / divide a cell.", "gộp ô tách ô"),
    ("AutoFit", "WORD", "Table Layout > AutoFit", "Co giãn bảng theo nội dung (Contents) hoặc theo chiều rộng trang (Window).",
     "Resize a table to its contents or the page width.", "co giãn bảng"),
    ("Repeat Header Rows", "WORD", "Table Layout > Repeat Header Rows", "Lặp lại hàng tiêu đề bảng ở đầu mỗi trang.",
     "Repeats the table header row on each page.", "lặp tiêu đề bảng"),
    ("Sort", "ALL", "Table Layout > Sort / Data > Sort", "Sắp xếp dữ liệu tăng dần / giảm dần theo cột.",
     "Order data ascending / descending by column.", "sắp xếp"),
    ("Bullets / Numbering", "ALL", "Home > Bullets / Numbering", "Danh sách gạch đầu dòng / đánh số; Define New để tự chọn ký hiệu, kiểu số.",
     "Bulleted / numbered lists; Define New for custom symbols or formats.", "danh sách đầu dòng đánh số"),
    ("Alt Text", "ALL", "Right-click > View / Edit Alt Text", "Mô tả thay thế cho ảnh, biểu đồ, SmartArt – giúp người khiếm thị.",
     "Alternative description for pictures, charts, SmartArt.", "mô tả thay thế"),
    ("Text Wrapping", "ALL", "Picture Format > Wrap Text", "Cách chữ bao quanh ảnh: Square, Tight, Through, Top and Bottom, In Line with Text…",
     "How text flows around a picture.", "ngắt dòng quanh ảnh"),
    ("Picture Styles / Effects", "ALL", "Picture Format > Picture Styles / Picture Effects",
     "Khung, bóng, phản chiếu, viền mềm cho ảnh.", "Frames, shadows, reflections, soft edges for pictures.", "kiểu ảnh hiệu ứng ảnh"),
    ("Remove Background", "ALL", "Picture Format > Remove Background", "Xóa nền ảnh, giữ lại chủ thể.",
     "Removes the background of a picture.", "xóa nền"),
    ("SmartArt", "ALL", "Insert > SmartArt", "Sơ đồ dạng danh sách, quy trình, chu trình…; SmartArt Design để đổi bố cục, màu, kiểu.",
     "Diagrams for lists, processes, cycles…; SmartArt Design changes layout, colors, style.", "sơ đồ"),
    ("Text Box", "ALL", "Insert > Text Box", "Khung chữ đặt tự do trên trang; Create Link để nối chữ tràn giữa hai khung.",
     "A free-floating box of text; Create Link flows text between boxes.", "khung chữ hộp văn bản"),
    ("3D Models", "ALL", "Insert > 3D Models", "Chèn mô hình 3D có thể xoay.", "Insert a rotatable 3D model.", "mô hình 3d"),
    ("Comments", "ALL", "Review > New Comment", "Ghi chú góp ý; Reply để trả lời, Resolve để đánh dấu đã xử lý, Delete để xóa.",
     "Review notes; Reply, Resolve or Delete them.", "bình luận nhận xét"),
    ("Track Changes", "WORD", "Review > Track Changes", "Ghi lại mọi sửa đổi; Accept / Reject để chấp nhận / từ chối; Lock Tracking để khóa.",
     "Records every edit; Accept / Reject them; Lock Tracking locks it.", "theo dõi thay đổi"),
    ("Restrict Editing", "WORD", "Review > Restrict Editing", "Giới hạn định dạng / chỉnh sửa (chỉ comment, chỉ điền form…).",
     "Limit formatting or editing (comments only, forms…).", "hạn chế chỉnh sửa bảo vệ"),
    ("Inspect Document", "ALL", "File > Info > Check for Issues > Inspect Document",
     "Tìm và xóa thông tin ẩn: thuộc tính, header/footer, comment…", "Finds and removes hidden data.", "kiểm tra tài liệu"),
    ("Check Accessibility", "ALL", "Review > Check Accessibility", "Tìm lỗi trợ năng (thiếu alt text, thiếu hàng tiêu đề bảng…).",
     "Finds accessibility issues (missing alt text, table headers…).", "trợ năng"),
    ("Compatibility Mode", "ALL", "File > Info > Convert", "Chế độ tương thích bản cũ; Convert để nâng cấp.",
     "Older-version mode; Convert upgrades the file.", "chế độ tương thích"),
    ("Document Properties", "ALL", "File > Info > Properties", "Title, Subject, Tags (Keywords), Categories, Company…",
     "Title, Subject, Tags (Keywords), Categories, Company…", "thuộc tính tệp tags keywords"),
    ("Save As / Save a Copy", "ALL", "File > Save As / Save a Copy", "Lưu dạng khác: PDF, RTF, TXT, Word 97-2003, Template (.dotx)…",
     "Save in another format: PDF, RTF, TXT, 97-2003, Template…", "lưu bản sao pdf"),
    ("Mail Merge", "WORD", "Mailings > Start Mail Merge", "Trộn thư: tạo nhiều thư từ danh sách người nhận (Excel).",
     "Create many letters from a recipient list.", "trộn thư"),
    ("Quick Parts", "WORD", "Insert > Quick Parts", "Lưu và chèn lại khối nội dung, trường (field) như NumWords, SaveDate.",
     "Store reusable content and insert fields.", "field trường"),
    ("Macro", "ALL", "View > Macros > Record Macro", "Ghi lại chuỗi thao tác để chạy lại; lưu file dạng .docm / .xlsm.",
     "Record actions to replay; save as .docm / .xlsm.", "macro"),
    # ---------------- Excel
    ("Freeze Panes", "EXCEL", "View > Freeze Panes", "Cố định hàng / cột để luôn hiển thị khi cuộn.",
     "Keep rows / columns visible while scrolling.", "cố định hàng cột"),
    ("Format as Table", "EXCEL", "Home > Format as Table (Ctrl + T)", "Biến vùng dữ liệu thành bảng có tiêu đề, lọc, kiểu màu.",
     "Turn a range into a table with headers, filters, styles.", "định dạng bảng"),
    ("Named Range", "EXCEL", "Formulas > Define Name / Name Box", "Đặt tên cho ô / vùng để dùng trong công thức.",
     "Give a cell or range a name for formulas.", "đặt tên vùng"),
    ("Conditional Formatting", "EXCEL", "Home > Conditional Formatting", "Tô màu ô theo điều kiện (Data Bars, Color Scales, Highlight Rules).",
     "Format cells by rules (Data Bars, Color Scales…).", "định dạng có điều kiện"),
    ("Number Format", "EXCEL", "Home > Number", "Định dạng số: Currency, Accounting, Percentage, số chữ số thập phân…",
     "Number formats: Currency, Accounting, Percentage, decimals…", "định dạng số"),
    ("SUM / AVERAGE / COUNT", "EXCEL", "Formulas > AutoSum", "Hàm tổng, trung bình, đếm. Ví dụ =SUM(E2:E11).",
     "Sum, average and count functions. e.g. =SUM(E2:E11).", "hàm tổng trung bình đếm"),
    ("IF", "EXCEL", "Formulas > Logical > IF", "Hàm điều kiện: =IF(điều kiện, giá trị đúng, giá trị sai).",
     "Conditional function: =IF(test, value if true, value if false).", "hàm if điều kiện"),
    ("COUNTIF / SUMIF", "EXCEL", "Formulas > More Functions", "Đếm / cộng các ô thỏa điều kiện.", "Count / sum cells that meet a condition.",
     "đếm có điều kiện"),
    ("Chart", "ALL", "Insert > Chart", "Biểu đồ; Chart Design để đổi kiểu, màu, bố cục, Switch Row/Column.",
     "Charts; Chart Design changes type, colors, layout, Switch Row/Column.", "biểu đồ"),
    ("Print Area / Page Setup", "EXCEL", "Page Layout > Print Area / Orientation", "Vùng in, hướng giấy, thu nhỏ khi in.",
     "Print area, orientation, scaling.", "vùng in"),
    # ---------------- PowerPoint
    ("Slide Master", "POWERPOINT", "View > Slide Master", "Mẫu chung của mọi slide (font, màu, logo).",
     "The template behind every slide.", "mẫu slide"),
    ("Transitions", "POWERPOINT", "Transitions", "Hiệu ứng chuyển slide; Apply To All để áp cho tất cả.",
     "Effects between slides; Apply To All.", "chuyển trang"),
    ("Animations", "POWERPOINT", "Animations", "Hiệu ứng cho đối tượng: Entrance, Emphasis, Exit, Motion Paths.",
     "Object effects: Entrance, Emphasis, Exit, Motion Paths.", "hiệu ứng động"),
    ("Hide Slide", "POWERPOINT", "Slide Show > Hide Slide", "Ẩn slide khi trình chiếu.", "Hides a slide during the show.", "ẩn slide"),
    ("Sections", "POWERPOINT", "Home > Section > Add Section", "Nhóm các slide thành phần có tên.", "Group slides into named sections.",
     "phần section"),
    ("Slide Size", "POWERPOINT", "Design > Slide Size", "Tỉ lệ slide 16:9 (Widescreen) hoặc 4:3 (Standard).",
     "Slide ratio 16:9 or 4:3.", "kích thước slide"),
    ("Speaker Notes", "POWERPOINT", "View > Notes", "Ghi chú cho người thuyết trình dưới mỗi slide.",
     "Notes for the presenter under each slide.", "ghi chú"),
]


@dataclass
class Term:
    term: str
    subject: str
    path: str
    vi: str
    en: str
    extra: str


GLOSSARY = [Term(*row) for row in _GLOSSARY]


def fold(text: str) -> str:
    """Chữ thường, bỏ dấu tiếng Việt, gộp khoảng trắng."""
    text = unicodedata.normalize("NFKD", str(text or "").replace("đ", "d").replace("Đ", "D"))
    text = "".join(c for c in text if not unicodedata.combining(c)).casefold()
    return " ".join(re.sub(r"[^\w#.,+<>/-]+", " ", text).split())


def _match(words: list[str], *fields: str) -> bool:
    hay = fold(" ".join(f for f in fields if f))
    return all(w in hay for w in words)


def search_terms(query: str, subject: str | None = None) -> list[Term]:
    words = fold(query).split()
    if not words:
        return []
    out = [t for t in GLOSSARY if (subject is None or t.subject in (subject, "ALL"))
           and _match(words, t.term, t.path, t.vi, t.en, t.extra)]
    q = fold(query)
    return sorted(out, key=lambda t: (not fold(t.term).startswith(q), fold(t.term)))


@dataclass
class TaskHit:
    exam: object
    project: object
    task: object


def _phrase_first(query: str, run):
    """Tìm nguyên cụm trước; không có kết quả mới tìm theo từng từ."""
    q = fold(query)
    if not q:
        return []
    if " " in q:
        hits = run([q])
        if hits:
            return hits
    return run(q.split())


def search_tasks(query: str, exams: list, limit: int = 60) -> list[TaskHit]:
    return _phrase_first(query, lambda words: _search_tasks(words, exams, limit))


def _search_tasks(words: list[str], exams: list, limit: int) -> list[TaskHit]:
    out, seen = [], set()
    for exam in exams:
        for p in exam.projects:
            for t in p.tasks:
                if t.title in seen:
                    continue
                if _match(words, t.title, t.title_en, t.hint, t.hint_en):
                    seen.add(t.title)
                    out.append(TaskHit(exam, p, t))
                    if len(out) >= limit:
                        return out
    return out


@dataclass
class SlideHit:
    lesson: object
    number: int
    snippet: str


def search_slides(query: str, lessons: list, limit: int = 40) -> list[SlideHit]:
    return _phrase_first(query, lambda words: _search_slides(words, lessons, limit))


def _search_slides(words: list[str], lessons: list, limit: int) -> list[SlideHit]:
    out = []
    for lesson in lessons:
        for i, s in enumerate(lesson.slides, start=1):
            text = (s.get("text") or "") + "\n" + (s.get("notes") or "")
            if _match(words, text):
                line = next((x for x in text.splitlines() if _match(words[:1], x)), text.splitlines()[0] if text else "")
                out.append(SlideHit(lesson, i, line.strip()[:160]))
                if len(out) >= limit:
                    return out
    return out
