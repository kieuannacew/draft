# Hướng dẫn tạo máy chủ Lớp học (Supabase – miễn phí)

Làm **một lần**, khoảng 10 phút. Sau đó học sinh tải app về laptop, đăng ký bằng **mã lớp**. Mỗi lần làm bài xong,
điểm tự gửi lên máy chủ và giáo viên xem được ở mục **Lớp học**.

> Không cần biết lập trình. Chỉ cần một tài khoản Google hoặc GitHub để đăng nhập Supabase.

---

## Bước 1 – Tạo dự án Supabase

1. Vào **https://supabase.com** → bấm **Start your project** → đăng nhập bằng GitHub hoặc Google.
2. Bấm **New project**:
   - **Name**: `luyen-thi-mos` (tùy ý)
   - **Database Password**: bấm *Generate a password*, **lưu lại** mật khẩu này (app không dùng, nhưng cần khi khôi phục)
   - **Region**: chọn **Southeast Asia (Singapore)**, gần Việt Nam nên nhanh nhất
   - Gói **Free**
3. Bấm **Create new project**, rồi đợi khoảng 1–2 phút cho dự án chạy xong.

## Bước 2 – Cài đặt dữ liệu (chạy 1 file SQL)

1. Menu bên trái → **SQL Editor** → **New query**.
2. Trong thư mục app, mở file **`may_chu\supabase_lop_hoc.sql`** bằng Notepad → **Ctrl+A**, **Ctrl+C**.
3. Dán vào ô soạn thảo của Supabase → bấm **Run**. Thấy dòng *Success. No rows returned* là xong.

(Chạy lại file này nhiều lần vẫn an toàn, ví dụ khi app có bản cập nhật máy chủ mới.)

## Bước 3 – Tắt xác nhận email

Học sinh đăng ký bằng tên đăng nhập, không dùng email thật, nên phải tắt bước xác nhận email:

1. Menu bên trái → **Authentication** → **Sign In / Providers** (bản cũ: *Providers*) → **Email**.
2. **Tắt** mục **Confirm email** → **Save**.
3. Kiểm tra mục **Allow new users to sign up** đang **bật**.

## Bước 4 – Lấy địa chỉ và khóa

Menu bên trái → **Project Settings** → **API** (hoặc **Data API / API Keys**), chép 2 thứ:

- **Project URL**, dạng `https://abcdxyz.supabase.co`
- Khóa công khai: **anon public** (dạng `eyJhbGci…`) hoặc **publishable** (dạng `sb_publishable_…`)

> ⚠️ **Không** dùng khóa `service_role` / `secret`. Khóa đó có toàn quyền, tuyệt đối không đưa vào app.

## Bước 5 – Kết nối app với máy chủ

1. Mở app (`run.bat`). Ở màn đăng nhập, bấm **⚙ Máy chủ lớp học** (góc dưới bên phải).
2. Dán **Project URL** và **khóa** → bấm **Kiểm tra kết nối**. Thấy *✓ Kết nối tốt* là được.
3. Bấm **Lưu**. App ghi cấu hình vào file **`cau_hinh.json`** cạnh `run.bat`.

## Bước 6 – Tạo tài khoản quản trị (làm NGAY)

Ở màn đăng nhập, bấm **Chưa có tài khoản? Đăng ký bằng mã lớp**, rồi đăng ký cho chính bạn (để trống mã lớp).
**Người đăng ký đầu tiên tự động là Quản trị.** Vì vậy hãy tự đăng ký trước khi gửi app cho học sinh.

## Bước 7 – Giáo viên và lớp

- **Thêm giáo viên**: giáo viên tự đăng ký giống học sinh. Quản trị vào **Tài khoản** → chọn người đó →
  **Đổi vai trò…** → **Giáo viên**.
- **Tạo lớp**: giáo viên (hoặc quản trị) vào **🏫 Lớp học** → **+ Tạo lớp** → app cho **mã lớp 6 ký tự**, ví dụ `K7M2QX`.
- Gửi mã lớp cho học sinh qua Zalo hoặc bảng lớp.

## Bước 8 – Gửi app cho học sinh

Nén **cả thư mục app, có file `cau_hinh.json`**, thành ZIP rồi gửi cho học sinh (qua Google Drive, Zalo…). Học sinh:

1. Giải nén → chạy **`run.bat`**. App tự kết nối máy chủ, không phải nhập gì thêm.
2. Bấm **Đăng ký bằng mã lớp** → nhập họ tên, tên đăng nhập, mật khẩu (ít nhất 6 ký tự) và **mã lớp**.
3. Học và làm bài như bình thường. Điểm tự gửi cho giáo viên.

(Nếu dùng bản `.exe`: đặt `cau_hinh.json` cạnh file `LuyenThiMOS.exe`.)

---

## Giáo viên xem được gì?

Mục **🏫 Lớp học**:

- Danh sách lớp, **mã lớp** (nút sao chép), sĩ số, lượt làm bài, tỉ lệ đạt, điểm trung bình.
- Bảng học sinh: số lần làm, số lần đạt, **điểm cao nhất Word / Excel / PowerPoint**, lần làm gần nhất, số bài giảng
  đã học xong.
- **Xem chi tiết** một học sinh: mọi lần làm bài, các câu sai ở lần gần nhất, tiến độ từng bài giảng (và điểm bài SCORM).
- **Xuất Excel** bảng điểm cả lớp. **Đặt lại mật khẩu** khi học sinh quên. **Xóa khỏi lớp**.

Quyền xem được máy chủ kiểm tra, học sinh không thể "lách" bằng cách sửa app:

| Ai | Thấy gì |
|---|---|
| Học sinh | chỉ điểm và tiến độ của chính mình |
| Giáo viên | học sinh trong **lớp mình** |
| Quản trị | tất cả; cấp quyền giáo viên, khóa / xóa tài khoản |

## Khi mất mạng

- Học sinh đã đăng nhập máy chủ ít nhất một lần thì lúc mất mạng **vẫn đăng nhập và học được**.
- Điểm và tiến độ được lưu trên máy (`cho_dong_bo.json`), **tự gửi** lần sau đăng nhập lúc có mạng.
  Học sinh cũng có thể bấm **Gửi ngay** trong mục Lớp học.

## Gói miễn phí có đủ không?

Đủ cho vài nghìn học sinh: 500 MB dữ liệu (mỗi lần làm bài chỉ vài KB) và 50.000 người dùng mỗi tháng.
**Lưu ý:** dự án miễn phí **tự tạm dừng nếu 7 ngày liền không ai dùng** (ví dụ nghỉ hè). Khi đó vào supabase.com,
mở dự án, bấm **Restore**. Dữ liệu vẫn còn.

## Gặp lỗi?

| Thông báo trong app | Cách xử lý |
|---|---|
| *Khóa máy chủ (anon key) không đúng* | Chép lại khóa ở Bước 4 (anon hoặc publishable, **không** phải service_role) |
| *Máy chủ chưa được cài đặt…* | Làm lại Bước 2 (chạy file SQL) |
| *Máy chủ đang bật xác nhận email…* | Làm lại Bước 3 |
| *Không kết nối được máy chủ lớp học* | Kiểm tra Internet; kiểm tra dự án có bị tạm dừng (Restore) |
| *Tên đăng nhập này đã có người dùng* | Chọn tên đăng nhập khác |
| *Mã lớp không đúng* | Hỏi lại giáo viên mã lớp (6 ký tự) |

**Tắt chế độ Lớp học**: màn đăng nhập → ⚙ Máy chủ lớp học → **Tắt máy chủ**. App quay về dùng tài khoản trên
từng máy như trước.
