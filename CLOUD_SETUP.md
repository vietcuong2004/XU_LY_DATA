# Hướng dẫn triển khai Web: Vercel + Cloudflare R2 (Hoàn toàn Miễn phí)

Tài liệu này hướng dẫn chi tiết cách triển khai ứng dụng lên **Vercel** kết hợp lưu trữ **Cloudflare R2**. 
Giải pháp này **hoàn toàn 0đ/tháng**, không giới hạn 12 tháng như AWS, và **miễn phí 100% băng thông tải xuống (Zero Egress Fees)**.

---

## 1. Tổng quan kiến trúc & Hạn mức miễn phí

- **Vercel (Gói Hobby - Miễn phí):** Chạy giao diện web HTML/JS và Serverless Python API.
- **Cloudflare R2 (Gói Free Tier vĩnh viễn):**
  - **10 GB dung lượng lưu trữ** miễn phí mỗi tháng (lưu được hàng chục nghìn file Excel).
  - **1.000.000 lượt ghi (Class A - PUT, POST, LIST)** miễn phí mỗi tháng.
  - **10.000.000 lượt đọc (Class B - GET)** miễn phí mỗi tháng.
  - **Miễn phí băng thông tải về (Egress $0)**: Không lo bị tính tiền khi người dùng tải nhiều file Excel/ZIP.
- **Luồng hoạt động:** 
  Trình duyệt tải file Excel nguồn trực tiếp lên Cloudflare R2 qua **Signed URL** (tránh giới hạn payload của Vercel). Vercel Function tải file từ R2 về xử lý, tính toán công thức, sinh các Family và lưu kết quả ngược lại R2.

---

## 2. Các bước thiết lập Cloudflare R2

### Bước 2.1: Đăng ký / Đăng nhập Cloudflare
1. Truy cập [dash.cloudflare.com](https://dash.cloudflare.com/) và đăng ký tài khoản miễn phí (nếu chưa có).
2. Tại menu bên trái, chọn **R2** (hoặc **Storage & Databases** → **R2**).
3. Nếu là lần đầu vào R2, Cloudflare có thể yêu cầu kích hoạt dịch vụ R2 (chọn gói Free).

### Bước 2.2: Tạo Bucket trên R2
1. Trong trang quản lý R2, bấm nút **Create bucket**.
2. Đặt tên bucket (ví dụ: `shipment-data` - chỉ dùng chữ thường, số và dấu gạch ngang `-`).
3. Vị trí (Location): Chọn **Automatic** (hoặc khu vực APAC / Châu Á).
4. Bấm **Create bucket**.

### Bước 2.3: Cấu hình CORS cho Bucket
Cấu hình CORS để trình duyệt web có thể tải file trực tiếp lên Cloudflare R2:
1. Trong Bucket vừa tạo (`shipment-data`), chuyển sang tab **Settings**.
2. Kéo xuống mục **CORS Policy**, chọn **Add CORS policy** (hoặc Edit).
3. Dán đoạn JSON sau vào:
   ```json
   [
     {
       "AllowedOrigins": [
         "https://*.vercel.app",
         "http://localhost:8765"
       ],
       "AllowedMethods": [
         "GET",
         "PUT",
         "HEAD"
       ],
       "AllowedHeaders": [
         "*"
       ],
       "ExposeHeaders": [
         "ETag"
       ],
       "MaxAgeSeconds": 3600
     }
   ]
   ```
   > **Lưu ý:** Sau khi Vercel cấp tên miền chính thức của bạn (ví dụ `https://ten-du-an.vercel.app`), bạn có thể thêm chính xác domain đó vào `AllowedOrigins`.
4. Bấm **Save**.

### Bước 2.4: Tạo API Token (Lấy S3 Credentials)
1. Quay lại trang chủ **R2** (menu trái chọn R2 Overview).
2. Ở cột bên phải, tìm mục **Account Details**, copy lại **Account ID** (chuỗi ký tự ví dụ: `a1b2c3d4e5f6...`).
3. Cũng tại cột bên phải, bấm vào link **Manage R2 API Tokens**.
4. Bấm nút **Create API token**:
   - **Token name:** `vercel-shipment-token`
   - **Permissions:** Chọn **Object Read & Write**.
   - **Specify bucket(s):** Chọn **Apply to specific buckets only** và chọn bucket vừa tạo (`shipment-data`), hoặc chọn **Apply to all buckets**.
   - **TTL:** Để mặc định (Forever) trừ khi bạn muốn đổi định kỳ.
5. Bấm **Create API Token**.
6. **LƯU Ý QUAN TRỌNG:** Màn hình sẽ hiển thị thông tin token một lần duy nhất, hãy sao chép lại ngay:
   - **Access Key ID**: Chuỗi ký tự (dùng cho `SHIPMENT_S3_ACCESS_KEY_ID`)
   - **Secret Access Key**: Chuỗi ký tự bí mật (dùng cho `SHIPMENT_S3_SECRET_ACCESS_KEY`)

---

## 3. Cấu hình Biến Môi Trường trên Vercel

1. Đăng nhập vào [Vercel Dashboard](https://vercel.com/dashboard) và chọn dự án của bạn.
2. Vào **Settings** → **Environment Variables**.
3. Thêm lần lượt các biến môi trường sau (áp dụng cho cả Production, Preview, Development):

| Tên biến | Giá trị mẫu | Ghi chú |
| :--- | :--- | :--- |
| `SHIPMENT_S3_BUCKET` | `shipment-data` | Tên bucket bạn đã tạo ở Bước 2.2 |
| `SHIPMENT_S3_REGION` | `auto` | Vùng Cloudflare R2 (để `auto`) |
| `SHIPMENT_S3_ENDPOINT` | `https://<ACCOUNT_ID>.r2.cloudflarestorage.com` | Thay `<ACCOUNT_ID>` bằng Account ID ở Bước 2.4 |
| `SHIPMENT_S3_ACCESS_KEY_ID` | `c123...` | Access Key ID tạo ở Bước 2.4 |
| `SHIPMENT_S3_SECRET_ACCESS_KEY` | `89abc...` | Secret Access Key tạo ở Bước 2.4 |
| `SHIPMENT_SESSION_SECRET` | *(chuỗi ngẫu nhiên 32+ ký tự)* | Khóa bí mật mã hóa cookie phiên (xem lệnh tạo bên dưới) |
| `SHIPMENT_CALCULATOR` | `python` | Dùng bộ tính toán Python tích hợp sẵn |

> **Cách tạo `SHIPMENT_SESSION_SECRET`:**  
> Mở terminal máy tính và chạy lệnh:
> ```powershell
> python -c "import secrets; print(secrets.token_urlsafe(48))"
> ```
> Copy chuỗi kết quả và dán vào giá trị của `SHIPMENT_SESSION_SECRET`.

4. Sau khi thêm đủ các biến, vào tab **Deployments** trên Vercel → bấm nút **...** ở bản deploy mới nhất → chọn **Redeploy** để Vercel nạp các biến môi trường mới.

---

## 4. Kiểm tra hoạt động sau khi triển khai

1. Mở trang web Vercel của bạn trên trình duyệt (hoặc kiểm tra `https://ten-du-an.vercel.app/api/config`):
   - Đảm bảo trả về `"cloud_storage": true`.
2. **Thử tải file Excel:** Kéo thả một file Excel nguồn vào trang web.
   - Quan sát tab *Network* của trình duyệt (F12): File sẽ được `PUT` trực tiếp lên endpoint của Cloudflare R2 (`...r2.cloudflarestorage.com`).
   - Serverless function xử lý và hiển thị tiến trình xử lý mượt mà.
3. **Thử tải kết quả:** Bấm tải từng Family hoặc "Tải tất cả các file", file sẽ được tải về trực tiếp từ R2 qua Signed URL.
4. **Bản nháp & Phiên:** Tải lại trang, phiên làm việc trước đó vẫn còn nguyên vẹn.

---

## 5. Dọn dẹp dữ liệu tự động (Lifecycle Rules - Tùy chọn)

Để giữ dung lượng luôn dưới hạn mức 10 GB miễn phí:
1. Vào Cloudflare Dashboard → R2 → Bucket `shipment-data` → **Settings**.
2. Tìm mục **Object lifecycle rules** → Chọn **Add rule**.
3. Thiết lập:
   - **Rule name:** `auto-delete-temp`
   - **Prefix:** `shipment/uploads/`
   - **Action:** Delete objects sau **1 ngày** (vì file tạm tải lên sau khi xử lý xong không cần giữ lại).
4. Bạn có thể thêm rule tương tự cho `shipment/downloads/` sau **1 ngày**.
5. Thư mục `shipment/jobs/` và `shipment/snapshots/` giữ nguyên nếu bạn muốn lưu lại lịch sử kết quả.
