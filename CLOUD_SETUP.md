# Bản web: Python + S3, không cần máy cá nhân chạy liên tục

## Luồng hiện tại

Vercel phục vụ giao diện và chạy Python. File nguồn tải trực tiếp vào S3 riêng tư bằng URL có hạn. Function tải nguồn, tạo và đối chiếu Family, lưu cả phiên vào S3 rồi mới báo thành công. Instance khác sẽ khôi phục phiên từ S3; `/tmp` chỉ là bản làm việc. Excel/ZIP cũng tải trực tiếp qua signed URL, tránh giới hạn payload Vercel.

Khi bấm **Lưu** trên nguồn, Python tính lại các sheet rồi tạo toàn bộ Family trong phiên mới. Bộ tính hỗ trợ phép toán, phần trăm, so sánh, tham chiếu ô/vùng giữa các sheet, hợp vùng, `SUM` và `WEEKNUM`. Hàm chưa hỗ trợ, công thức mảng, liên kết ngoài và tham chiếu vòng sẽ chặn lưu và giữ phiên trước. Đây là bộ tính dành cho lịch hiện tại, không thay thế mọi tính năng Excel.

IndexedDB giữ bản nháp; mở lại cùng file trong cùng phiên sẽ khôi phục bản nháp có cùng revision. Cookie HttpOnly có chữ ký tách workspace mỗi trình duyệt, tối đa một năm. Xóa cookie/đổi trình duyệt sẽ tạo workspace khác: phiên vẫn còn trên S3 nhưng chưa có đăng nhập để lấy lại trên thiết bị khác.

## Biến môi trường Vercel

Tạo bucket AWS S3 **Private**, bật Block Public Access. Cấu hình cho đúng môi trường rồi Redeploy:

```text
SHIPMENT_S3_BUCKET=ten-bucket
SHIPMENT_S3_REGION=ap-southeast-1
SHIPMENT_S3_ACCESS_KEY_ID=...
SHIPMENT_S3_SECRET_ACCESS_KEY=...
SHIPMENT_SESSION_SECRET=chuoi-ngau-nhien-it-nhat-32-ky-tu
SHIPMENT_CALCULATOR=python
```

Tạo secret bằng `python -c "import secrets; print(secrets.token_urlsafe(48))"`. Giữ nguyên secret qua các lần deploy để cookie cũ còn hiệu lực. Không đưa credentials vào JavaScript hay Git. Không đặt `SHIPMENT_API_TOKEN` trên Vercel (biến của backend Windows cũ). `SHIPMENT_BACKEND_URL` và `EXCEL_WORKER_URL` không cần trong chế độ S3 + Python.

Tùy chọn: `SHIPMENT_S3_PREFIX=shipment`, `SHIPMENT_S3_ENDPOINT=https://...` cho dịch vụ tương thích S3 có hỗ trợ signed URL và conditional writes/deletes `If-None-Match`, `If-Match`. Khuyến nghị kiểm chứng trên AWS S3 trước. Chưa có bucket/credentials của bạn để kiểm thử dịch vụ thật; kiểm thử lưu/khôi phục dùng S3 giả lập.

Quyền IAM: `s3:ListBucket` trên bucket; `s3:GetObject`, `s3:PutObject`, `s3:DeleteObject` trên prefix dự án. Không cấp public ACL. Mã không tự tạo dịch vụ có phí; chi phí lưu trữ/truyền dữ liệu thuộc tài khoản cloud của bạn.

## CORS của bucket

AWS Console → S3 → bucket → Permissions → CORS. Thay origin đúng domain website, không có dấu `/` cuối:

```json
[
  {
    "AllowedOrigins": ["https://TEN-DU-AN.vercel.app"],
    "AllowedMethods": ["PUT", "GET", "HEAD"],
    "AllowedHeaders": ["*"],
    "ExposeHeaders": ["ETag"],
    "MaxAgeSeconds": 3600
  }
]
```

Signed URL có hạn 10 phút. Thêm origin Preview cụ thể khi cần. File tối đa 45 MB được kiểm tra trước giải nén. Bucket chứa `shipment/jobs`, `shipment/snapshots`, `shipment/uploads`, `shipment/downloads`, `shipment/locks`; mỗi prefix có workspace riêng.

Thiết lập lifecycle xóa `shipment/uploads/` và `shipment/downloads/` sau 1 ngày. Không tự xóa `jobs/` hoặc `snapshots/` nếu cần giữ lịch sử. Xóa phiên trong app xóa manifest; các snapshot còn lại cần quy trình retention quản trị riêng. Khóa `locks/` ngăn hai request cùng lưu một phiên. Nếu Function bị dừng cưỡng bức, khóa có thể được giành lại sau 10 phút bằng conditional write; không xóa khóa của request khác. Giữ giới hạn Function 300 giây như `vercel.json`.

## Thời gian và cập nhật

Python xử lý trong Function hiện tại và gửi progress trực tiếp. Cần Fluid Compute và đủ thời gian thực thi (mục tiêu 300 giây). Chưa có hàng đợi chạy tác vụ độc lập: workbook vượt giới hạn thời gian cần chuyển phần xử lý sang worker có hàng đợi. Không chạy Excel desktop hoặc phụ thuộc máy Windows.

Deploy giao diện và server cùng nhau. Giữ bucket, prefix, session secret ổn định. Dữ liệu S3 không phụ thuộc vòng đời bản deploy. Phiên cũ đã mất trong `/tmp` cần tải lại nguồn.

## Kiểm tra sau deploy

1. `/api/config` trả `cloud_storage: true`; browser có cookie `shipment_workspace`.
2. Upload >4,5 MB phải PUT trực tiếp tới S3; POST `/api/jobs` chỉ chứa JSON nhỏ với `upload_key`.
3. Tạo xong, tải lại trang, mở lại phiên, tải Excel/ZIP qua signed URL.
4. Sửa ô trên sheet thị trường, Lưu và đối chiếu SUM/Family. Công thức chưa hỗ trợ phải giữ bản nháp và không công bố kết quả.
5. Nhập vài ô chưa lưu, tải lại trang, mở lại đúng file: bản nháp được khôi phục.
6. Trình duyệt khác không được thấy lịch sử workspace trước.

Kiểm thử: `uv sync --frozen`, `uv run python -m unittest discover -s test`, `node --test test/test_inline_editor.cjs test/test_progress.cjs test/test_session_ui.cjs`.

Nguồn: [S3 CORS](https://docs.aws.amazon.com/AmazonS3/latest/userguide/cors.html), [WEEKNUM](https://support.microsoft.com/en-us/excel/functions/weeknum-function), [Vercel Function limits](https://vercel.com/docs/functions/limitations).
