# Kết nối Vercel với máy Windows lưu phiên

Lỗi “Không tìm thấy phiên xử lý” xuất hiện khi tạo file trong `/tmp` của một Vercel Function nhưng lần xem/sửa chạy ở instance khác hoặc sau khi instance cũ mất. Chỉ sửa progress hoặc tính Excel qua worker không giải quyết việc lưu phiên.

Bản này dùng Vercel phục vụ giao diện; mọi API tạo, xem, sửa nguồn, sửa Family, lịch sử và tải xuống đều chuyển tới **cùng một Shipment Studio trên Windows**. File được lưu trên ổ đĩa Windows; Excel desktop tính lại nguồn ngay tại đó. Không thay đổi thuật toán tạo workbook hoặc định dạng template.

## 1. Chạy backend trên Windows

Cài Microsoft Excel và mở/kích hoạt một lần. Chạy dưới tài khoản Windows đã đăng nhập, giữ máy hoạt động. Trong PowerShell ở thư mục dự án:

```powershell
python -m pip install -r requirements.txt
python -c "import secrets; print(secrets.token_urlsafe(48))"
$env:SHIPMENT_API_TOKEN = 'DAN_KHOA_NGAU_NHIEN_VUA_TAO'
# Thư mục này giữ toàn bộ nguồn, kết quả và lịch sử giữa các lần khởi động.
# Mặc định nếu không đặt: .ui_jobs trong thư mục dự án.
$env:SHIPMENT_DATA_DIR = 'D:\ShipmentStudioData'
.\start_backend.ps1
```

Giữ khóa riêng, không commit vào Git hoặc đưa vào JavaScript. Không đặt `VERCEL`, `SHIPMENT_BACKEND_URL` trên máy Windows. Không cần `EXCEL_WORKER_URL`/`EXCEL_WORKER_TOKEN` nếu Excel chạy ngay trên máy này. Nếu đã đặt `EXCEL_WORKER_URL` cho cách cũ, bỏ biến đó trên Windows để dùng Excel cục bộ.

Lưu ý: biến `$env:` trong ví dụ chỉ tồn tại trong phiên PowerShell hiện tại; khi mở cửa sổ mới cần đặt lại hoặc cấu hình biến môi trường Windows. Đổi thư mục dữ liệu sẽ tạo kho phiên khác; giữ nguyên đường dẫn để mở lại lịch sử. Không xóa thư mục này khi cập nhật mã.

## 2. Cấp HTTPS cho backend

Cấu hình reverse proxy/tunnel HTTPS tới `http://127.0.0.1:8765`. Cần giữ header `Authorization`, cho phép streaming không buffer và thời gian xử lý tối thiểu 300 giây. Dùng URL gốc, không có đường dẫn hoặc redirect. Không mở trực tiếp cổng 8765 ra Internet.

**Hiện chưa có URL HTTPS được cấu hình trong phiên làm việc này.** Đây là bước hạ tầng bắt buộc trước khi Vercel có thể hoạt động. Endpoint cũ cổng 8767 của `excel_worker.py` chỉ tính Excel, không phải backend lưu phiên; không dùng endpoint đó cho cấu hình dưới đây.

## 3. Cấu hình Environment Variables trên Vercel và redeploy

```text
SHIPMENT_BACKEND_URL=https://DOMAIN_BACKEND_CUA_BAN
SHIPMENT_BACKEND_TOKEN=CUNG_KHOA_SHIPMENT_API_TOKEN_TREN_WINDOWS
```

Áp dụng cho đúng môi trường Production/Preview cần dùng. URL và token chỉ dùng phía server. Khi chưa cấu hình, API báo rõ thiếu backend và không tiếp tục tạo phiên tạm sẽ mất sau đó.

Backend trả 401 nếu khóa không khớp. Lỗi 502 thường là máy Windows/tunnel không hoạt động hoặc URL chưa đúng. Truy cập `https://DOMAIN_VERCEL/api/config` sau deploy phải trả JSON chứa `jobs`, `template`; sau đó tạo một file và mở lại link phiên ở tab mới, sửa input, bấm Lưu và tải kết quả.

## 4. Phiên cũ và giới hạn

- Phiên đã mất trong `/tmp` của Vercel không thể phục hồi chỉ từ mã phiên. Sau khi nối backend, tải lại file nguồn (ưu tiên bản đã lưu các chỉnh sửa) để tạo phiên mới.
- Phiên vẫn còn trong `.ui_jobs` trên Windows có thể dùng tiếp nếu giữ nguyên thư mục dữ liệu đó.
- Proxy không loại bỏ giới hạn dung lượng và thời gian của Vercel Functions. File request/response vượt giới hạn 4,5 MB cần truyền trực tiếp qua kho file hoặc chạy giao diện trên backend qua HTTPS có xác thực; không chia chunk vào `/tmp` của Vercel.
- Đây vẫn là workspace dùng chung của người vận hành. Khi mở cho nhiều người, cần thêm đăng nhập/phân quyền trước khi cho truy cập dữ liệu. Khóa backend bảo vệ kết nối server; không thay thế đăng nhập người dùng trên giao diện Vercel.

Tham khảo: [Vercel Function limits](https://vercel.com/docs/functions/limitations).
