# Triển khai Shipment Studio

## Lỗi Vercel không tìm thấy Python entrypoint

Repository có `ui_server.py` với lớp `Handler`, nhưng tên file này không nằm trong danh sách entrypoint mặc định của Vercel. `pyproject.toml` ở thư mục gốc khai báo rõ:

```toml
[tool.vercel]
entrypoint = "ui_server:Handler"
```

Đưa file này lên cùng nhánh đang deploy. Root Directory trên Vercel phải là thư mục chứa `pyproject.toml`, `ui_server.py` và `requirements.txt`.

Đây là sửa cấu hình phát hiện entrypoint, **chưa phải bản chuyển đổi ứng dụng sang Vercel hoàn chỉnh**. Chưa xác nhận bằng một lần build/deploy trên tài khoản Vercel.

## Những phần cần chuyển đổi trước khi dùng trên Vercel

- `Handler.guard()` hiện chỉ nhận Host/Origin localhost. Domain triển khai cần được cho phép rõ ràng.
- File nguồn, kết quả và lịch sử hiện được lưu trong `.ui_jobs` cạnh mã nguồn. Cần nơi lưu bền vững, dùng chung giữa các lần gọi function; chỉ chuyển sang `/tmp` không bảo đảm mở lại hay chỉnh sửa được phiên.
- `create_job()` hiện trả về ngay và chạy `generate()` trong thread nền. Cần cơ chế xử lý có vòng đời phù hợp với function hoặc hàng đợi bền vững.
- Upload hiện gửi JSON/base64, tối đa 45 MB mỗi file. Vercel Functions giới hạn payload request/response 4.5 MB: cần upload/download trực tiếp qua kho file cho các file lớn, đồng thời tránh trả báo cáo hoặc preview quá lớn.
- Khi phục vụ nhiều người dùng, danh sách phiên và quyền đọc/sửa/tải file cần được tách theo người dùng. API hiện được thiết kế cho ứng dụng localhost một người dùng.

Có hai hướng: giữ Vercel và bổ sung lưu trữ, xử lý phiên phù hợp; hoặc chạy Python trên máy chủ liên tục với ổ đĩa bền vững và cấu hình truy cập qua domain. Không nên coi build thành công là đã kiểm chứng toàn bộ luồng tải lên → tạo file → sửa → tải về.

Tài liệu chính thức: [Python entrypoints](https://vercel.com/docs/functions/runtimes/python#python-entrypoints), [Function limits](https://vercel.com/docs/functions/limitations).
