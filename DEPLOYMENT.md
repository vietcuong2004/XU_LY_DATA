# Triển khai Shipment Studio

**Hướng triển khai hiện tại:** [CLOUD_SETUP.md](CLOUD_SETUP.md). Bộ tính Python và kho S3 thay cho máy Windows luôn bật. Các phần backend Windows dưới đây là phương án cũ/tùy chọn.

## Cấu hình đang dùng: Vercel + backend Windows lưu phiên

Làm theo [WINDOWS_BACKEND_SETUP.md](WINDOWS_BACKEND_SETUP.md). Đặt `SHIPMENT_BACKEND_URL` và `SHIPMENT_BACKEND_TOKEN` trên Vercel; máy Windows chạy `ui_server.py` với `SHIPMENT_API_TOKEN` và thư mục dữ liệu bền vững `SHIPMENT_DATA_DIR`. Toàn bộ API được chuyển tới Windows, gồm tạo phiên, tiến độ, đọc, sửa nguồn/Family và tải file. Vercel thiếu cấu hình sẽ báo 503, không tiếp tục tạo phiên trong `/tmp`.

Phiên đã mất khỏi `/tmp` không thể khôi phục bằng URL; cần tải lại file nguồn sau khi backend đã kết nối. Mã đã được kiểm thử cục bộ, chưa có endpoint HTTPS để kiểm chứng trên bản deploy thật.

## Sửa mọi sheet với Microsoft Excel

Với backend Windows ở trên, Excel desktop tính lại ngay trên máy lưu phiên. [EXCEL_WORKER_SETUP.md](EXCEL_WORKER_SETUP.md) mô tả lựa chọn tách riêng máy tính Excel; chỉ cấu hình worker đó không giải quyết lưu trữ phiên trên Vercel.

## Lỗi Vercel không tìm thấy Python entrypoint

Repository có `ui_server.py` với lớp `Handler`, nhưng tên file này không nằm trong danh sách entrypoint mặc định của Vercel. `pyproject.toml` ở thư mục gốc khai báo rõ:

```toml
[tool.vercel]
entrypoint = "ui_server:Handler"
```

`pyproject.toml` cũng phải có bảng `[project]` để Vercel chạy `uv lock`: tên dự án, phiên bản, `requires-python` và `dependencies`. Bản hiện tại chọn Python 3.12 và khai báo `openpyxl`, `pandas` với cùng giới hạn phiên bản như `requirements.txt`. `[tool.uv] package = false` xác định đây là ứng dụng chạy trực tiếp, không cần đóng gói thành thư viện.

Đưa cả `pyproject.toml` và `uv.lock` lên cùng nhánh đang deploy, rồi Redeploy. Root Directory trên Vercel phải là thư mục chứa `pyproject.toml`, `ui_server.py` và `requirements.txt`. Khi đổi dependencies, cập nhật cả `requirements.txt` và chạy lại `uv lock`.

Đây là sửa cấu hình phát hiện entrypoint, **chưa phải bản chuyển đổi ứng dụng sang Vercel hoàn chỉnh**. Chưa xác nhận bằng một lần build/deploy trên tài khoản Vercel.

## Giới hạn của cách chạy xử lý trực tiếp trong Vercel Function (không dùng nữa)

- `Handler.guard()` hiện chỉ nhận Host/Origin localhost. Domain triển khai cần được cho phép rõ ràng.
- File nguồn, kết quả và lịch sử hiện được lưu trong `.ui_jobs` cạnh mã nguồn. Cần nơi lưu bền vững, dùng chung giữa các lần gọi function; chỉ chuyển sang `/tmp` không bảo đảm mở lại hay chỉnh sửa được phiên.
- Trên Vercel, `create_job()` chờ xử lý xong trong request hiện tại, tránh thread nền bị treo sau khi function trả về. Tác vụ vẫn phải hoàn tất trong giới hạn thời gian function; file lớn cần hàng đợi/worker bền vững. Local tiếp tục dùng thread nền và polling tiến độ.
- Upload dùng multipart nhị phân (giảm khoảng 25% dung lượng so với JSON/base64), tối đa 45 MB mỗi file ở ứng dụng local. Vercel Functions vẫn giới hạn payload request/response 4.5 MB: cần upload/download trực tiếp qua kho file cho các file lớn. Nén gzip báo cáo/preview không thay thế việc xử lý giới hạn nền tảng.
- Khi phục vụ nhiều người dùng, danh sách phiên và quyền đọc/sửa/tải file cần được tách theo người dùng. API hiện được thiết kế cho ứng dụng localhost một người dùng.

Có hai hướng: giữ Vercel và bổ sung lưu trữ, xử lý phiên phù hợp; hoặc chạy Python trên máy chủ liên tục với ổ đĩa bền vững và cấu hình truy cập qua domain. Không nên coi build thành công là đã kiểm chứng toàn bộ luồng tải lên → tạo file → sửa → tải về.

Tài liệu chính thức: [Python entrypoints](https://vercel.com/docs/functions/runtimes/python#python-entrypoints), [Function limits](https://vercel.com/docs/functions/limitations).

## Tối ưu hiệu năng đã thực hiện

- Đã tắt tối ưu cắt workbook chỉ còn SUM. Tính năng sửa mọi sheet cần giữ đầy đủ công thức và dữ liệu nguồn. Upload hiện gửi nguyên workbook; dung lượng lớn cần object storage thay vì bỏ sheet.
- Trình duyệt gửi nguyên bytes Excel bằng FormData, không đọc và chuyển cả file sang base64. Thanh tiến độ hiển thị phần trăm upload thực tế; tiến độ xử lý Family vẫn tách riêng.
- JSON và nội dung text lớn được nén gzip khi trình duyệt hỗ trợ. File Excel/ZIP tải xuống giữ nguyên bytes.
- Khi tạo phiên, mở mỗi workbook kết quả một lần cho cả đối chiếu SUM và phân tích lỗi. Bản baseline vừa sao chép có cùng bytes nên dùng chung workbook đã đọc. Khi sửa dữ liệu, vẫn đọc baseline riêng để so sánh đúng.
- Không sửa thuật toán tạo Excel, công thức, định dạng hay bỏ phép kiểm tra nguồn.

Đo cục bộ trên 26 Family: bước đối chiếu/phân tích sau xuất từ 4,143 giây xuống 1,614 giây; JSON báo cáo mẫu từ 3.097.217 bytes xuống 174.966 bytes khi gzip. Đây không phải số đo end-to-end trên Vercel; tốc độ còn phụ thuộc đường truyền, cold start và tài nguyên function.

Nếu riêng sheet `SUM` sau khi đóng gói vẫn vượt giới hạn request 4,5 MB của Vercel, bước tiếp theo bắt buộc là client upload trực tiếp lên Vercel Blob hoặc object storage tương đương. Chia chunk rồi lưu vào `/tmp` của nhiều Function không an toàn vì các request không được bảo đảm chạy trên cùng instance.
