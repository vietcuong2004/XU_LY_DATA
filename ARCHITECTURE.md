# Cấu trúc ứng dụng

Ứng dụng dùng kiến trúc module trong một dịch vụ Python. Chưa tách thành nhiều
microservice: cần ổn định giao diện, phiên xử lý và tính đúng Excel trước khi thêm
vận hành mạng, hàng đợi và lưu trữ dùng chung.

## Giao diện

`ui/app.entry.js` là điểm kết nối các module và khởi động trang. Chạy
`node scripts/build_ui.mjs` để tạo `ui/app.js` đã gộp cho trình duyệt. Cách này
giữ mã nguồn dễ quản lý nhưng vẫn chạy được trên server cũ chỉ phục vụ `app.js`.

| Module | Trách nhiệm |
| --- | --- |
| `state.mjs` | Tạo trạng thái riêng cho mỗi workspace |
| `api.mjs` | Gọi API, chuyển lỗi HTTP, nhận AbortSignal |
| `week-picker.mjs` | Chọn tuần, phân tích nhãn tuần và khoảng ngày |
| `file-list.mjs` | Danh sách Family, bộ lọc, trạng thái đã xem/tải |
| `workbook-controller.mjs` | Chọn file/sheet, quản lý yêu cầu preview, đăng ký sự kiện |
| `preview-view.mjs` | Tiêu đề, tabs, trạng thái tải và lỗi preview |
| `grid-view.mjs` | Render bảng, ô gộp, độ rộng cột và trạng thái ô |

Controller chỉ phụ thuộc API và các callback của view. Mỗi lần chuyển file/sheet
tạo một mã yêu cầu mới và hủy fetch cũ. Chỉ kết quả cùng mã yêu cầu, phiên, file,
sheet hiện tại được hiển thị. Một nơi duy nhất đăng ký click của danh sách file.
Gọi `bind()` lại sẽ tháo listener cũ trước khi đăng ký mới.

Thêm module vào danh sách cho phép trong `ui_server.py` để server phục vụ đúng
MIME JavaScript. Không mở truy cập tùy ý tới đường dẫn hệ thống.

Kiểm thử: `node --test test/test_workbook_controller.mjs test/test_ui_bootstrap.mjs test/test_progress.cjs`.
Kiểm thử route Python: `python -m unittest test.test_module_routes -v`.
Kiểm thử khởi tạo đọc ID từ HTML thật để phát hiện đăng ký sự kiện vào phần tử
đã bị bỏ. Thành phần tùy chọn như lịch sử phải được kiểm tra tồn tại trước khi dùng.

Sau khi sửa module hoặc entrypoint, luôn chạy build UI và commit cả `ui/app.js`.
Không cần khởi động lại Python server khi chỉ đổi giao diện.

## Phần cần tách tiếp khi mở rộng

Upload, lịch sử phiên, chỉnh sửa và preview định dạng vẫn được kết nối trong
`app.js`. Tách tiếp theo cùng cách: factory nhận phụ thuộc, trả phương thức công
khai, đăng ký sự kiện tại bước khởi động. Không tạo thêm biến toàn cục hay handler
cho chức năng đã có controller sở hữu.

Python giữ các vai trò hiện có: `ui_server.py` xử lý HTTP, `process.py` tạo Excel,
`template_format.py` định dạng template, `source_edit.py` điều phối sửa nguồn.

Nếu thời gian xử lý vượt giới hạn web hoặc lượng người dùng tăng, phần tạo Excel
có thể chuyển thành worker riêng: API tạo job → hàng đợi → worker xử lý → kho file
dùng chung. Cần có lưu trữ phiên bền vững, retry/idempotency và API tiến độ trước
khi triển khai. Đây là hướng mở rộng, chưa phải microservice đã được triển khai.
