# Vercel + máy Windows tính Excel

## Trạng thái

Mã nguồn đã có worker và luồng sửa mọi sheet. Chưa có địa chỉ HTTPS công khai của máy Windows, nên chưa kết nối với bản Vercel đang chạy.

Luồng: trình duyệt tải **đủ workbook** lên Vercel → sửa ô ở bất kỳ sheet nào → Vercel gửi workbook và yêu cầu sửa tới worker → Microsoft Excel tính lại toàn bộ công thức → worker trả workbook đã lưu → Vercel tạo và đối chiếu toàn bộ Family → giao diện mở phiên mới.

Mọi thay đổi diễn ra trên bản sao. Phiên trước được giữ trong lịch sử. Nếu Excel, kết nối hoặc tạo Family thất bại, giao diện không công bố bộ kết quả mới. Các lỗi công thức vốn có trong nguồn vẫn được báo, không đổi thành số 0.

## 1. Chuẩn bị máy Windows

- Cài Microsoft Excel desktop, mở Excel một lần để hoàn tất đăng nhập/kích hoạt.
- Chạy worker trong phiên người dùng Windows đã đăng nhập, không dùng tài khoản dịch vụ không có desktop.
- Máy phải hoạt động trong lúc dùng tính năng chỉnh sửa. Worker xử lý một workbook mỗi lần; yêu cầu đồng thời được báo bận để thử lại.
- Cài dependencies bằng `python -m pip install -r requirements.txt` trong thư mục dự án.

Tạo khóa riêng (chạy trên máy của bạn, không commit khóa vào Git):

```powershell
python -c "import secrets; print(secrets.token_urlsafe(48))"
$env:EXCEL_WORKER_TOKEN = 'DAN_KHOA_VUA_TAO_VAO_DAY'
python excel_worker.py --host 127.0.0.1 --port 8767
```

Giữ cửa sổ worker chạy. Có thể dùng `start_excel_worker.bat` khi biến môi trường đã được cấu hình.

## 2. Cấp địa chỉ HTTPS

Cần reverse proxy hoặc tunnel HTTPS chuyển tiếp tới `http://127.0.0.1:8767` trên máy Windows. Hiện chưa cấu hình hoặc xuất bản tunnel nào. Endpoint phải giữ header `Authorization`, hỗ trợ POST nhị phân và timeout ít nhất 240 giây; không chuyển hướng sang domain khác.

Chỉ public qua HTTPS, không mở trực tiếp cổng 8767 ra Internet. Worker yêu cầu khóa cho mọi lần tính lại và tắt macro, sự kiện Excel, cập nhật liên kết ngoài khi mở workbook. Dịch vụ này dành cho workbook tin cậy của người vận hành, không phải dịch vụ Excel công cộng nhận file tùy ý.

## 3. Cấu hình Vercel

Thêm vào Environment Variables **phía server**, không đưa vào JavaScript trình duyệt:

```text
EXCEL_WORKER_URL=https://DIA_CHI_HTTPS_CUA_MAY_WINDOWS
EXCEL_WORKER_TOKEN=CUNG_KHOA_DA_CAU_HINH_TREN_WINDOWS
```

Redeploy sau khi có địa chỉ và khóa. Giao diện tự mở quyền sửa mọi sheet khi có cấu hình engine. Có cấu hình chưa có nghĩa là worker đang online: lần lưu chỉ thành công khi nhận lại workbook thực tế và đối chiếu đầu ra đạt.

Trên máy Windows worker **không đặt EXCEL_WORKER_URL**. Worker luôn dùng Excel tại máy đó.

## 4. Sử dụng

1. Tải lại workbook nguồn đầy đủ. Phiên cũ từng chỉ tải SUM không chứa các sheet còn lại; cần tạo phiên mới từ file gốc.
2. Mở file nguồn, chọn sheet, bấm Toàn màn hình nếu cần.
3. Bấm **Chỉnh sửa** cạnh **Toàn màn hình**, nhập trực tiếp vào nhiều ô, chuyển sheet hoặc dán một vùng từ Excel. Ô chưa lưu có màu vàng. Công thức bắt đầu bằng `=`; thêm dấu nháy đơn trước nội dung để giữ dạng văn bản. Bỏ trống để xóa nội dung.
4. Bấm **Lưu** sau khi sửa xong; toàn bộ ô được gửi tới Excel trong một lần và chỉ tính lại một lần. Bấm **Hủy** để bỏ bản nháp. Với workbook mẫu 14 sheet/26 Family, lần kiểm thử thực tế mất khoảng 128 giây; thời gian thật phụ thuộc workbook và máy Windows.
5. Giao diện chuyển sang phiên mới với tất cả Family vừa tạo. Tải nguồn riêng bằng nút tải ở file nguồn để nhận workbook đã tính lại; ZIP tải toàn bộ Family.

Nguồn là chuẩn: các sửa riêng trên Family của phiên trước được giữ trong phiên trước, không tự ghi đè số liệu vừa tính từ nguồn. Có thể sửa tiếp Family sau khi hoàn tất cập nhật nguồn.

Sửa nội dung ô và công thức ở mọi sheet; không bao gồm chèn/xóa sheet, chèn dòng/cột, bỏ bảo vệ sheet, hoặc sửa một phần công thức mảng. Công thức liên kết tới workbook/file bên ngoài vẫn cần dữ liệu bên ngoài; worker không tự tải hay cập nhật các liên kết đó.

## 5. Điều kiện vận hành Vercel còn lại

- Function phải có thời gian chạy đủ cho Excel và tạo Family; thao tác kéo dài cần job queue/worker bền vững.
- Phiên hiện vẫn nằm trong `/tmp/.ui_jobs` trên Vercel. `/tmp` không phải kho lưu bền vững dùng chung giữa các instance. Cần object storage/database cho vận hành nhiều người hoặc giữ lịch sử lâu dài.
- Upload/download toàn workbook vẫn chịu giới hạn request/response Vercel. Workbook lớn cần upload trực tiếp vào object storage, không cắt bỏ sheet để giảm dung lượng.
- Worker này xử lý tính toán; chưa thay thế lưu trữ, xác thực người dùng hay hàng đợi của ứng dụng.

## Kiểm thử

```powershell
python -m unittest discover -s test -v
$env:RUN_EXCEL_TESTS = '1'
python -m unittest test.test_full_source.DesktopExcelTests -v
```

Kiểm thử Excel dùng workbook tạm: chuỗi `Injection → Production → SUM` và luồng HTTP worker `Bel!C16 → SUM!W16 → 26 Family`. Không chỉnh file nguồn gốc của dự án.

Tham khảo: [CalculateFullRebuild của Microsoft](https://learn.microsoft.com/en-us/office/vba/api/excel.application.calculatefullrebuild), [giới hạn Vercel Functions](https://vercel.com/docs/functions/limitations).
