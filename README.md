# Shipment Studio · Tool Bóc Tách Kế Hoạch Excel Theo Family

Bản web mới dùng **Python để tính lại nguồn** và **S3 để lưu phiên bền vững**, không cần Excel desktop hoặc máy cá nhân luôn bật. Xem [hướng dẫn cấu hình cloud](CLOUD_SETUP.md). Có bản nháp IndexedDB và upload/download trực tiếp S3. Công thức chưa hỗ trợ sẽ báo lỗi và giữ phiên cũ.

Công cụ tự động hóa bóc tách ma trận kế hoạch xuất hàng tổng thể (**Master Shipment Schedule - sheet `SUM`**) thành các file Excel chuẩn theo từng dòng sản phẩm (**Family**), phục vụ công tác điều hành sản xuất và logistics tại **Niigata**.

---

## 1. Hướng dẫn sử dụng Giao diện (Shipment Studio UI)

### 1.1. Khởi động ứng dụng
* **Cách 1 (Nhanh nhất):** Nhấp đúp chuột vào file [`start_ui.bat`](file:///d:/XU_LY_DATA/start_ui.bat).
* **Cách 2 (Dòng lệnh):**
  ```powershell
  python ui_server.py --open
  ```
  Ứng dụng sẽ tự động khởi chạy tại cổng nội bộ và mở trình duyệt tại: **`http://127.0.0.1:8765`** (nhấn `Ctrl + C` tại cửa sổ dòng lệnh để tắt server).

---

### 1.2. Thao tác trên 2 Tab giao diện

Giao diện được phân chia thành **2 Tab nghiệp vụ trực quan**:

#### Tab 1: "Tải file lên" (↥)
1. **Chọn file nguồn (.xlsx):** Kéo thả hoặc bấm vào khung tải lên để chọn file Master (ví dụ: `2026 INTERNAL SCHEDULE FERRERO-WK40(LOG)-T.xlsx`). Nếu file đã nằm sẵn trong thư mục làm việc, có thể bấm nút *"Dùng file nguồn có sẵn"*.
2. **Bấm "Tạo các file Family →":** Trình duyệt tải đủ workbook để giữ mọi sheet và công thức liên kết. Hệ thống tự dùng template BARBIE mặc định, giữ toàn bộ khoảng tuần trong nguồn và bóc tách tất cả Family. Tuần đặt tên file được nhận từ `WKxx` trong tên file nguồn; nếu không tìm thấy thì dùng `WK40`. Mùa đặt tên file được nhận từ template, mặc định `2728`. Hai giá trị này chỉ dùng trong tên file kết quả.
3. Quá trình hoàn tất sẽ **tự động chuyển ngay sang Tab 2**. Các số Family, item, thị trường và tuần hiển thị tại đây đều là kết quả đọc thật sau khi xử lý, không phải số minh họa.
4. **Mở lại kết quả gần đây:** Cho phép mở lại tức thì các phiên làm việc trước đó mà không cần tạo lại từ đầu.

#### Tab 2: "Sửa & xuất kết quả" (▦)
1. **Xem tổng quan (Metrics):** Thống kê số lượng file Family đã tạo (kèm huy hiệu trên Tab), số tuần kế hoạch, số Family cần kiểm tra cảnh báo và số Family đã chỉnh sửa.
2. **Danh sách Family & Bộ lọc:**
   - Cột bên trái hiển thị danh sách toàn bộ các Family tìm thấy (ví dụ 26 Family).
   - File nguồn nằm đầu danh sách. Khi đã kết nối máy xử lý Microsoft Excel, có thể sửa nội dung ô và công thức ở mọi sheet. Excel tính lại toàn workbook, rồi tool tạo lại và đối chiếu toàn bộ Family. Xem [cấu hình Windows worker cho Vercel](EXCEL_WORKER_SETUP.md).
   - **Tải file nguồn** trả workbook hiện tại đã tính lại; ZIP tải toàn bộ Family. Mỗi lần lưu nguồn tạo một phiên mới; phiên trước được giữ trong lịch sử. Các chỉnh sửa riêng trên Family cũ không áp dụng sang bộ mới tạo từ nguồn. Nếu chưa có engine, giao diện báo rõ và khóa sửa nguồn.
   - Nút **"Toàn màn hình"** mở rộng bảng; nút **"Chỉnh sửa"** bên cạnh bật sửa trực tiếp. Bấm **"Thu nhỏ"** hoặc `Esc` để trở lại.
   - Hỗ trợ ô tìm kiếm nhanh và các bộ lọc: *Tất cả*, *Cần xem*, *Đã sửa*.
3. **Bảng lưới tính Excel (Interactive Grid Viewer):**
   - Chuyển đổi linh hoạt giữa 3 sheet: `Breakdown `, `[Tên Family]`, `Release Qty`.
   - Hiển thị màu sắc trạng thái: Xanh lá (Khớp nguồn), Xanh lam (Đã sửa), Đỏ (Lỗi dữ liệu).
4. **Chỉnh sửa trực tiếp trên bảng:**
   - Bấm **Chỉnh sửa**, nhấp ô để nhập. Có thể sửa nhiều ô, chuyển sheet trong cùng file hoặc dán một vùng từ Excel. Enter/Tab chuyển ô.
   - Các ô chưa lưu có màu vàng. Mọi thay đổi chỉ nằm trong bản nháp cho tới khi bấm **Lưu**; **Hủy** bỏ toàn bộ bản nháp của file đang sửa.
   - Một lần Lưu gửi toàn bộ thay đổi, tính lại và tạo phiên mới. File và phiên trước được giữ nguyên. Nếu có lỗi, giữ bản nháp để sửa và thử lại.
   - Sửa nguồn: công thức bắt đầu bằng `=`, thêm dấu nháy đơn trước nội dung để giữ văn bản như `'00123`, dùng dấu chấm cho số thập phân. Sửa mọi sheet nguồn cần Windows worker như đã cấu hình.
   - Lưu hoặc Hủy trước khi đổi file, tải xuống hay mở phiên khác. Tối đa 1000 ô mỗi lần lưu. Bản nháp chưa lưu không được lưu bền vững khi đóng trang.
5. **Kiểm tra lỗi & Nhật ký (Audit Trail):**
   - Bảng *"Những điểm cần kiểm tra"*: Liệt kê các ô lỗi nguồn (`#REF!`, tuần không hợp lệ theo chuẩn ISO) để bấm nhảy trực tiếp tới ô cần xử lý.
   - Bảng *"Lịch sử chỉnh sửa"*: Lưu vết chi tiết ai đã sửa ô nào, từ giá trị nào sang giá trị nào và vào thời gian nào.
6. **Xuất kết quả:**
   - Bấm nút **"Xuất kết quả ↧"** ở góc phải trên để tải về trọn bộ file `.zip` chứa toàn bộ các file Excel Family hoàn chỉnh cùng báo cáo kiểm tra.
   - Bấm biểu tượng ↧ cạnh tên file để tải riêng file Excel của Family đang chọn.

---

## 2. Hướng dẫn chạy Dòng lệnh (CLI Mode)

Nếu muốn chạy trực tiếp bằng dòng lệnh hoặc tích hợp vào hệ thống tự động:

```powershell
# Cài đặt thư viện phụ thuộc
python -m pip install -r requirements.txt

# Chạy bóc tách toàn bộ file
python process.py --input "2026 INTERNAL SCHEDULE FERRERO-WK40(LOG)-T.xlsx"

# Chạy với tham số tùy chọn nâng cao
python process.py --input "master.xlsx" --template "BARBIE_template.xlsx" --output "output_families" --week 40 --season 2728

# Chỉ xuất 1 hoặc một số Family cụ thể trong khoảng tuần mong muốn
python process.py --input "2026 INTERNAL SCHEDULE FERRERO-WK40(LOG)-T.xlsx" --family "BARBIE 2728" --start-week 2026/42 --end-week 2027/52 --output "barbie_filtered"
```

### Các cờ lệnh hữu ích:
- `--family`: Lọc tên Family cần xuất (có thể khai báo nhiều lần).
- `--start-week`, `--end-week`: Giới hạn phạm vi tuần (định dạng `YYYY/WW`).
- `--date YYYY-MM-DD`: Ngày báo cáo ghi trên header; mặc định lấy ngày hiện tại.
- `--overwrite`: Cho phép ghi đè các file Excel đã tồn tại trong thư mục đích.
- `--strict`: Bật chế độ kiểm tra nghiêm ngặt (dừng lại nếu phát hiện lỗi công thức nguồn).

---

## 3. Nghiệp vụ xử lý dữ liệu chi tiết (Business Logic)

1. **Nhận diện Family và giải ô gộp:**
   - Đọc Dòng 3 của sheet `SUM` từ Cột 22 trở đi để nhận diện ranh giới các Family.
   - Xử lý chuẩn xác các ô gộp (Merged cells), gom cả các Family có cột rời rạc không liền nhau.
   - Tự động lọc bỏ các cột phụ tính tổng trung gian như `Total` và `Nr. Pallet` để không bị trùng lặp số lượng.
2. **Quy tắc phân nhóm Thị trường (Markets):**
   - Phân nhóm dựa trên bộ ba: **Quốc gia + Điểm đến (Destination) + Phiên bản (Version)** nhằm tách biệt các thị trường đặc thù (ví dụ: `INTERNATIONAL` vs `ISRAEL`).
   - Các thông tin `PO No.` và màu `Capsule` đa dạng trong cùng một nước được gom gọn khoa học.
3. **Cấu trúc 3 Sheet đầu ra:**
   - **Sheet 1: `Breakdown ` (Chi tiết theo con hàng):**
     * Dòng 2–13: Metadata (Loại KS/KJ, Quốc gia, Mã cảng, Tên con hàng, Mã MPG, Version, confirm, Option, Total, Code, Capsule color).
     * Dòng 14 trở đi: Số lượng giao hàng theo từng tuần của từng con hàng. Cột cuối bảng là khối `TOTAL KS` tính tổng sản lượng từng con qua các nước.
   - **Sheet 2: `[Tên Family]` (Tổng hợp theo thị trường):**
     * Dòng 1–4: Thông tin nhà cung ứng (`NIIGATA`), Tên Family, Ngày báo cáo.
     * Dòng 5–14: Thống kê theo từng quốc gia (Mã cảng, Version, Capsule, Số lượng item, Tổng confirm, Tổng option, Tổng cộng, PO, Format T00).
     * Dòng 15 trở đi: Tổng số lượng theo từng tuần của mỗi nước, cột `Total (weekly)` tính tổng tuần, cột `CUM (Total)` tính lũy kế cộng dồn, cột `Date` (thứ Hai ISO), số tuần và các cột logistics (`ASN NO.`, `ETD`...).
   - **Sheet 3: `Release Qty` (Lịch ngày cut-off):**
     * Danh sách ngày thứ Hai đầu tuần theo chuẩn ISO tương ứng với các tuần có lịch giao hàng.
4. **Giữ nguyên 100% định dạng từ Template mẫu:**
   - Kế thừa toàn bộ style từ file mẫu `BARBIE 2728`: Phông chữ, cỡ chữ, màu sắc, viền ô, căn lề, độ rộng cột, độ cao dòng, cố định dòng/cột (Freeze Panes).
   - Tự động mở rộng dải định dạng mẫu nếu Family có số lượng item hoặc số lượng thị trường nhiều hơn mẫu.
5. **Tính toàn vẹn công thức & Độc lập file:**
   - Toàn bộ công thức trong file kết quả là **công thức nội bộ (Internal formulas)**, không dùng liên kết file ngoài (`External Links`) để tránh lỗi đứt gãy link (`#REF!`) khi gửi file qua máy tính khác.
   - Giá trị số lượng được giữ nguyên đơn vị gốc từ nguồn, không tự ý làm tròn hay nhân chia quy đổi.

---

## 4. Cấu trúc thư mục dự án

```text
d:\XU_LY_DATA/
├── ui/                                # Giao diện Web Client
│   ├── index.html                     # Bố cục 2 Tab (Tải file lên & Sửa/Xuất kết quả)
│   ├── app.js                         # Logic xử lý giao diện, chuyển Tab, đối chiếu & sửa ô
│   └── style.css                      # Hệ thống style giao diện Niigata Shipment Studio
├── templates/                         # Thư mục chứa các file Excel mẫu định dạng
│   ├── BARBIE 2728 _Weekly shipment schedule 2728_WK39.xlsx
│   └── NATOONS WOODLAND_Weekly shipment schedule 2728_WK39.xlsx
├── test/                              # Thư mục chứa toàn bộ mã nguồn kiểm thử tự động
│   ├── __init__.py
│   ├── test_ui.py                     # Bộ kiểm thử hồi quy cho API chỉnh sửa và audit
│   ├── test_process.py                # Bộ kiểm thử lõi bóc tách dữ liệu
│   └── test_template_format.py        # Bộ kiểm thử kế thừa định dạng template
├── ui_server.py                       # HTTP server xử lý API, nạp file, session & xuất ZIP
├── process.py                         # Thuật toán lõi bóc tách dữ liệu từ sheet SUM
├── template_format.py                 # Module sao chép và kế thừa định dạng từ Template mẫu
├── start_ui.bat                       # File nhấp đúp chạy nhanh UI trên Windows
├── requirements.txt                   # Danh sách thư viện Python (openpyxl, pandas)
├── .gitignore                         # File cấu hình bỏ qua file tạm, cache và build
└── 2026 INTERNAL SCHEDULE FERRERO-WK40(LOG)-T.xlsx # File kế hoạch nguồn
```

---

## 5. Kiểm thử hệ thống (Unit & Regression Tests)

Chạy các bài kiểm thử Python và giao diện:

```powershell
python -m unittest discover
node --test test/test_inline_editor.cjs test/test_progress.cjs
```
Mọi trường hợp về ô gộp phức tạp, công thức lũy kế, khôi phục giá trị gốc, xử lý tuần ISO và độc lập liên kết đều được kiểm thử tự động thành công (100% PASS).

Tiến độ tạo file được gửi trực tiếp qua cùng yêu cầu tải lên (`application/x-ndjson`), kể cả trên Vercel. Thanh tiến độ dành 0–80% cho tạo Family, 80–99% cho đối chiếu với nguồn và chỉ hiển thị 100% khi báo cáo đã lưu xong. Đây là tỷ lệ theo các bước hoàn thành, không phải phần trăm thời gian. Đồng hồ chạy trên trình duyệt từ lúc bấm tạo; hoàn tất sẽ đóng lớp tiến độ và mở tab kết quả. Cần deploy lại cả giao diện và `ui_server.py` để áp dụng.
