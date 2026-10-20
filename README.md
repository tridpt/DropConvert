# Thả & Đổi

Ứng dụng Windows để xử lý tài liệu ngay trên máy. Một nút nổi nằm ở mép phải màn hình; khi kéo tệp qua nút, nó mở rộng và nhận tệp. Thả tệp vào nút để mở menu thao tác phù hợp. Bạn cũng có thể nhấp nút rồi chọn tệp.

## Dùng bản đã đóng gói

Sau khi chạy `build.ps1`, mở `dist\ThaVaDoi\ThaVaDoi.exe`, hoặc giải nén `ThaVaDoi-Windows-0.4.0.zip` rồi chạy `ThaVaDoi.exe`. Giữ nguyên cả thư mục `ThaVaDoi` khi sao chép sang nơi khác. Bản đóng gói đã kèm bộ OCR tiếng Việt và tiếng Anh, không cần cài Tesseract riêng.

1. Kéo DOCX, DOC, PDF, TXT, Markdown hoặc ảnh vào nút tím ở mép phải.
2. Chọn thao tác trong menu. Nếu gộp nhiều PDF, kéo các dòng trong danh sách để sắp thứ tự.
3. Tệp kết quả nằm trong thư mục `DaChuyenDoi` cạnh tệp gốc đầu tiên. Nút **Đổi…** cho phép chọn nơi khác.
4. Bấm **Mở tệp** hoặc **Mở thư mục** khi hoàn tất. Biểu tượng ở khay hệ thống có lệnh mở lại, ẩn nút và thoát.
5. Nếu muốn mở từ File Explorer, bấm **Bật chuột phải** ở đầu menu. Nhấp phải một tệp được hỗ trợ rồi chọn **Mở bằng Thả & Đổi**. Với nhiều tệp, chọn tất cả, nhấp phải và dùng **Gửi tới → Thả & Đổi**. Bấm **Tắt chuột phải** để gỡ cả hai lệnh. Trên Windows 11, chúng có thể nằm dưới **Hiển thị thêm tùy chọn**.

Tệp gốc không bị sửa hoặc ghi đè. Tệp trùng tên được thêm số. Tất cả thao tác chạy cục bộ, không tải tài liệu lên mạng.

## Chức năng

| Nguồn | Thao tác |
| --- | --- |
| DOCX | PDF, TXT, Markdown |
| DOC | PDF, DOCX |
| PDF | Word từ phần chữ, TXT, ảnh PNG từng trang, OCR sang TXT/Word/PDF tìm kiếm được, gộp, tách trang, sắp xếp trang bằng ảnh thu nhỏ, xoay, tối ưu hoặc nén theo chất lượng, đặt mật khẩu |
| Ảnh PNG/JPG/BMP/WebP/TIFF | PDF (có thể ghép nhiều ảnh), OCR sang TXT/Word/PDF tìm kiếm được |
| TXT/Markdown | DOCX, PDF |

Chọn nhiều tệp cùng loại để xử lý hàng loạt. PDF và ảnh từ nhiều tệp có thể gộp thành một PDF. Với tệp lẫn định dạng, bấm **Mở hàng đợi từng tệp**: mỗi dòng chọn thao tác và mức nén riêng, xem trạng thái, dừng sau tệp hiện tại hoặc thử lại dòng lỗi. Có thể kéo thêm tệp vào cửa sổ hàng đợi.

**Nén PDF theo chất lượng** có ba mức: rõ nét (220 DPI), cân bằng (160 DPI), dung lượng nhỏ (110 DPI). Chữ và hình vector vẫn giữ nguyên; ảnh nhúng được giảm kích thước/chất lượng. Nếu kết quả không nhỏ hơn bản gốc, ứng dụng lưu một bản sao nguyên chất lượng.

Với một PDF, chọn **Sắp xếp trang** để xem ảnh từng trang, kéo đổi thứ tự, xoay riêng trang hoặc bỏ trang. Bấm **Lưu PDF mới**; tệp nguồn vẫn giữ nguyên.

## Điều kiện và giới hạn

- Chuyển DOCX/DOC/TXT/Markdown sang PDF và đổi DOC sang DOCX dùng **Microsoft Word đã cài trên máy**. Các thao tác đó sẽ bị vô hiệu hóa nếu không tìm thấy Word.
- Bản Windows đã kèm Tesseract OCR và dữ liệu nhận dạng Việt + Anh. Khi chạy từ mã nguồn, máy cần cài Tesseract riêng.
- **PDF tìm kiếm được** giữ nội dung và hình ảnh gốc; chỉ các trang có rất ít chữ mới được OCR và thêm lớp chữ ẩn để tìm kiếm, sao chép. Trang có chữ sẵn được giữ nguyên. OCR tự tăng tương phản và dò chiều trang; chất lượng nhận dạng tùy độ rõ của trang.
- PDF sang Word lấy phần chữ có sẵn. Với PDF scan, dùng **OCR sang Word**. Bảng, hình và bố cục phức tạp có thể không giữ nguyên.
- **Tối ưu PDF** dùng nén không giảm chất lượng. Nếu tệp vốn đã tối ưu, dung lượng có thể không giảm.
- PDF có mật khẩu phải được mở khóa trước khi xử lý.
- Nút nổi luôn nằm ở mép phải và mở rộng khi kéo tệp vào. Windows không cho ứng dụng nhận biết một tệp đang được kéo ở nơi khác trên màn hình nếu chưa kéo qua vùng của nút.
- Tích hợp File Explorer chỉ được thêm khi bạn bấm bật, dùng registry và thư mục Gửi tới của người dùng hiện tại, không cần quyền quản trị. Lệnh **Mở bằng Thả & Đổi** nhận một tệp; **Gửi tới → Thả & Đổi** nhận nhiều tệp.

## Chạy từ mã nguồn

Yêu cầu Python 3.11 trở lên trên Windows:

```powershell
python -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
.\.venv\Scripts\python.exe run.py
```

Đóng gói lại thành ứng dụng Windows:

Máy đóng gói cần cài Tesseract OCR; `build.ps1` sao chép bộ chạy và DLL của nó vào gói ứng dụng. Có thể đặt biến môi trường `TESSERACT_HOME` trỏ đến thư mục cài đặt.

```powershell
.\build.ps1
```

Bài kiểm tra:

```powershell
.\.venv\Scripts\python.exe -m pip install pytest
.\.venv\Scripts\python.exe -m pytest -q
```

## Thành phần bên thứ ba

Giao diện dùng PySide6 (LGPL). Các thư viện xử lý tài liệu được liệt kê trong `requirements.txt`. Tesseract OCR và tệp dữ liệu `eng.traineddata`, `vie.traineddata`, `osd.traineddata` được phân phối theo Apache 2.0; xem `ocr_data/LICENSE-Tesseract.txt`. Bản Windows cũng kèm giấy phép này. Nguồn của `vie.traineddata`: <https://github.com/tesseract-ocr/tessdata_fast>.
