# Trình biên tập Giáo trình Phần mềm mã nguồn mở

Website tĩnh + hàm serverless chạy trên **Vercel**, mã nguồn lưu trên **GitHub**. Cho phép xem, chỉnh sửa trực tiếp nội dung 6 chương, thêm đoạn văn, tiêu đề, danh sách, hộp nhấn mạnh, bảng, khối mã, hình ảnh và **xuất ra file Word** theo mục “Định dạng Word” của `rules.md`.

## Tính năng

- 6 chương đã đưa sẵn lên website (thư mục `content/`, ảnh trong `assets/img/`).
- Sửa trực tiếp trên trang: bấm vào đoạn văn, tiêu đề, ô bảng, chú thích hình; Enter trong đoạn để thêm đoạn mới.
- Thêm nội dung: đoạn văn, tiêu đề (1.1 / 1.1.1 / a), danh sách, 5 loại hộp nhấn mạnh (Ghi nhớ, Lưu ý, Ví dụ, Tình huống, Mở rộng), bảng (thêm/bớt hàng, cột), khối mã, trích dẫn mẫu, **hình ảnh tải từ máy**.
- Sắp xếp khối (▲ ▼), nhân đôi, xóa.
- Định dạng chữ đậm, nghiêng, mã; **trích dẫn số `[n]` tự đánh số theo thứ tự xuất hiện** trong chương; bảng quản lý tài liệu tham khảo.
- Hình, bảng tự đánh số theo chương (Hình 2.3, Bảng 3.1).
- **Xuất Word** một chương hoặc cả giáo trình (A4, lề trái 3 cm, các lề còn lại 2 cm, Times New Roman 13, giãn dòng 1,3, đầu trang “Chương n”, số trang ở chân trang, bảng/hình/hộp nhấn mạnh theo quy định). Việc xuất chạy hoàn toàn trên trình duyệt.
- Tự lưu trên trình duyệt (IndexedDB), sao lưu/nhập JSON, khôi phục chương về bản gốc.
- Lưu lâu dài lên GitHub bằng nút **Khác → Lưu lên GitHub** (hàm serverless `api/save.js`).

## Cấu trúc thư mục

```
index.html            giao diện
css/style.css         kiểu dáng
js/markup.js          định dạng trong dòng, trích dẫn [@khóa], đánh số hình/bảng
js/export.js          xuất Word (thư viện docx) theo rules.md
js/store.js           lưu cục bộ (IndexedDB)
js/app.js             trình soạn thảo
vendor/docx.min.js    thư viện docx 8.5.0 (đã kèm sẵn, không cần cài)
api/save.js           hàm serverless ghi nội dung/ảnh vào GitHub
content/chuong-1..6.json   nội dung 6 chương
assets/img/           hình của các chương (và assets/img/uploads/ cho ảnh tải lên)
test/export.test.js   kiểm thử xuất Word bằng Node
tools/export_content.py    tạo lại content/*.json từ tập lệnh dựng chương (dùng trên máy tác giả)
vercel.json, package.json
```

## Chạy thử trên máy

```bash
npx serve .          # rồi mở http://localhost:3000
```
(Không mở trực tiếp bằng `file://` vì trang cần nạp `content/*.json`.) Hàm `/api/save` chỉ hoạt động khi đã triển khai trên Vercel hoặc chạy `vercel dev`.

## Triển khai: GitHub → Vercel

1. **Tạo kho GitHub** (riêng tư nếu không muốn công khai nội dung) rồi đẩy mã nguồn:
   ```bash
   cd website
   git init && git add . && git commit -m "Trình biên tập giáo trình PMNM"
   git branch -M main
   git remote add origin https://github.com/<tên-người-dùng>/<tên-kho>.git
   git push -u origin main
   ```
2. **Vercel** → *Add New… → Project* → chọn kho vừa tạo → *Framework Preset: Other* (không cần lệnh build, thư mục gốc là `./`) → *Deploy*. Từ đó mỗi lần có commit mới trên `main`, Vercel tự triển khai lại.
3. **Bật lưu lên GitHub** (tùy chọn): tạo *fine-grained personal access token* tại GitHub → *Settings → Developer settings*, chỉ chọn kho này, quyền **Contents: Read and write**. Trong Vercel → *Settings → Environment Variables* thêm:

   | Biến | Giá trị |
   |---|---|
   | `GITHUB_TOKEN` | token vừa tạo |
   | `GITHUB_REPO` | `tên-người-dùng/tên-kho` |
   | `GITHUB_BRANCH` | `main` (mặc định) |
   | `EDIT_PASSWORD` | mật khẩu do bạn đặt, cung cấp cho người biên tập |

   Sau đó *Redeploy*. Trong website: **Khác → Lưu lên GitHub**, nhập `EDIT_PASSWORD`, chọn “Lưu chương” hoặc “Lưu mọi chương đã sửa”.

## Lưu ý về dữ liệu

- Chỉnh sửa được lưu **trên trình duyệt đang dùng** cho tới khi bấm “Lưu lên GitHub”. Dùng “Tải bản sao lưu” để lưu một bản JSON (kèm ảnh) về máy.
- Khi nhiều người cùng biên tập, mỗi người có bản riêng trên trình duyệt của mình; người bấm “Lưu lên GitHub” sau cùng sẽ ghi đè nội dung chương đó. Nên thống nhất ai phụ trách chương nào.
- Ảnh tải lên được thu nhỏ còn tối đa 1400 px, nén JPEG; ảnh lớn hơn khoảng 2,4 MB sau nén có thể bị hàm serverless từ chối (giới hạn 4,5 MB mỗi yêu cầu của Vercel).
- Số trích dẫn `[n]` chỉ là hiển thị; trong dữ liệu mỗi trích dẫn dùng một khóa (`[@khóa]`), nên thêm/bớt nguồn không làm sai số.

## Cập nhật nội dung gốc

Nếu chỉnh sửa các chương bằng tập lệnh trong `BIEN SOAN MOI/cong-cu` thì chạy lại `tools/export_content.py` để sinh lại `content/*.json` và ảnh (sẽ ghi đè các chỉnh sửa đã lưu trên GitHub).

## Kiểm thử

```bash
npm install          # chỉ cần để chạy kiểm thử (docx@8.5.0)
npm test             # xuất 6 chương ra file .docx tạm và báo kết quả
```
