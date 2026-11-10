# Pipeline trích xuất và làm sạch văn bản

Dự án gồm các script Python để trích xuất văn bản từ PDF Tam Quốc Diễn Nghĩa, làm sạch nội dung theo từng đoạn, tách câu và tùy chọn chuyển một file CSV sang TSV/XLSX.

> `extract_metadata.py` dùng `pdfplumber` để đọc lớp văn bản có sẵn trong PDF; script không thực hiện OCR ảnh. Nếu PDF chỉ gồm ảnh quét, cần OCR PDF trước khi chạy pipeline.

## 1. Cấu trúc thư mục

```text
.
├── CV05_CPR_03/
│   ├── 01_nguon/                 # PDF nguồn
│   ├── 02_ocr_tho/               # Văn bản trích xuất thô
│   │   ├── han/
│   │   ├── han_viet/
│   │   └── viet/
│   ├── 03_text_sach/             # Văn bản đã làm sạch
│   │   └── sentence/             # Văn bản đã tách câu
│   └── 04_dong_hang/             # TSV/XLSX tùy chọn
├── extract_metadata.py
├── extract_text_v3.py            # Entrypoint thay thế cho extract_metadata.py
├── clean_data.py
├── turn_csv_xlxs_tsv.py
└── requirements.txt
```

Các thư mục kết quả được tạo tự động khi chạy script. `.gitignore` bỏ qua môi trường ảo, bytecode và các thư mục kết quả này; các file vẫn tồn tại trên máy để sử dụng.

## 2. Tạo môi trường Python

Chạy các lệnh sau từ thư mục gốc của repository:

```bash
python3 -m venv .venv
source .venv/bin/activate
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
```

`requirements.txt` cài các thư viện `pdfplumber`, `pandas` và `openpyxl`. Khi mở terminal mới, kích hoạt lại môi trường bằng:

```bash
source .venv/bin/activate
```

Hoặc không kích hoạt môi trường và gọi trực tiếp `.venv/bin/python` trong các lệnh bên dưới.

## 3. Đặt PDF đầu vào

Đặt các PDF cần xử lý trong:

```text
CV05_CPR_03/01_nguon/
```

Hiện script mặc định trỏ tới file tập 3:

```text
Tam Quoc Dien Nghia (Han - Viet) - quyen 3.pdf
```

Đường dẫn này được khai báo trong `extract_metadata.py` tại biến `PDF_PATH`. Để trích xuất tập 1 hoặc tập 2, sửa tên PDF trong biến này cho đúng file cần chạy. Script hiện xử lý một PDF mỗi lần và chưa nhận đường dẫn PDF qua tham số dòng lệnh.

## 4. Trích xuất văn bản thô

Từ thư mục gốc, chạy:

```bash
python extract_metadata.py
```

Hoặc dùng entrypoint tương đương:

```bash
python extract_text_v3.py
```

Chỉ chạy một trong hai lệnh trên cho mỗi PDF; `extract_text_v3.py` gọi lại `extract_metadata.main()`, không phải một bước xử lý riêng.

Script đọc PDF đã cấu hình và ghi các cột tương ứng vào:

```text
CV05_CPR_03/02_ocr_tho/han/<tên-PDF-không-phần-mở-rộng>.txt
CV05_CPR_03/02_ocr_tho/han_viet/<tên-PDF-không-phần-mở-rộng>.txt
CV05_CPR_03/02_ocr_tho/viet/<tên-PDF-không-phần-mở-rộng>.txt
CV05_CPR_03/02_ocr_tho/metadata.jsonl
```

Trong đó:

- `han`: cột chữ Hán.
- `han_viet`: cột phiên âm Hán-Việt.
- `viet`: cột dịch nghĩa tiếng Việt.
- `metadata.jsonl`: thông tin đoạn, trang, cột và tọa độ.

### Xử lý nhiều tập

Lặp lại bước trích xuất cho từng tập:

1. Sửa `PDF_PATH` trong `extract_metadata.py` thành đường dẫn PDF tập cần xử lý.
2. Chạy `python extract_metadata.py`.
3. Lặp lại với các tập còn lại.
4. Sau khi đã trích xuất tất cả PDF cần dùng, chạy bước làm sạch ở mục tiếp theo.

Tên file OCR dựa trên tên PDF nên các tập được lưu thành các file riêng. `metadata.jsonl` hiện có tên cố định và sẽ bị ghi lại mỗi lần trích xuất; file này chỉ giữ metadata của PDF chạy gần nhất.

## 5. Làm sạch và tách câu

Sau khi đã tạo văn bản thô cho các tập cần xử lý, chạy:

```bash
python clean_data.py
```

Script đọc tất cả file `.txt` trong ba thư mục `han/`, `han_viet/` và `viet/` của `02_ocr_tho`. Kết quả được lưu với tên gồm tên PDF nguồn và loại văn bản, ví dụ:

```text
CV05_CPR_03/03_text_sach/
├── Tam Quoc Dien Nghia (Han - Viet) - quyen 3_han.txt
├── Tam Quoc Dien Nghia (Han - Viet) - quyen 3_han_viet.txt
├── Tam Quoc Dien Nghia (Han - Viet) - quyen 3_viet.txt
└── sentence/
    ├── Tam Quoc Dien Nghia (Han - Viet) - quyen 3_han.txt
    ├── Tam Quoc Dien Nghia (Han - Viet) - quyen 3_han_viet.txt
    └── Tam Quoc Dien Nghia (Han - Viet) - quyen 3_viet.txt
```

Với tập 1 và tập 2, phần tên tập trong tên file đầu ra sẽ lấy theo tên PDF tương ứng. Bước làm sạch lọc một số header, số trang, ký hiệu tham chiếu và chú thích; sau đó xuất bản sạch và bản tách câu. Script không xóa các file đầu vào hay kết quả cũ.

## 6. Chuyển CSV sang TSV/XLSX (tùy chọn)

`turn_csv_xlxs_tsv.py` là bước riêng, không được gọi bởi bước trích xuất hoặc làm sạch. Script đọc CSV có đường dẫn khai báo ở cuối file, mặc định:

```text
Demo_CV05_CPR_03(Sheet1).csv
```

Đặt CSV này ở thư mục gốc hoặc sửa `csv_file_path` trong `turn_csv_xlxs_tsv.py`, sau đó chạy:

```bash
python turn_csv_xlxs_tsv.py
```

Hai file được tạo trong:

```text
CV05_CPR_03/04_dong_hang/CPR_03_parallel.tsv
CV05_CPR_03/04_dong_hang/CPR_03_parallel.xlsx
```

Script này chỉ chuyển đổi CSV có sẵn; nó không tự ghép các file trong `03_text_sach` thành dữ liệu song song.

## 7. Kiểm tra cú pháp

Có thể kiểm tra các script mà không chạy pipeline:

```bash
python -m py_compile extract_metadata.py extract_text_v3.py clean_data.py turn_csv_xlxs_tsv.py
```

## 8. Lỗi thường gặp

- **Không tìm thấy PDF:** kiểm tra tên file trong `PDF_PATH` và vị trí `CV05_CPR_03/01_nguon/`.
- **Không tìm thấy file `.txt` để làm sạch:** chạy trích xuất trước, xác nhận các file đã được tạo trong ba thư mục ngôn ngữ của `02_ocr_tho`.
- **CSV không tồn tại:** đặt CSV tại đường dẫn cấu hình trong `turn_csv_xlxs_tsv.py` hoặc cập nhật biến `csv_file_path`.
- **Không trích xuất được nội dung từ PDF:** xác nhận PDF có lớp văn bản. PDF chỉ chứa ảnh cần được OCR trước.
