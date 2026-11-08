## 1. Cài đặt và chạy

```bash
python -m venv .venv
source .venv/bin/activate        # Windows: .venv\Scripts\activate
pip install -r requirements.txt
```

Yêu cầu: Python 3.8 trở lên.

## 2. Git

**Nhánh**: `<loại>-<mô-tả-ngắn>` (chữ thường, dấu gạch ngang).

| Loại       | Dùng cho                        | Ví dụ                          |
| ---------- | ------------------------------- | ------------------------------ |
| `feat`     | tính năng mới                   | `feat/b-dpll-unit-propagation` |
| `fix`      | sửa lỗi                         | `fix/c-exactly-k0-encoding`    |
| `test`     | thêm test                       | `test/c-evaluator-vs-cnf`      |
| `docs`     | báo cáo, tài liệu               | `docs/e-report-skeleton`       |
| `refactor` | đổi cấu trúc, không đổi hành vi | `refactor/a-trace-format`      |

Nhánh chính: `main` (luôn chạy được, không push trực tiếp). Nhánh tích hợp tuỳ chọn: `dev`.

**Commit message**: `<loại>(<phạm vi>): <mô tả ngắn, thể mệnh lệnh, tiếng Anh>`
