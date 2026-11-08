#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
tsv_to_xlsx.py - Chuyển file dóng hàng Hán–Việt từ TSV sang XLSX theo đúng định dạng đồ án.

Đầu ra:
  - Sheet "DONG_HANG": hàng 1 là "Hán" | "Việt"; từ hàng 2, mỗi hàng đúng một dòng của TSV.
    Không gộp ô, không hàng trống, không công thức, không cột số thứ tự.
    Mọi ô được ghi là văn bản (kể cả khi bắt đầu bằng "=", "+", "-", "@" hay toàn chữ số).
  - Sheet "THONG_KE" (tùy chọn, mặc định có): số cặp, số dòng Hán/Việt... chỉ là giá trị tĩnh, không ảnh hưởng cấu trúc sheet DONG_HANG. Tắt bằng --no-stats.

Quy trình:
  1. (tùy chọn --normalize) sửa TSV tại chỗ: bỏ khoảng trắng đầu/cuối mỗi vế, chuẩn hóa Unicode NFC, đổi CRLF -> LF, xóa dòng trống. File gốc được sao lưu thành <tên>.tsv.bak (nhớ xóa file .bak trước khi nộp).
  2. Kiểm tra TSV bằng validate_tsv.py. Nếu có LỖI thì dừng, không tạo XLSX.
  3. Ghi XLSX.
  4. Đọc lại XLSX và so với TSV (trừ khi dùng --no-verify) để bảo đảm hai file khớp nhau.

Cách dùng:
  python scripts/tsv_to_xlsx.py CPR_01/04_dong_hang/CPR_01_parallel.tsv
  python scripts/tsv_to_xlsx.py file.tsv -o file.xlsx --normalize
  python scripts/tsv_to_xlsx.py "CPR_0*/04_dong_hang/*_parallel.tsv"

Mã thoát: 0 = thành công, 1 = TSV có lỗi hoặc XLSX không khớp, 2 = lỗi sử dụng.
"""
from __future__ import annotations

import argparse
import datetime
import glob
import shutil
import sys
import unicodedata
from pathlib import Path
from types import SimpleNamespace

sys.path.insert(0, str(Path(__file__).resolve().parent))
import validate_tsv as V  # noqa: E402

try:
    from openpyxl import Workbook
    from openpyxl.styles import Alignment, Font, PatternFill
except ImportError:
    sys.exit("Chưa cài openpyxl. Chạy: pip install openpyxl")

FONT = "Arial"


def normalize_tsv_inplace(path: Path) -> str:
    """Chuẩn hóa nhẹ TSV tại chỗ (xem docstring). Trả về mô tả những gì đã đổi."""
    raw = path.read_bytes()
    body = raw[3:] if raw.startswith(b"\xef\xbb\xbf") else raw
    try:
        text = body.decode("utf-8")
    except UnicodeDecodeError:
        return "không chuẩn hóa được vì file không phải UTF-8"
    lines = text.split("\n")
    if lines and lines[-1] == "":
        lines.pop()
    out, dropped = [], 0
    for line in lines:
        line = line.rstrip("\r")
        if not line.strip():
            dropped += 1
            continue
        if line.count("\t") == 1:
            han, viet = line.split("\t")
            line = unicodedata.normalize("NFC", han.strip()) + "\t" + unicodedata.normalize("NFC", viet.strip())
        out.append(line)
    new = ("\n".join(out) + "\n").encode("utf-8")
    if new == raw:
        return "không có gì cần đổi"
    shutil.copy2(path, str(path) + ".bak")
    path.write_bytes(new)
    return f"đã chuẩn hóa (xóa {dropped} dòng trống); bản gốc lưu tại {path.name}.bak"


def build_xlsx(pairs, out_path: Path, src_name: str, with_stats: bool = True):
    wb = Workbook()
    ws = wb.active
    ws.title = V.SHEET

    head_font = Font(name=FONT, bold=True, color="FFFFFF")
    head_fill = PatternFill(start_color="1F3864", end_color="1F3864", fill_type="solid")
    ws.cell(1, 1, V.HDR_HAN)
    ws.cell(1, 2, V.HDR_VIET)
    for c in (1, 2):
        cell = ws.cell(1, c)
        cell.font, cell.fill = head_font, head_fill
        cell.alignment = Alignment(horizontal="center", vertical="center")

    body_font = Font(name=FONT, size=11)
    wrap = Alignment(wrap_text=True, vertical="top")
    for r, (han, viet) in enumerate(pairs, start=2):
        for c, val in ((1, han), (2, viet)):
            cell = ws.cell(r, c)
            cell.value = val
            cell.data_type = "s"            # ép kiểu văn bản: không bao giờ thành công thức
            cell.number_format = "@"
            cell.font, cell.alignment = body_font, wrap
    ws.column_dimensions["A"].width = 55
    ws.column_dimensions["B"].width = 80
    ws.freeze_panes = "A2"

    if with_stats:
        st = wb.create_sheet("THONG_KE")
        n_cjk = sum(len(V.CJK_RE.findall(h)) for h, _ in pairs)
        n_syl = sum(len(v.split()) for _, v in pairs)
        n = len(pairs)
        rows = [
            ("Tệp nguồn (TSV)", src_name),
            ("Số cặp dóng hàng", n),
            ("Số dòng Hán", n),
            ("Số dòng Việt", n),
            ("Tổng số chữ Hán", n_cjk),
            ("Tổng số âm tiết Việt (tách theo dấu cách)", n_syl),
            ("Trung bình chữ Hán / cặp", round(n_cjk / n, 1) if n else 0),
            ("Trung bình âm tiết Việt / cặp", round(n_syl / n, 1) if n else 0),
            ("Ngày tạo", datetime.datetime.now().strftime("%Y-%m-%d %H:%M")),
            ("Công cụ", "scripts/tsv_to_xlsx.py"),
        ]
        st.cell(1, 1, "Chỉ tiêu").font = head_font
        st.cell(1, 2, "Giá trị").font = head_font
        for c in (1, 2):
            st.cell(1, c).fill = head_fill
        for i, (k, v) in enumerate(rows, start=2):
            st.cell(i, 1, k).font = Font(name=FONT)
            cell = st.cell(i, 2, v)
            cell.font = Font(name=FONT)
            cell.alignment = Alignment(horizontal="left")
        st.cell(len(rows) + 3, 1,
                "Giá trị tĩnh, được tính lúc xuất file. Sheet này không thuộc dữ liệu dóng hàng.").font = \
            Font(name=FONT, italic=True, color="595959")
        st.column_dimensions["A"].width = 46
        st.column_dimensions["B"].width = 34

    wb.save(out_path)


def convert_one(tsv: Path, out: Path, args) -> bool:
    print(f"\n>>> {tsv}")
    if args.normalize:
        print("Chuẩn hóa TSV:", normalize_tsv_inplace(tsv))

    opts = SimpleNamespace(ratio_low=V.DEFAULT_RATIO_LOW, ratio_high=V.DEFAULT_RATIO_HIGH)
    rep = V.Report(tsv.name)
    entries = V.read_tsv(tsv, rep, opts)
    if rep.n_err:
        V.print_report(rep, args.max_show)
        print("\nKhông tạo XLSX vì TSV còn lỗi. Sửa lỗi rồi chạy lại.")
        return False
    if rep.n_warn:
        codes = sorted(rep.warnings)
        print(f"TSV có {rep.n_warn} cảnh báo ({', '.join(codes)}). "
              "Xem chi tiết: python scripts/validate_tsv.py " + str(tsv))
        if {"W_SPACES", "W_NOT_NFC", "W_CRLF"} & set(codes) and not args.normalize:
            print("Gợi ý: chạy lại với --normalize để tự sửa khoảng trắng, Unicode NFC và CRLF.")

    pairs = [(h, v) for _, h, v in entries]
    build_xlsx(pairs, out, tsv.name, with_stats=not args.no_stats)
    print(f"Đã ghi {len(pairs)} cặp -> {out}")

    if args.no_verify:
        return True
    rep2 = V.Report(out.name)
    xrows = V.read_xlsx(out, rep2)
    if xrows is not None:
        V.compare(rep2, entries, xrows)
    if rep2.n_err:
        V.print_report(rep2, args.max_show)
        print("XLSX KHÔNG khớp TSV. Không nộp file này.")
        return False
    print("Đã đọc lại và so khớp: XLSX giống hệt TSV.")
    return True


def main(argv=None) -> int:
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8", errors="replace")
        except Exception:                   # noqa: BLE001
            pass
    ap = argparse.ArgumentParser(description="Chuyển TSV dóng hàng Hán–Việt sang XLSX (sheet DONG_HANG).")
    ap.add_argument("tsv", nargs="+", help="một hoặc nhiều file .tsv (hỗ trợ ký tự đại diện * ?)")
    ap.add_argument("-o", "--output", help="đường dẫn .xlsx đầu ra (chỉ khi có đúng một file .tsv)")
    ap.add_argument("--normalize", action="store_true",
                    help="sửa TSV tại chỗ: bỏ khoảng trắng đầu/cuối vế, NFC, CRLF->LF, xóa dòng trống")
    ap.add_argument("--no-stats", action="store_true", help="không thêm sheet THONG_KE")
    ap.add_argument("--no-verify", action="store_true", help="bỏ bước đọc lại XLSX để so khớp")
    ap.add_argument("--max-show", type=int, default=10)
    args = ap.parse_args(argv)

    files = []
    for a in args.tsv:
        files += sorted(glob.glob(a)) if any(ch in a for ch in "*?[") else [a]
    if not files:
        print("Không tìm thấy file nào khớp.", file=sys.stderr)
        return 2
    if args.output and len(files) != 1:
        print("-o chỉ dùng được với đúng một file .tsv.", file=sys.stderr)
        return 2
    for f in files:
        if not Path(f).is_file():
            print(f"Không tìm thấy file: {f}", file=sys.stderr)
            return 2

    ok = True
    for f in files:
        tsv = Path(f)
        out = Path(args.output) if args.output else tsv.with_suffix(".xlsx")
        ok &= convert_one(tsv, out, args)
    return 0 if ok else 1


if __name__ == "__main__":
    sys.exit(main())