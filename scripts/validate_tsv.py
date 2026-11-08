#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
validate_tsv.py - Kiểm tra file dóng hàng Hán–Việt (.tsv) và file Excel (.xlsx) tương ứng.

Kiểm tra theo đúng yêu cầu định dạng của đồ án:

  TSV   : mỗi dòng đúng một cặp  <Hán><TAB><Việt>; không dòng trống; không tab/xuống dòng bên trong một vế; không vế rỗng; UTF-8; không có dòng tiêu đề.
  XLSX  : sheet DONG_HANG, đúng 2 cột, hàng 1 là "Hán" / "Việt"; không gộp ô; không dòng trống; không công thức; không cột số thứ tự; mọi ô là văn bản.
  KHỚP  : số hàng XLSX = số dòng TSV và từng ô giống hệt từng vế trong TSV.

Mức độ:
  LỖI (E_*, X_*)   : phải sửa, script trả về mã thoát 1.
  CẢNH BÁO (W_*)   : nên xem lại (khoảng trắng thừa, Unicode không chuẩn NFC, chú thích còn sót, tỷ lệ độ dài bất thường, dòng trùng...). Không làm script thất bại, trừ khi dùng --strict.

Cách dùng:
  python scripts/validate_tsv.py CPR_01/04_dong_hang/CPR_01_parallel.tsv
  python scripts/validate_tsv.py file.tsv --xlsx file.xlsx --report bao_cao_kiem_tra.md
  python scripts/validate_tsv.py "CPR_0*/04_dong_hang/*_parallel.tsv" --report kiem_tra.md
  python scripts/validate_tsv.py file.tsv --tsv-only        # chưa có file xlsx

File .xlsx mặc định là file cùng tên, đuôi .xlsx, cùng thư mục với .tsv.
Mã thoát: 0 = đạt, 1 = có lỗi, 2 = lỗi sử dụng (không tìm thấy file...).
"""
from __future__ import annotations

import argparse
import datetime
import glob
import hashlib
import re
import sys
import unicodedata
from collections import defaultdict
from pathlib import Path
from types import SimpleNamespace

SHEET = "DONG_HANG"
HDR_HAN, HDR_VIET = "Hán", "Việt"
EXCEL_MAX_CELL = 32767            # giới hạn ký tự của một ô Excel
DEFAULT_RATIO_LOW = 0.3           # âm tiết Việt / chữ Hán: dưới ngưỡng này bị cảnh báo
DEFAULT_RATIO_HIGH = 8.0          # ... trên ngưỡng này bị cảnh báo
MIN_HAN_FOR_RATIO = 5             # chỉ xét tỷ lệ khi vế Hán có ít nhất ngần này chữ Hán

CJK_RE = re.compile(r"[\u3400-\u4dbf\u4e00-\u9fff\uf900-\ufaff\U00020000-\U0003134f]")
CTRL_RE = re.compile(r"[\x00-\x08\x0b-\x1f\x7f\ufffe\uffff]")
LINEBREAK_RE = re.compile(r"[\u2028\u2029\x85]")
INVISIBLE_RE = re.compile(r"[\u00ad\u200b-\u200d\u2060\ufeff]")
FOOTNOTE_RE = re.compile(r"\[\d+\]|\(\d+\)|［\d+］|（\d+）|〔\d+〕")

DESC = {
    # ---- lỗi TSV ----
    "E_EMPTY_FILE": "File rỗng",
    "E_ENCODING": "File không phải UTF-8",
    "E_BOM": "File có BOM ở đầu (dính vào vế Hán của dòng 1)",
    "E_BLANK_LINE": "Dòng trống",
    "E_NO_TAB": "Dòng không có ký tự tab phân cách",
    "E_MANY_TAB": "Dòng có nhiều hơn một tab (tab nằm bên trong một vế)",
    "E_EMPTY_HAN": "Vế Hán rỗng (loại lỗi MISSING)",
    "E_EMPTY_VIET": "Vế Việt rỗng (loại lỗi MISSING)",
    "E_HEADER_IN_TSV": "Dòng tiêu đề Hán/Việt lọt vào TSV",
    "E_CTRL_CHAR": "Ký tự điều khiển trong vế (kể cả \\r lạc, \\x0b, \\x0c)",
    "E_LINEBREAK_CHAR": "Ký tự xuống dòng lạ trong vế (U+2028, U+2029, U+0085)",
    "E_TOO_LONG": "Vế dài hơn 32.767 ký tự (vượt giới hạn một ô Excel)",
    "E_HAN_NO_CJK": "Vế Hán không chứa chữ Hán nào (nghi dòng tiêu đề/số trang/rác, hoặc đảo vế)",
    "E_VIET_NO_LETTER": "Vế Việt không chứa chữ cái nào (nghi số trang/rác OCR)",
    # ---- lỗi XLSX ----
    "X_MISSING": "Không tìm thấy file .xlsx tương ứng",
    "X_OPEN": "Không mở được file .xlsx",
    "X_NO_SHEET": "Không có sheet tên DONG_HANG",
    "X_HEADER": "Hàng 1 không phải tiêu đề 'Hán' / 'Việt'",
    "X_EXTRA_COL": "Có dữ liệu ngoài hai cột A, B của sheet DONG_HANG",
    "X_MERGED": "Có ô bị gộp",
    "X_BLANK_ROW": "Có hàng trống xen giữa dữ liệu",
    "X_EMPTY_CELL": "Hàng chỉ có một vế (ô còn lại rỗng)",
    "X_FORMULA": "Ô chứa công thức",
    "X_NOT_TEXT": "Ô không phải văn bản (số/ngày/logic)",
    "X_NEWLINE": "Ô chứa ký tự xuống dòng",
    "X_TAB": "Ô chứa ký tự tab",
    "X_ROWCOUNT": "Số hàng XLSX khác số dòng TSV",
    "X_MISMATCH": "Nội dung ô XLSX khác nội dung dòng TSV",
    # ---- cảnh báo ----
    "W_CRLF": "Dòng kết thúc kiểu Windows (CRLF); nên dùng LF",
    "W_SPACES": "Khoảng trắng thừa ở đầu/cuối vế",
    "W_DOUBLE_SPACE": "Hai dấu cách liên tiếp trong vế",
    "W_NOT_NFC": "Unicode chưa chuẩn NFC (dấu tiếng Việt có thể gõ ở dạng tổ hợp)",
    "W_INVISIBLE": "Ký tự vô hình (zero-width, soft hyphen, BOM giữa dòng)",
    "W_FOOTNOTE": "Có thể còn sót chỉ mục chú thích như [1], (1)",
    "W_CJK_IN_VIET": "Vế Việt có chứa chữ Hán",
    "W_RATIO": "Tỷ lệ độ dài Hán/Việt bất thường (có thể dóng sai hoặc thiếu/thừa nội dung)",
    "W_DUPLICATE": "Cặp trùng lặp hoàn toàn (có thể hợp lệ với câu lặp như 曰：)",
    "W_FILENAME": "Tên file không theo mẫu <MA_TAC_PHAM>_parallel.tsv",
}


# --------------------------------------------------------------------------- #
#  Báo cáo
# --------------------------------------------------------------------------- #
class Report:
    def __init__(self, label: str):
        self.label = label
        self.errors = defaultdict(list)     # code -> [(where, msg)]
        self.warnings = defaultdict(list)
        self.info = {}

    def err(self, code, where, msg=""):
        self.errors[code].append((where, msg))

    def warn(self, code, where, msg=""):
        self.warnings[code].append((where, msg))

    @property
    def n_err(self):
        return sum(len(v) for v in self.errors.values())

    @property
    def n_warn(self):
        return sum(len(v) for v in self.warnings.values())


def snip(s: str, n: int = 70) -> str:
    s = s.replace("\t", "⇥").replace("\r", "␍").replace("\n", "⏎")
    return s if len(s) <= n else s[:n] + "…"


def nfc(s: str) -> str:
    return unicodedata.normalize("NFC", s)


def sha256(path: Path) -> str:
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(1 << 20), b""):
            h.update(chunk)
    return h.hexdigest()


# --------------------------------------------------------------------------- #
#  Đọc và kiểm tra TSV
# --------------------------------------------------------------------------- #
def check_cell(rep: Report, lineno: int, side: str, s: str):
    """Kiểm tra một vế. Trả về True nếu vế không rỗng."""
    where = f"dòng {lineno}"
    if not s.strip():
        rep.err("E_EMPTY_HAN" if side == "Hán" else "E_EMPTY_VIET", where, "vế rỗng")
        return False
    if CTRL_RE.search(s):
        rep.err("E_CTRL_CHAR", where, f"vế {side}: {snip(s)}")
    if LINEBREAK_RE.search(s):
        rep.err("E_LINEBREAK_CHAR", where, f"vế {side}: {snip(s)}")
    if len(s) > EXCEL_MAX_CELL:
        rep.err("E_TOO_LONG", where, f"vế {side} dài {len(s)} ký tự")
    if s != s.strip():
        rep.warn("W_SPACES", where, f"vế {side}: '{snip(s, 40)}'")
    if "  " in s:
        rep.warn("W_DOUBLE_SPACE", where, f"vế {side}: '{snip(s, 50)}'")
    if nfc(s) != s:
        rep.warn("W_NOT_NFC", where, f"vế {side}: {snip(s, 50)}")
    if INVISIBLE_RE.search(s):
        rep.warn("W_INVISIBLE", where, f"vế {side}: {snip(s, 50)}")
    if FOOTNOTE_RE.search(s):
        m = FOOTNOTE_RE.search(s)
        rep.warn("W_FOOTNOTE", where, f"vế {side}: '{m.group(0)}' trong {snip(s, 50)}")
    return True


def read_tsv(path, rep: Report, opts=None):
    """
    Đọc TSV, ghi lỗi vào rep. Trả về danh sách (số_dòng, hán, việt) cho MỌI dòng;
    với dòng sai cấu trúc (trống, thiếu tab, thừa tab) hán = việt = None.
    """
    opts = opts or SimpleNamespace(ratio_low=DEFAULT_RATIO_LOW, ratio_high=DEFAULT_RATIO_HIGH)
    path = Path(path)
    raw = path.read_bytes()
    rep.info["sha256_tsv"] = sha256(path)
    if not raw.strip():
        rep.err("E_EMPTY_FILE", "-", "file không có nội dung")
        return []
    if raw.startswith(b"\xef\xbb\xbf"):
        rep.err("E_BOM", "dòng 1", "xóa BOM (lưu lại dạng UTF-8 không BOM)")
        raw = raw[3:]
    try:
        text = raw.decode("utf-8")
    except UnicodeDecodeError as e:
        rep.err("E_ENCODING", f"byte {e.start}", "không giải mã được UTF-8; lưu lại file bằng UTF-8")
        return []

    # Chỉ tách theo '\n' (không dùng splitlines) để U+2028/U+0085... không bị coi là xuống dòng.
    lines = text.split("\n")
    if lines and lines[-1] == "":
        lines.pop()

    entries, crlf = [], 0
    seen = {}
    dup_extra = defaultdict(list)
    n_pairs = n_han = n_viet = 0

    for i, line in enumerate(lines, 1):
        if line.endswith("\r"):
            line = line[:-1]
            crlf += 1
        where = f"dòng {i}"
        if not line.strip():
            rep.err("E_BLANK_LINE", where, "dòng trống")
            entries.append((i, None, None))
            continue
        ntab = line.count("\t")
        if ntab == 0:
            hint = ""
            if re.search(r"\S {2,}\S", line):
                hint = " (có nhiều dấu cách liên tiếp; có thể đã dùng dấu cách thay cho tab)"
            rep.err("E_NO_TAB", where, snip(line) + hint)
            entries.append((i, None, None))
            continue
        if ntab > 1:
            rep.err("E_MANY_TAB", where, f"{ntab} tab: {snip(line)}")
            entries.append((i, None, None))
            continue

        han, viet = line.split("\t")
        entries.append((i, han, viet))

        if nfc(han.strip()) == HDR_HAN and nfc(viet.strip()) == HDR_VIET:
            rep.err("E_HEADER_IN_TSV", where, "TSV không có dòng tiêu đề; chỉ chứa cặp dóng hàng")
            continue

        ok_h = check_cell(rep, i, "Hán", han)
        ok_v = check_cell(rep, i, "Việt", viet)
        n_han += ok_h
        n_viet += ok_v
        if not (ok_h and ok_v):
            continue
        n_pairs += 1

        n_cjk = len(CJK_RE.findall(han))
        if n_cjk == 0:
            rep.err("E_HAN_NO_CJK", where, snip(han))
        if not any(ch.isalpha() for ch in viet):
            rep.err("E_VIET_NO_LETTER", where, snip(viet))
        if CJK_RE.search(viet):
            rep.warn("W_CJK_IN_VIET", where, snip(viet))
        if n_cjk >= MIN_HAN_FOR_RATIO:
            ratio = len(viet.split()) / n_cjk
            if ratio < opts.ratio_low or ratio > opts.ratio_high:
                rep.warn("W_RATIO", where,
                         f"{n_cjk} chữ Hán ↔ {len(viet.split())} âm tiết Việt (tỷ lệ {ratio:.1f})")

        key = (han, viet)
        if key in seen:
            dup_extra[key].append(i)
        else:
            seen[key] = i

    if crlf:
        rep.warn("W_CRLF", "-", f"{crlf} dòng kết thúc bằng CRLF")
    for key, extra in dup_extra.items():
        rep.warn("W_DUPLICATE", f"dòng {seen[key]} và {', '.join(map(str, extra))}", snip(key[0], 40))
    if not path.name.endswith("_parallel.tsv"):
        rep.warn("W_FILENAME", "-", f"'{path.name}'; tên khuyến nghị: <MA_TAC_PHAM>_parallel.tsv")

    rep.info.update(total_lines=len(lines), pairs=n_pairs, han_lines=n_han, viet_lines=n_viet)
    return entries


# --------------------------------------------------------------------------- #
#  Đọc và kiểm tra XLSX
# --------------------------------------------------------------------------- #
def read_xlsx(path, rep: Report):
    """Kiểm tra cấu trúc sheet DONG_HANG. Trả về [(số_hàng, hán, việt)] hoặc None nếu không đọc được."""
    try:
        from openpyxl import load_workbook
    except ImportError:
        rep.err("X_OPEN", "-", "chưa cài openpyxl: pip install openpyxl")
        return None
    path = Path(path)
    rep.info["sha256_xlsx"] = sha256(path)
    try:
        wb = load_workbook(path)            # không dùng data_only để phát hiện công thức
    except Exception as e:                  # noqa: BLE001
        rep.err("X_OPEN", "-", f"{type(e).__name__}: {e}")
        return None
    rep.info["sheets"] = list(wb.sheetnames)
    if SHEET not in wb.sheetnames:
        rep.err("X_NO_SHEET", "-", f"các sheet hiện có: {', '.join(wb.sheetnames)}")
        return None
    ws = wb[SHEET]

    # tiêu đề
    def hv(c):
        v = ws.cell(1, c).value
        return nfc(str(v).strip()) if v is not None else ""
    if hv(1) != HDR_HAN or hv(2) != HDR_VIET:
        rep.err("X_HEADER", "hàng 1", f"đang là '{hv(1)}' | '{hv(2)}'; cần 'Hán' | 'Việt'")

    # gộp ô
    for rng in ws.merged_cells.ranges:
        rep.err("X_MERGED", str(rng), "bỏ gộp ô")

    # dữ liệu ngoài cột A, B
    if ws.max_column > 2:
        for row in ws.iter_rows(min_row=1, min_col=3):
            for c in row:
                if c.value is not None and str(c.value).strip() != "":
                    rep.err("X_EXTRA_COL", c.coordinate, snip(str(c.value), 40))

    def empty(c):
        v = c.value
        return v is None or (isinstance(v, str) and v.strip() == "")

    rows, blank_rows, last_nonempty = [], [], 1
    if ws.max_row >= 2:
        for r, pair in enumerate(ws.iter_rows(min_row=2, max_col=2), start=2):
            a, b = pair[0], pair[1]
            ea, eb = empty(a), empty(b)
            if ea and eb:
                blank_rows.append(r)
                rows.append((r, "", ""))
                continue
            last_nonempty = r
            if ea != eb:
                rep.err("X_EMPTY_CELL", f"hàng {r}", "thiếu vế Hán" if ea else "thiếu vế Việt")
            texts = []
            for c in (a, b):
                v = c.value
                if empty(c):
                    texts.append("")
                    continue
                if c.data_type == "f":
                    rep.err("X_FORMULA", c.coordinate, snip(str(v), 40))
                    texts.append(str(v))
                    continue
                if not isinstance(v, str):
                    rep.err("X_NOT_TEXT", c.coordinate, f"{type(v).__name__}: {v!r}")
                    texts.append(str(v))
                    continue
                if "\n" in v or "\r" in v:
                    rep.err("X_NEWLINE", c.coordinate, snip(v))
                if "\t" in v:
                    rep.err("X_TAB", c.coordinate, snip(v))
                texts.append(v)
            rows.append((r, texts[0], texts[1]))
    rows = [x for x in rows if x[0] <= last_nonempty]
    for r in blank_rows:
        if r <= last_nonempty:
            rep.err("X_BLANK_ROW", f"hàng {r}", "xóa hàng trống")
    rep.info["xlsx_rows"] = len(rows)
    return rows


# --------------------------------------------------------------------------- #
#  So khớp TSV ↔ XLSX
# --------------------------------------------------------------------------- #
def first_diff(a: str, b: str) -> str:
    k = 0
    while k < min(len(a), len(b)) and a[k] == b[k]:
        k += 1
    lo = max(0, k - 12)
    return f"khác từ vị trí {k + 1}: TSV '…{snip(a[lo:k + 18], 40)}' ↔ XLSX '…{snip(b[lo:k + 18], 40)}'"


def compare(rep: Report, tsv_entries, xrows):
    n_t, n_x = len(tsv_entries), len(xrows)
    if n_t != n_x:
        rep.err("X_ROWCOUNT", "-", f"TSV {n_t} dòng, XLSX {n_x} hàng dữ liệu "
                                   "(các ô phía sau điểm lệch có thể bị báo khác nhau)")
    for k in range(min(n_t, n_x)):
        i, han, viet = tsv_entries[k]
        r, xh, xv = xrows[k]
        if han is None:                     # dòng TSV sai cấu trúc, đã báo lỗi riêng
            continue
        for side, t, x in (("Hán", han, xh), ("Việt", viet, xv)):
            if t != x:
                note = " (chỉ khác khoảng trắng/Unicode)" if nfc(t.strip()) == nfc(x.strip()) else ""
                rep.err("X_MISMATCH", f"TSV dòng {i} ↔ XLSX hàng {r}, vế {side}", first_diff(t, x) + note)


# --------------------------------------------------------------------------- #
#  In và ghi báo cáo
# --------------------------------------------------------------------------- #
def print_report(rep: Report, max_show: int = 10, out=None):
    out = out or sys.stdout
    p = lambda s="": print(s, file=out)
    p(f"\n=== {rep.label} ===")
    i = rep.info
    if "pairs" in i:
        p(f"Số dòng TSV: {i['total_lines']} | cặp hợp lệ: {i['pairs']} | "
          f"số dòng Hán: {i['han_lines']} | số dòng Việt: {i['viet_lines']}")
    if "xlsx_rows" in i:
        p(f"Số hàng dữ liệu XLSX: {i['xlsx_rows']} | các sheet: {', '.join(i.get('sheets', []))}")
    for title, bag in (("LỖI", rep.errors), ("CẢNH BÁO", rep.warnings)):
        for code in sorted(bag):
            items = bag[code]
            p(f"[{title}] {code} ({len(items)}): {DESC.get(code, '')}")
            for where, msg in items[:max_show]:
                p(f"    {where}: {msg}" if msg else f"    {where}")
            if len(items) > max_show:
                p(f"    ... và {len(items) - max_show} mục nữa (tăng --max-show để xem thêm)")
    verdict = "ĐẠT" if rep.n_err == 0 else "KHÔNG ĐẠT"
    p(f"=> {verdict}: {rep.n_err} lỗi, {rep.n_warn} cảnh báo")


def write_markdown(path, results, max_show=20):
    now = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")
    L = ["# Kết quả tự kiểm tra định dạng dóng hàng", "", f"Thời điểm kiểm tra: {now}",
         "Công cụ: scripts/validate_tsv.py", ""]
    for rep, tsv, xlsx in results:
        i = rep.info
        L += [f"## {Path(tsv).name}", ""]
        L += ["| Mục | Giá trị |", "|---|---|",
              f"| Kết quả | **{'ĐẠT' if rep.n_err == 0 else 'KHÔNG ĐẠT'}** ({rep.n_err} lỗi, {rep.n_warn} cảnh báo) |",
              f"| Số dòng TSV | {i.get('total_lines', '-')} |",
              f"| Số cặp hợp lệ | {i.get('pairs', '-')} |",
              f"| Số dòng Hán | {i.get('han_lines', '-')} |",
              f"| Số dòng Việt | {i.get('viet_lines', '-')} |",
              f"| Số hàng XLSX (sheet DONG_HANG) | {i.get('xlsx_rows', '-')} |",
              f"| SHA-256 TSV | `{i.get('sha256_tsv', '-')}` |",
              f"| SHA-256 XLSX | `{i.get('sha256_xlsx', '-')}` |", ""]
        for title, bag in (("Lỗi", rep.errors), ("Cảnh báo", rep.warnings)):
            if not bag:
                continue
            L += [f"### {title}", ""]
            for code in sorted(bag):
                items = bag[code]
                L.append(f"- **{code}** ({len(items)}): {DESC.get(code, '')}")
                for where, msg in items[:max_show]:
                    L.append(f"  - {where}: {msg}" if msg else f"  - {where}")
                if len(items) > max_show:
                    L.append(f"  - ... và {len(items) - max_show} mục nữa")
            L.append("")
    Path(path).write_text("\n".join(L), encoding="utf-8")


# --------------------------------------------------------------------------- #
#  Điểm vào
# --------------------------------------------------------------------------- #
def validate_pair(tsv_path, xlsx_path, opts, tsv_only=False) -> Report:
    tsv_path = Path(tsv_path)
    rep = Report(tsv_path.name + ("" if tsv_only else f"  +  {Path(xlsx_path).name}"))
    entries = read_tsv(tsv_path, rep, opts)
    if tsv_only:
        return rep
    xlsx_path = Path(xlsx_path)
    if not xlsx_path.exists():
        rep.err("X_MISSING", str(xlsx_path), "đề yêu cầu nộp đồng thời .tsv và .xlsx")
        return rep
    xrows = read_xlsx(xlsx_path, rep)
    if xrows is not None and entries:
        compare(rep, entries, xrows)
    return rep


def main(argv=None) -> int:
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8", errors="replace")
        except Exception:                   # noqa: BLE001
            pass
    ap = argparse.ArgumentParser(description="Kiểm tra file dóng hàng Hán–Việt (TSV + XLSX).")
    ap.add_argument("tsv", nargs="+", help="một hoặc nhiều file .tsv (hỗ trợ ký tự đại diện * ?)")
    ap.add_argument("--xlsx", help="đường dẫn .xlsx (chỉ dùng khi kiểm tra một file .tsv)")
    ap.add_argument("--tsv-only", action="store_true", help="chỉ kiểm tra TSV, bỏ qua XLSX")
    ap.add_argument("--report", help="ghi kết quả ra file Markdown (dùng cho báo cáo nộp)")
    ap.add_argument("--max-show", type=int, default=10, help="số ví dụ tối đa mỗi loại (mặc định 10)")
    ap.add_argument("--strict", action="store_true", help="coi cảnh báo như lỗi")
    ap.add_argument("--ratio-low", type=float, default=DEFAULT_RATIO_LOW)
    ap.add_argument("--ratio-high", type=float, default=DEFAULT_RATIO_HIGH)
    args = ap.parse_args(argv)

    files = []
    for a in args.tsv:
        files += sorted(glob.glob(a)) if any(ch in a for ch in "*?[") else [a]
    if not files:
        print("Không tìm thấy file nào khớp.", file=sys.stderr)
        return 2
    if args.xlsx and len(files) != 1:
        print("--xlsx chỉ dùng được với đúng một file .tsv.", file=sys.stderr)
        return 2
    for f in files:
        if not Path(f).is_file():
            print(f"Không tìm thấy file: {f}", file=sys.stderr)
            return 2

    opts = SimpleNamespace(ratio_low=args.ratio_low, ratio_high=args.ratio_high)
    results, failed = [], False
    for f in files:
        xlsx = args.xlsx or str(Path(f).with_suffix(".xlsx"))
        rep = validate_pair(f, xlsx, opts, args.tsv_only)
        print_report(rep, args.max_show)
        results.append((rep, f, xlsx))
        failed |= rep.n_err > 0 or (args.strict and rep.n_warn > 0)
    if args.report:
        write_markdown(args.report, results)
        print(f"\nĐã ghi báo cáo: {args.report}")
    return 1 if failed else 0


if __name__ == "__main__":
    sys.exit(main())