#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
extract_cols_cpr03.py - Tách 4 cột của Tam quốc diễn nghĩa (CPR_03, PDF text) thành các file text.

Bố cục mỗi trang (theo trang mẫu): 4 cột phân biệt bằng vị trí và màu chữ
    cột 1  Pinyin                      -> bỏ (trừ khi dùng --keep-pinyin)
    cột 2  Hán văn (giản thể)          -> han.txt        (vế Hán của đồ án)
    cột 3  Phiên âm Hán Việt           -> hanviet.txt    (KHÔNG nộp; chỉ làm cầu nối dóng hàng)
    cột 4  Bản dịch Việt (Phan Kế Bính)-> viet.txt       (vế Việt của đồ án)
Ngoài ra bỏ: số thứ tự hàng ở lề trái (1, 2...), dòng tiêu đề sách căn giữa trang.

Mỗi hồi là một đơn vị. Đầu hồi có khối tiêu đề in đậm (第八十一回 / HỒI 81...), được tách riêng thành tiêu đề và KHÔNG lẫn vào thân bài. Mỗi hồi được ghép thành MỘT dòng trong han.txt, viet.txt, hanviet.txt (dòng i của ba file là cùng hồi, có bảng đối chiếu trong index.tsv), để bước tách câu và dóng hàng chạy được TRONG TỪNG HỒI, tránh trôi lệch dây chuyền.

Đầu ra trong thư mục -o:
    han.txt, viet.txt, hanviet.txt   mỗi dòng một hồi (đã nối dòng bị ngắt, nối câu qua trang)
    index.tsv                        hồi, trang đầu/cuối, tiêu đề, số chữ Hán, số âm tiết, cảnh báo
    cells.jsonl                      bản đầy đủ từng hồi, kèm vị trí ký tự bắt đầu của từng trang 
                                     (để truy vết về PDF)
    hoi/HOI_081.han.txt ...          (tùy chọn --per-hoi) mỗi hồi một file

Kiểm tra tự động: số âm tiết cột Hán Việt phải xấp xỉ số chữ Hán của cột Hán (đọc từng chữ một).
Lệch hơn --hv-tol (mặc định 5%) nghĩa là cột nào đó bị mất/dư dòng; script cảnh báo hồi đó.

Cách dùng:
    # 1) Hiệu chỉnh: xem script nhận dạng 4 cột và các hồi có đúng không (làm trước, mỗi tập một lần)
    python scripts/extract_cols_cpr03.py calibrate CPR_03/01_nguon/tap1.pdf --save cols_tap1.json
    python scripts/extract_cols_cpr03.py calibrate tap1.pdf --dump --page 3     # soi từng dòng của một trang

    # 2) Tách
    python scripts/extract_cols_cpr03.py extract tap1.pdf -o CPR_03/03_text_sach/tap1
    python scripts/extract_cols_cpr03.py extract tap1.pdf -o out --config cols_tap1.json --hoi 81-100 --per-hoi

Giả định: mỗi hồi bắt đầu bằng một dòng in đậm kiểu "第八十一回" ở cột Hán (thường ở trang đầu của hồi). Nếu thực tế khác, chạy calibrate để xem script thấy gì rồi chỉnh --config / --drop-regex.
"""
from __future__ import annotations

import argparse
import json
import re
import sys
import unicodedata
from collections import Counter
from pathlib import Path

try:
    import fitz  # PyMuPDF
except ImportError:  # PyMuPDF >= 1.24 cũng cho phép "import pymupdf"
    try:
        import pymupdf as fitz
    except ImportError:
        sys.exit("Chưa cài PyMuPDF. Chạy: pip install pymupdf")

ROLES = ["pinyin", "han", "hv", "viet"]
CJK_RE = re.compile(r"[\u3400-\u4dbf\u4e00-\u9fff\uf900-\ufaff\U00020000-\U0003134f]")
SYL_RE = re.compile(r"[^\W\d_]+", re.UNICODE)       # một âm tiết = một cụm chữ cái
HOI_RE = re.compile(r"^第\s*([〇零一二三四五六七八九十百两兩\d]+)\s*[回章]")
HOI_VI_RE = re.compile(r"H[ỒO]I\s*(\d+)")
DEFAULT_DROP = [                                      # dòng tiêu đề sách lặp lại (so khớp chữ HOA)
    r"^TAM\s+QU[ỐỒOÓ]C\s+DI[ỄE]N\s+NGH",
    r"^D[ỊI]CH\s+GI[ẢA]\s*:",
    r"^T[ÁA]C\s+GI[ẢA]\s*:",
]
ZH = {"〇": 0, "零": 0, "一": 1, "二": 2, "两": 2, "兩": 2, "三": 3, "四": 4,
      "五": 5, "六": 6, "七": 7, "八": 8, "九": 9}


# --------------------------------------------------------------------------- #
#  Tiện ích
# --------------------------------------------------------------------------- #
def zh_to_int(s: str):
    """八十一 -> 81, 一百二十 -> 120, 〇八一 -> 81, '81' -> 81. Trả về None nếu không đọc được."""
    s = s.strip()
    if s.isdigit():
        return int(s)
    if not s or any(c not in ZH and c not in "十百" for c in s):
        return None
    if "十" not in s and "百" not in s:
        return int("".join(str(ZH[c]) for c in s))
    total = num = 0
    for c in s:
        if c in ZH:
            num = ZH[c]
        elif c == "十":
            total += (num or 1) * 10
            num = 0
        else:                                         # 百
            total += (num or 1) * 100
            num = 0
    return total + num


def clean(t: str, role: str) -> str:
    t = unicodedata.normalize("NFC", t.replace("\u00a0", " "))
    return re.sub(r"\s+", "", t) if role == "han" else re.sub(r"\s+", " ", t).strip()


def parse_range(s):
    if not s:
        return None
    a, _, b = s.partition("-")
    lo = int(a)
    hi = int(b) if b else lo
    return lo, hi


# --------------------------------------------------------------------------- #
#  Đọc dòng chữ từ PDF
# --------------------------------------------------------------------------- #
def page_lines(page):
    """Các dòng chữ của một trang: x0, x1, y0, text, bold, size, color."""
    out = []
    for b in page.get_text("dict")["blocks"]:
        if b.get("type", 0) != 0:
            continue
        for ln in b.get("lines", []):
            spans = ln.get("spans", [])
            text = "".join(s["text"] for s in spans)
            if not text.strip():
                continue
            w = Counter()
            bold_w = size_w = 0
            sizes, colors = Counter(), Counter()
            for s in spans:
                n = len(s["text"].strip())
                if not n:
                    continue
                w[0] += n
                is_bold = bool(s["flags"] & 16) or "bold" in s["font"].lower()
                bold_w += n if is_bold else 0
                sizes[round(s["size"], 1)] += n
                colors[s["color"]] += n
            x0, y0, x1, _ = ln["bbox"]
            out.append({
                "x0": x0, "x1": x1, "y0": y0, "text": text,
                "bold": bold_w * 2 > w[0],
                "size": sizes.most_common(1)[0][0],
                "color": "#%06x" % colors.most_common(1)[0][0],
            })
    return out


# --------------------------------------------------------------------------- #
#  Nhận dạng cột
# --------------------------------------------------------------------------- #
def cluster_x(lines, tol=4.0):
    """Gom các dòng theo x0 gần nhau. Trả về danh sách cụm đã sắp theo x."""
    clusters = []
    for ln in sorted(lines, key=lambda l: l["x0"]):
        n = len(ln["text"].strip())
        if clusters and ln["x0"] - clusters[-1]["x"] <= tol:
            c = clusters[-1]
            c["x"] = (c["x"] * c["lines"] + ln["x0"]) / (c["lines"] + 1)
            c["lines"] += 1
            c["chars"] += n
            c["items"].append(ln)
        else:
            clusters.append({"x": ln["x0"], "lines": 1, "chars": n, "items": [ln]})
    return clusters


def sample_pages(n, k=14):
    if n <= k:
        return list(range(n))
    return sorted({int(i * (n - 1) / (k - 1)) for i in range(k)})


def detect_columns(doc, pages=None):
    """Tự nhận dạng 4 cột từ các trang mẫu. Trả về (cfg, cụm, cảnh báo)."""
    pages = pages if pages is not None else sample_pages(len(doc))
    lines = []
    page_w = None
    for p in pages:
        pg = doc.load_page(p)
        page_w = pg.rect.width
        lines += page_lines(pg)
    if not lines:
        raise SystemExit("Không đọc được chữ nào từ các trang mẫu. Đây có phải PDF scan (cần OCR) không?")
    clusters = cluster_x(lines)
    top = sorted(sorted(clusters, key=lambda c: -c["chars"])[:4], key=lambda c: c["x"])
    warns = []
    if len(top) < 4:
        raise SystemExit(f"Chỉ tìm thấy {len(top)} cụm cột (cần 4). Chạy: calibrate --dump --page N để xem.")
    extra = [c for c in clusters if c not in top and c["chars"] >= 0.08 * top[1]["chars"]]
    if extra:
        warns.append("Có cụm x khác đáng kể ngoài 4 cột: " + ", ".join(f"x≈{c['x']:.0f} ({c['chars']} ký tự)" for c in extra))

    for role, c in zip(ROLES, top):
        tot = sum(len(i["text"].strip()) for i in c["items"])
        cjk = sum(len(CJK_RE.findall(i["text"])) for i in c["items"])
        c["role"], c["cjk"] = role, cjk / tot if tot else 0
    if top[1]["cjk"] < 0.5 or any(c["cjk"] > 0.2 for i, c in enumerate(top) if i != 1):
        raise SystemExit(
            "Nhận dạng cột không hợp lý (cột 2 phải là chữ Hán, các cột khác là chữ Latinh). "
            "Chi tiết %CJK theo cột: " + ", ".join(f"{c['role']}={c['cjk']:.0%}" for c in top) +
            ". Hãy chạy 'calibrate' và dùng --config.")

    cols = []
    for k, c in enumerate(top):
        nxt = top[k + 1]["x"] - 8 if k + 1 < len(top) else (page_w or 10_000) + 1
        cols.append({"role": c["role"], "x_min": round(c["x"] - 8, 1), "x_max": round(nxt, 1)})
    spacing = sorted(top[k + 1]["x"] - top[k]["x"] for k in range(3))[1]      # trung vị khoảng cách cột
    body_size = {}
    for c in top:
        cnt = Counter()
        for i in c["items"]:
            if not i["bold"]:
                cnt[i["size"]] += len(i["text"].strip())
        body_size[c["role"]] = cnt.most_common(1)[0][0] if cnt else 0
    cfg = {"cols": cols, "span_limit": round(spacing * 1.2, 1), "body_size": body_size}
    return cfg, top, warns


def classify(lines, cfg):
    out = []
    for ln in lines:
        role = next((c["role"] for c in cfg["cols"] if c["x_min"] <= ln["x0"] < c["x_max"]), None)
        ln = dict(ln, role=role)
        ln["width"] = ln["x1"] - ln["x0"]
        bs = cfg["body_size"].get(role) or 0
        ln["heading"] = bool(ln["bold"] or (bs and ln["size"] >= bs * 1.08))
        out.append(ln)
    return out


def is_marker(ln):
    if ln["role"] != "han":
        return None
    t = clean(ln["text"], "han")
    m = HOI_RE.match(t)
    if m and (ln["heading"] or len(t) <= 12):
        return zh_to_int(m.group(1)), t
    return None


# --------------------------------------------------------------------------- #
#  Tách theo hồi
# --------------------------------------------------------------------------- #
def new_record(hoi, marker_text, page):
    return {"hoi": hoi, "marker": marker_text, "page_start": page, "page_end": page,
            "titles": {"han": [], "hv": [], "viet": []},
            "body": {"pinyin": [], "han": [], "hv": [], "viet": []}}


def extract_records(doc, cfg, page_idx, drop_res, keep_pinyin=False, tail_gap=6.0):
    records, cur = [], None
    stats = Counter()
    for pno in page_idx:
        lines = classify(page_lines(doc.load_page(pno)), cfg)
        page = pno + 1
        marker = None
        for ln in sorted(lines, key=lambda l: (l["y0"], l["x0"])):
            if ln["width"] <= cfg["span_limit"] and (mk := is_marker(ln)):
                marker = (mk, ln["y0"])
                break
        if marker:
            (hoi, mtext), marker_y = marker
            new = new_record(hoi, mtext, page)
        else:
            marker_y = None
            new = None
        for ln in sorted(lines, key=lambda l: (l["role"] or "", round(l["y0"], 1), l["x0"])):
            role = ln["role"]
            text = ln["text"]
            if role is None:
                stats["lề trái/số hàng"] += 1
                continue
            if ln["width"] > cfg["span_limit"]:
                stats["dòng tiêu đề căn giữa"] += 1
                continue
            if any(r.search(text.upper().strip()) for r in drop_res):
                stats["khớp --drop-regex"] += 1
                continue
            if role == "pinyin" and not keep_pinyin:
                stats["pinyin"] += 1
                continue
            if new is not None and ln["heading"] and role != "pinyin":
                new["titles"][role].append(clean(text, role))
                continue
            target = cur
            if new is not None:                      # trang mở đầu hồi mới
                if ln["y0"] < marker_y - tail_gap and role != "pinyin":
                    target = cur                     # phần cuối hồi trước nằm cùng trang
                else:
                    target = new
            if target is None:
                stats["trước hồi đầu tiên"] += 1
                continue
            if ln["heading"] and new is None:
                stats["chữ đậm ngoài trang đầu hồi (giữ trong thân bài)"] += 1
            target["body"][role].append((page, text))
            target["page_end"] = max(target["page_end"], page)
        if new is not None:
            if cur is not None:
                records.append(cur)
            cur = new
    if cur is not None:
        records.append(cur)
    return records, stats


def join_lines(items, role):
    sep = "" if role == "han" else " "
    parts, offsets, pos, last = [], [], 0, None
    for pg, t in items:
        t = clean(t, role)
        if not t:
            continue
        if pg != last:
            offsets.append([pg, pos])
            last = pg
        parts.append(t)
        pos += len(t) + len(sep)
    return sep.join(parts), offsets


def finalize(rec, args, prev_hoi):
    out = {"hoi": rec["hoi"], "page_start": rec["page_start"], "page_end": rec["page_end"]}
    for role in ("han", "hv", "viet"):
        out["title_" + role] = " ".join(t for t in rec["titles"][role] if t)
        out[role], out["pages_" + role] = join_lines(rec["body"][role], role)
    if args.keep_pinyin:
        out["pinyin"], out["pages_pinyin"] = join_lines(rec["body"]["pinyin"], "pinyin")

    n_han = len(CJK_RE.findall(out["han"]))
    n_hv = len(SYL_RE.findall(out["hv"]))
    n_vi = len(SYL_RE.findall(out["viet"]))
    out.update(han_chars=n_han, hv_syll=n_hv, viet_syll=n_vi,
               hv_diff=(n_hv - n_han) / n_han if n_han else None)

    w = []
    m = HOI_VI_RE.search(out["title_viet"].upper())
    if out["hoi"] is None and m:
        out["hoi"] = int(m.group(1))
    elif m and out["hoi"] is not None and int(m.group(1)) != out["hoi"]:
        w.append(f"số hồi Hán ({out['hoi']}) ≠ số hồi Việt ({m.group(1)})")
    if out["hoi"] is not None and prev_hoi is not None and out["hoi"] != prev_hoi + 1:
        w.append(f"số hồi không liên tục (trước đó là {prev_hoi})")
    for role, label in (("han", "Hán"), ("hv", "Hán Việt"), ("viet", "Việt")):
        if not out[role]:
            w.append(f"cột {label} rỗng")
        if not out["title_" + role]:
            w.append(f"không thấy tiêu đề {label}")
    if n_han and out["hv_diff"] is not None and abs(out["hv_diff"]) > args.hv_tol:
        w.append(f"âm tiết Hán Việt ({n_hv}) lệch {out['hv_diff']:+.1%} so với chữ Hán ({n_han})")
    if n_han and not (args.ratio_low <= n_vi / n_han <= args.ratio_high):
        w.append(f"tỷ lệ âm tiết Việt/chữ Hán = {n_vi / n_han:.2f} bất thường")
    out["warnings"] = w
    return out


# --------------------------------------------------------------------------- #
#  Lệnh: calibrate
# --------------------------------------------------------------------------- #
def open_pdf(path):
    if not Path(path).is_file():
        sys.exit(f"Không tìm thấy file: {path}")
    doc = fitz.open(path)
    if getattr(doc, "needs_pass", False):
        sys.exit("PDF có mật khẩu.")
    return doc


def cfg_from_args(doc, args):
    if getattr(args, "config", None):
        cfg = json.loads(Path(args.config).read_text(encoding="utf-8"))
        return cfg, None, []
    return detect_columns(doc)


def cmd_calibrate(args):
    doc = open_pdf(args.pdf)
    cfg, top, warns = cfg_from_args(doc, args)
    pg0 = doc.load_page(0)
    print(f"{args.pdf}: {len(doc)} trang, khổ {pg0.rect.width:.0f} x {pg0.rect.height:.0f} pt")
    if top:
        print("\nCụm vị trí x0 (4 cột nhiều chữ nhất, trên các trang mẫu):")
        print(f"{'vai trò':8} {'x0':>7} {'dòng':>6} {'ký tự':>7} {'%Hán':>6}  {'màu chữ chủ đạo':16} {'cỡ chữ thân bài':>15}")
        for c in top:
            col = Counter()
            for i in c["items"]:
                col[i["color"]] += len(i["text"].strip())
            print(f"{c['role']:8} {c['x']:7.1f} {c['lines']:6d} {c['chars']:7d} {c['cjk']:6.0%}  "
                  f"{col.most_common(1)[0][0]:16} {cfg['body_size'][c['role']]:>15}")
    print("\nVùng x của từng cột (dòng có x0 nằm trong vùng thì thuộc cột đó):")
    for c in cfg["cols"]:
        print(f"  {c['role']:7} [{c['x_min']:7.1f}, {c['x_max']:7.1f})")
    print(f"Dòng rộng hơn {cfg['span_limit']} pt bị coi là tiêu đề căn giữa và bỏ.")
    for w in warns:
        print("CẢNH BÁO:", w)
    if args.save:
        Path(args.save).write_text(json.dumps(cfg, ensure_ascii=False, indent=2), encoding="utf-8")
        print(f"\nĐã lưu cấu hình cột: {args.save} (dùng với: extract ... --config {args.save})")

    if args.dump:
        p = (args.page or 1) - 1
        print(f"\n--- Trang {p + 1}: từng dòng sau khi phân loại ---")
        for ln in sorted(classify(page_lines(doc.load_page(p)), cfg), key=lambda l: (l["y0"], l["x0"])):
            flag = ("B" if ln["bold"] else "-") + ("H" if ln["heading"] else "-")
            skip = " (bỏ: căn giữa)" if ln["width"] > cfg["span_limit"] else ""
            print(f"[{ln['role'] or '----':6}] y={ln['y0']:6.1f} x0={ln['x0']:6.1f} w={ln['width']:5.1f} "
                  f"{flag} cỡ={ln['size']:<4} {ln['color']} | {ln['text'][:60]}{skip}")
        print("Chú giải: B = đậm, H = bị coi là tiêu đề (đậm hoặc cỡ chữ lớn hơn thân bài).")
        return 0

    print("\nQuét toàn bộ PDF tìm đầu hồi...")
    drop_res = [re.compile(r, re.I) for r in DEFAULT_DROP + (args.drop_regex or [])]
    recs, stats = extract_records(doc, cfg, range(len(doc)), drop_res)
    if not recs:
        print("KHÔNG tìm thấy dòng đầu hồi nào (kiểu '第八十一回' in đậm ở cột Hán). "
              "Dùng --dump --page N để xem dòng đó được nhận dạng thế nào.")
    else:
        nums = [r["hoi"] for r in recs]
        print(f"Tìm thấy {len(recs)} hồi: " + ", ".join(f"{r['hoi']} (tr.{r['page_start']})" for r in recs[:12]) +
              (" ..." if len(recs) > 12 else ""))
        gaps = [(a, b) for a, b in zip(nums, nums[1:]) if a is None or b is None or b != a + 1]
        if gaps:
            print("CẢNH BÁO: số hồi không liên tục tại:", gaps[:10])
        print(f"Trang trước hồi đầu tiên (bỏ): {stats.get('trước hồi đầu tiên', 0)} dòng.")
    return 0


# --------------------------------------------------------------------------- #
#  Lệnh: extract
# --------------------------------------------------------------------------- #
def cmd_extract(args):
    doc = open_pdf(args.pdf)
    cfg, _, warns = cfg_from_args(doc, args)
    for w in warns:
        print("CẢNH BÁO:", w, file=sys.stderr)
    prange = parse_range(args.pages)
    first, last = (prange[0] - 1, prange[1] - 1) if prange else (0, len(doc) - 1)
    first, last = max(first, 0), min(last, len(doc) - 1)
    drop_res = [re.compile(r, re.I) for r in DEFAULT_DROP + (args.drop_regex or [])]
    print(f"Đang đọc trang {first + 1}–{last + 1} / {len(doc)} ...", file=sys.stderr)
    raw, stats = extract_records(doc, cfg, range(first, last + 1), drop_res, args.keep_pinyin)

    final, prev = [], None
    for r in raw:
        f = finalize(r, args, prev)
        prev = f["hoi"] if f["hoi"] is not None else prev
        final.append(f)
    hr = parse_range(args.hoi)
    if hr:
        final = [f for f in final if f["hoi"] is not None and hr[0] <= f["hoi"] <= hr[1]]
    if not final:
        sys.exit("Không có hồi nào để xuất (kiểm tra --pages/--hoi hoặc chạy calibrate).")

    out = Path(args.output)
    out.mkdir(parents=True, exist_ok=True)
    names = {"han": "han.txt", "viet": "viet.txt", "hv": "hanviet.txt"}
    if args.keep_pinyin:
        names["pinyin"] = "pinyin.txt"
    for role, fn in names.items():
        (out / fn).write_text("\n".join(f.get(role, "") for f in final) + "\n", encoding="utf-8")

    cols = ["dòng", "hồi", "trang_đầu", "trang_cuối", "tiêu_đề_hán", "tiêu_đề_hánviệt", "tiêu_đề_việt",
            "chữ_hán", "âm_tiết_hánviệt", "âm_tiết_việt", "lệch_hv_%", "cảnh_báo"]
    rows = ["\t".join(cols)]
    for i, f in enumerate(final, 1):
        diff = "" if f["hv_diff"] is None else f"{f['hv_diff'] * 100:+.1f}"
        rows.append("\t".join(str(x) for x in [
            i, f["hoi"] if f["hoi"] is not None else "", f["page_start"], f["page_end"],
            f["title_han"], f["title_hv"], f["title_viet"], f["han_chars"], f["hv_syll"],
            f["viet_syll"], diff, "; ".join(f["warnings"])]))
    (out / "index.tsv").write_text("\n".join(rows) + "\n", encoding="utf-8")
    with open(out / "cells.jsonl", "w", encoding="utf-8") as fh:
        for f in final:
            fh.write(json.dumps(f, ensure_ascii=False) + "\n")
    if args.per_hoi:
        d = out / "hoi"
        d.mkdir(exist_ok=True)
        for i, f in enumerate(final, 1):
            tag = f"HOI_{f['hoi']:03d}" if f["hoi"] is not None else f"DONG_{i:03d}"
            for role, fn in (("han", "han"), ("viet", "viet"), ("hv", "hanviet")):
                (d / f"{tag}.{fn}.txt").write_text(f[role] + "\n", encoding="utf-8")

    # ---- tóm tắt ----
    print(f"\nĐã xuất {len(final)} hồi -> {out}/ (han.txt, viet.txt, hanviet.txt, index.tsv, cells.jsonl)")
    print(f"{'hồi':>4} {'trang':>9} {'chữ Hán':>8} {'âm tiết HV':>10} {'âm tiết Việt':>12}  ghi chú")
    for f in final:
        note = "OK" if not f["warnings"] else "; ".join(f["warnings"])
        print(f"{f['hoi'] if f['hoi'] is not None else '?':>4} {str(f['page_start']) + '-' + str(f['page_end']):>9} "
              f"{f['han_chars']:8d} {f['hv_syll']:10d} {f['viet_syll']:12d}  {note}")
    tot = sum(f["han_chars"] for f in final)
    nwarn = sum(1 for f in final if f["warnings"])
    print(f"Tổng: {tot} chữ Hán, {sum(f['viet_syll'] for f in final)} âm tiết Việt; {nwarn} hồi có cảnh báo.")
    if stats:
        print("Đã bỏ qua: " + "; ".join(f"{k}: {v} dòng" for k, v in stats.items()))
    if nwarn:
        print("Hồi có cảnh báo: chạy 'calibrate --dump --page N' ở trang tương ứng để soi nguyên nhân.")
    return 0


# --------------------------------------------------------------------------- #
def main(argv=None):
    for stream in (sys.stdout, sys.stderr):
        try:
            stream.reconfigure(encoding="utf-8", errors="replace")
        except Exception:                             # noqa: BLE001
            pass
    ap = argparse.ArgumentParser(description="Tách 4 cột Tam quốc diễn nghĩa (CPR_03) thành han/viet/hanviet.")
    sub = ap.add_subparsers(dest="cmd", required=True)

    def common(p):
        p.add_argument("pdf")
        p.add_argument("--config", help="file cột .json tạo bởi calibrate --save (bỏ qua nhận dạng tự động)")
        p.add_argument("--drop-regex", action="append", help="bỏ các dòng khớp regex này (có thể lặp lại)")

    c = sub.add_parser("calibrate", help="xem script nhận dạng cột và hồi thế nào")
    common(c)
    c.add_argument("--save", help="lưu cấu hình cột ra file .json")
    c.add_argument("--dump", action="store_true", help="in từng dòng của một trang (kèm --page)")
    c.add_argument("--page", type=int, help="số trang (tính từ 1) cho --dump")
    c.set_defaults(fn=cmd_calibrate)

    e = sub.add_parser("extract", help="tách và xuất text")
    common(e)
    e.add_argument("-o", "--output", required=True, help="thư mục đầu ra")
    e.add_argument("--pages", help="chỉ đọc khoảng trang, ví dụ 5-300 (tính từ 1)")
    e.add_argument("--hoi", help="chỉ xuất khoảng hồi, ví dụ 81-100")
    e.add_argument("--keep-pinyin", action="store_true", help="xuất thêm pinyin.txt")
    e.add_argument("--per-hoi", action="store_true", help="ghi thêm mỗi hồi một file trong hoi/")
    e.add_argument("--hv-tol", type=float, default=0.05, help="ngưỡng lệch âm tiết Hán Việt/chữ Hán (mặc định 0.05)")
    e.add_argument("--ratio-low", type=float, default=0.5, help="ngưỡng dưới âm tiết Việt/chữ Hán")
    e.add_argument("--ratio-high", type=float, default=6.0, help="ngưỡng trên âm tiết Việt/chữ Hán")
    e.set_defaults(fn=cmd_extract)

    args = ap.parse_args(argv)
    return args.fn(args)


if __name__ == "__main__":
    sys.exit(main())