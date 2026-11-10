import json
from pathlib import Path
import re
from statistics import median

import pdfplumber


ROOT_DIR = Path(__file__).resolve().parent
PROJECT_DIR = ROOT_DIR / "CV05_CPR_03"
PDF_PATH = (
    PROJECT_DIR
    / "01_nguon"
    / "Tam Quoc Dien Nghia (Han - Viet) - quyen 3.pdf"
)
OUTPUT_DIR = PROJECT_DIR / "02_ocr_tho"
TEXT_OUTPUT_COLUMNS = {2: "han", 3: "han_viet", 4: "viet"}
START_PAGE_NUMBER = 4
COLUMN_COUNT = 4
MIN_GUTTER_WIDTH = 8
GUTTER_SEARCH_MARGIN = 0.15
LINE_TOP_TOLERANCE = 2.5
HEADER_TOP_LIMIT = 110
HEADER_MARKERS = ("TAM QUỐC DIỄN NGHĨA", "DỊCH GIẢ:")

PUNCTUATION_SPACING = re.compile(r"\s+([,.;:!?，。！？；：、])")
PAGE_NUMBER = re.compile(r"\d+")
CHAPTER_START = re.compile(
    r"^(?:第[一二三四五六七八九十百千万零〇0-9]+回|HỒI\s+\d+|"
    r"Đệ\s+.+?\bhồi\b|d[iìíĭ]\s+.+?\bhu[iíìĭ]\b)",
    re.I,
)


def find_column_boundaries(page):
    intervals = sorted(
        (word["x0"], word["x1"])
        for word in page.extract_words()
    )
    merged = []

    for left, right in intervals:
        if merged and left <= merged[-1][1]:
            merged[-1] = (merged[-1][0], max(merged[-1][1], right))
        else:
            merged.append((left, right))

    min_x = page.width * GUTTER_SEARCH_MARGIN
    max_x = page.width * (1 - GUTTER_SEARCH_MARGIN)
    gutters = []
    for (_, left_right), (right_left, _) in zip(merged, merged[1:]):
        gap = right_left - left_right
        midpoint = (left_right + right_left) / 2
        if gap >= MIN_GUTTER_WIDTH and min_x < midpoint < max_x:
            gutters.append((gap, midpoint))

    if len(gutters) < COLUMN_COUNT - 1:
        return None

    widest = sorted(gutters, reverse=True)[: COLUMN_COUNT - 1]
    return sorted(midpoint for _, midpoint in widest)


def find_header_cutoff(page, words):
    rows = {}
    for word in words:
        if word["top"] <= HEADER_TOP_LIMIT:
            key = round(word["top"], 1)
            rows.setdefault(key, []).append(word)

    header_bottoms = []
    for row_words in rows.values():
        row_text = " ".join(
            word["text"] for word in sorted(row_words, key=lambda word: word["x0"])
        ).upper()
        if any(marker in row_text for marker in HEADER_MARKERS):
            header_bottoms.append(max(word["bottom"] for word in row_words))

    return max(header_bottoms) if header_bottoms else None


def remove_page_furniture(page, words):
    header_cutoff = find_header_cutoff(page, words)
    rows = {}
    for word in words:
        key = round(word["top"], 1)
        rows.setdefault(key, []).append(word)

    filtered = []
    for row_words in rows.values():
        row_words.sort(key=lambda word: word["x0"])
        row_text = "".join(word["text"] for word in row_words).strip()
        row_top = min(word["top"] for word in row_words)

        if header_cutoff is not None and row_top <= header_cutoff:
            continue
        if row_top >= page.height * 0.85 and PAGE_NUMBER.fullmatch(row_text):
            continue

        filtered.extend(row_words)

    return filtered


def group_words_into_lines(words, left, right):
    selected = [
        word
        for word in words
        if left <= (word["x0"] + word["x1"]) / 2 < right
    ]
    selected.sort(key=lambda word: (word["top"], word["x0"]))

    lines = []
    for word in selected:
        if not lines or word["top"] - lines[-1]["top"] > LINE_TOP_TOLERANCE:
            lines.append(
                {
                    "top": word["top"],
                    "bottom": word["bottom"],
                    "words": [word],
                }
            )
        else:
            line = lines[-1]
            line["top"] = min(line["top"], word["top"])
            line["bottom"] = max(line["bottom"], word["bottom"])
            line["words"].append(word)

    return lines


def line_text(line, column_number):
    words = sorted(line["words"], key=lambda word: word["x0"])
    separator = "" if column_number == 2 else " "
    text = separator.join(word["text"] for word in words)
    return PUNCTUATION_SPACING.sub(r"\1", text).strip()


def group_lines_into_paragraphs(lines, column_number):
    if not lines:
        return []

    heights = [
        word["bottom"] - word["top"]
        for line in lines
        for word in line["words"]
    ]
    paragraph_gap = max(18.0, median(heights) * 1.8)
    groups = []
    current = [lines[0]]

    for line in lines[1:]:
        if line["top"] - current[-1]["top"] > paragraph_gap:
            groups.append(current)
            current = []
        current.append(line)
    groups.append(current)

    separator = "" if column_number == 2 else " "
    paragraphs = []
    for group in groups:
        text = separator.join(line_text(line, column_number) for line in group)
        paragraphs.append(
            {
                "text": PUNCTUATION_SPACING.sub(r"\1", text).strip(),
                "y_start": min(line["top"] for line in group),
                "y_end": max(line["bottom"] for line in group),
            }
        )

    return [paragraph for paragraph in paragraphs if paragraph["text"]]


def ends_with_sentence_punctuation(text, column_number):
    closing = "\"'”’」』〉》)]}» "
    text = text.rstrip(closing)
    punctuation = "。！？" if column_number == 2 else ".?!"
    return text.endswith(tuple(punctuation))


def join_page_continuation(previous, current, column_number, page_height):
    if not previous or not current:
        return False
    if CHAPTER_START.match(current["text"]):
        return False
    if previous["y_end"] < page_height * 0.85:
        return False
    if current["y_start"] > HEADER_TOP_LIMIT + 20:
        return False
    return not ends_with_sentence_punctuation(previous["text"], column_number)


def extract_paragraphs(pdf, boundaries):
    paragraphs_by_column = [[] for _ in range(COLUMN_COUNT)]
    x_edges = [0, *boundaries, pdf.pages[0].width]

    for page_number, page in enumerate(
        pdf.pages[START_PAGE_NUMBER - 1 :],
        start=START_PAGE_NUMBER,
    ):
        words = remove_page_furniture(page, page.extract_words())

        for column_index in range(COLUMN_COUNT):
            column_number = column_index + 1
            lines = group_words_into_lines(
                words,
                x_edges[column_index],
                x_edges[column_index + 1],
            )
            page_paragraphs = group_lines_into_paragraphs(lines, column_number)
            existing = paragraphs_by_column[column_index]

            if (
                page_paragraphs
                and existing
                and join_page_continuation(
                    existing[-1],
                    page_paragraphs[0],
                    column_number,
                    page.height,
                )
            ):
                separator = "" if column_number == 2 else " "
                existing[-1]["text"] += separator + page_paragraphs[0]["text"]
                existing[-1]["page_end"] = page_number
                existing[-1]["y_end"] = page_paragraphs[0]["y_end"]
                page_paragraphs = page_paragraphs[1:]

            for paragraph in page_paragraphs:
                paragraph["page_start"] = page_number
                paragraph["page_end"] = page_number
                paragraph["column"] = column_number
                existing.append(paragraph)

    return paragraphs_by_column


def write_outputs(paragraphs_by_column):
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    metadata_path = OUTPUT_DIR / "metadata.jsonl"

    with metadata_path.open("w", encoding="utf-8") as metadata_file:
        for column_index, paragraphs in enumerate(paragraphs_by_column, start=1):
            language = TEXT_OUTPUT_COLUMNS.get(column_index)
            if language is None:
                continue

            language_dir = OUTPUT_DIR / language
            language_dir.mkdir(parents=True, exist_ok=True)
            text_path = language_dir / f"{PDF_PATH.stem}.txt"
            with text_path.open("w", encoding="utf-8") as text_file:
                for paragraph_index, paragraph in enumerate(paragraphs, start=1):
                    paragraph_id = f"{paragraph_index:06d}"
                    text_file.write(f"===đoạn {paragraph_id}===\n")
                    text_file.write(paragraph["text"] + "\n\n")

                    record = {
                        "paragraph_id": f"c{column_index}-p{paragraph_id}",
                        "page_start": paragraph["page_start"],
                        "page_end": paragraph["page_end"],
                        "column": column_index,
                        "y_start": round(paragraph["y_start"], 2),
                        "y_end": round(paragraph["y_end"], 2),
                        "text": paragraph["text"],
                    }
                    metadata_file.write(
                        json.dumps(record, ensure_ascii=False) + "\n"
                    )


def main():
    if not PDF_PATH.is_file():
        raise FileNotFoundError(f"Không tìm thấy PDF: {PDF_PATH}")

    with pdfplumber.open(PDF_PATH) as pdf:
        if START_PAGE_NUMBER > len(pdf.pages):
            raise ValueError(
                f"PDF chỉ có {len(pdf.pages)} trang; "
                f"không thể bắt đầu từ trang {START_PAGE_NUMBER}."
            )

        boundaries = None
        for page in pdf.pages[START_PAGE_NUMBER - 1 :]:
            boundaries = find_column_boundaries(page)
            if boundaries is not None:
                break

        if boundaries is None:
            raise ValueError("Không tìm thấy ba rãnh để xác định bốn cột.")

        paragraphs = extract_paragraphs(pdf, boundaries)
        write_outputs(paragraphs)

    print(f"Đã trích xuất metadata và text thô vào: {OUTPUT_DIR}")
    print(f"Metadata: {OUTPUT_DIR / 'metadata.jsonl'}")
    print(
        "Số đoạn theo cột: "
        + ", ".join(str(len(column)) for column in paragraphs)
    )


if __name__ == "__main__":
    main()
