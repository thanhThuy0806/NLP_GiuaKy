from pathlib import Path
import re


ROOT_DIR = Path(__file__).resolve().parent
PROJECT_DIR = ROOT_DIR / "CV05_CPR_03"
INPUT_DIRS = {
    "han": PROJECT_DIR / "02_ocr_tho" / "han",
    "han_viet": PROJECT_DIR / "02_ocr_tho" / "han_viet",
    "viet": PROJECT_DIR / "02_ocr_tho" / "viet",
}
OUTPUT_DIR = PROJECT_DIR / "03_text_sach"
SENTENCE_OUTPUT_DIR = OUTPUT_DIR / "sentence"

PAGE_NUMBER_LINE = re.compile(r"[0-9]+")
REFERENCE_NUMBER = re.compile(r"\[\s*\d+\s*\]|\(\s*\d+\s*\)")
FOOTNOTE_LINE = re.compile(
    r"^\s*(?:\[\s*\d+\s*\]|\(\s*\d+\s*\)|[*†‡])\s*\S"
)
FOOTNOTE_LABEL = re.compile(r"^\s*(?:chú thích|ghi chú)\s*[:：]", re.IGNORECASE)
PARAGRAPH_MARKER = re.compile(r"^===đoạn\s+(\d+)===$", re.IGNORECASE)
PAGE_HEADER_TEXT = re.compile(
    r"1\s*三国演义\s*[-–—]\s*罗贯中|"
    r"PINYIN\s*[–—-]\s*QUAN\s*THOẠI",
    re.IGNORECASE,
)
PAGE_HEADER_LINES = {
    "1 三国演义 - 罗贯中",
    "pinyin – quan thoại",
}
HAN_CLAUSE_STARTERS = (
    "却说",
    "次日",
    "忽报",
    "于是",
    "且说",
    "不知",
    "当下",
    "再说",
    "只见",
    "正是",
)
CLOSING_QUOTES = "\"'”’」』〉》)]}»"
NAME_ABBREVIATIONS = {
    "tp.",
    "q.",
    "p.",
    "gs.",
    "pgs.",
    "ts.",
    "ths.",
    "bs.",
    "ks.",
    "dr.",
    "mr.",
    "mrs.",
}


def is_footer_page_number(lines, index):
    """Only treat a number as a page number when it precedes a page break."""
    if not PAGE_NUMBER_LINE.fullmatch(lines[index].strip()):
        return False

    next_line = lines[index + 1].strip() if index + 1 < len(lines) else ""
    return next_line == ""


def clean_text(text, remove_page_numbers=True):
    text = re.sub(r"[\x00-\x08\x0b\x0c\x0e-\x1f\ufffd]", "", text)
    text = PAGE_HEADER_TEXT.sub("", text)
    lines = text.splitlines()
    cleaned_lines = []
    removed_page_numbers = 0
    removed_footnotes = 0

    for index, line in enumerate(lines):
        if remove_page_numbers and is_footer_page_number(lines, index):
            removed_page_numbers += 1
            continue

        if FOOTNOTE_LINE.match(line) or FOOTNOTE_LABEL.match(line):
            removed_footnotes += 1
            continue

        if line.strip().casefold() in PAGE_HEADER_LINES:
            continue

        line = REFERENCE_NUMBER.sub("", line)
        line = re.sub(r"[ \t]+", " ", line).strip()
        line = re.sub(r"\s+([,.;:!?])", r"\1", line)

        if not line:
            if cleaned_lines and cleaned_lines[-1] == "":
                continue
            cleaned_lines.append("")
        else:
            if cleaned_lines and cleaned_lines[-1] == line:
                continue
            cleaned_lines.append(line)

    return "\n".join(cleaned_lines).strip(), removed_page_numbers, removed_footnotes


def parse_paragraph_blocks(text):
    blocks = []
    paragraph_id = None
    paragraph_lines = []
    found_markers = False

    for line in text.splitlines():
        match = PARAGRAPH_MARKER.fullmatch(line.strip())
        if match:
            found_markers = True
            if paragraph_id is not None:
                blocks.append((paragraph_id, "\n".join(paragraph_lines).strip()))
            paragraph_id = match.group(1)
            paragraph_lines = []
        else:
            paragraph_lines.append(line)

    if paragraph_id is not None:
        blocks.append((paragraph_id, "\n".join(paragraph_lines).strip()))

    if not found_markers and text.strip():
        return [("", text)]
    return blocks


def join_wrapped_lines(text, is_chinese):
    lines = [line.strip() for line in text.splitlines() if line.strip()]
    separator = "" if is_chinese else " "
    return separator.join(lines)


def next_nonspace(text, index):
    while index < len(text) and text[index].isspace():
        index += 1
    return text[index] if index < len(text) else ""


def separate_chapter_heading(text, language):
    if language == "han":
        compact_text = join_wrapped_lines(text, is_chinese=True)
        if re.match(r"^第[一二三四五六七八九十百千万零〇0-9]+回", compact_text):
            for marker in ("却说", "话说", "且说"):
                position = compact_text.find(marker)
                if position > 0:
                    return compact_text[:position], compact_text[position:]

    if language == "han_viet":
        compact_text = join_wrapped_lines(text, is_chinese=False)
        if compact_text.casefold().startswith("đệ bát thập"):
            position = compact_text.casefold().find("khước thuyết")
            if position > 0:
                return compact_text[:position], compact_text[position:]

    return "", text


def is_chinese_sentence_end(text, index):
    character = text[index]
    if character in "。！？":
        return True

    if character == "；":
        following_text = text[index + 1 :].lstrip()
        return following_text.startswith(
            ("“", "‘", "「", "『") + HAN_CLAUSE_STARTERS
        )

    return False


def is_abbreviation_period(text, index, sentence_start):
    prefix = text[sentence_start : index + 1]
    match = re.search(
        r"(?i)\b(?:v\.v|tp|q|p|pgs|gs|ts|ths|bs|ks|dr|mr|mrs)\.$",
        prefix,
    )
    following = next_nonspace(text, index + 1)

    if match:
        abbreviation = match.group().casefold()
        if abbreviation in NAME_ABBREVIATIONS:
            return True
        if abbreviation == "v.v." and following.islower():
            return True

    initials = re.search(r"(?:^|\s)(?:[A-ZĐ]\.){1,3}$", prefix)
    return bool(initials and following.isupper())


def is_vietnamese_sentence_end(text, index, sentence_start):
    character = text[index]
    following = next_nonspace(text, index + 1)

    if character in "?!":
        return True

    if character == ".":
        previous = text[index - 1] if index else ""
        next_character = text[index + 1] if index + 1 < len(text) else ""
        if previous.isdigit() and next_character.isdigit():
            return False

        prefix = text[sentence_start:index]
        if re.fullmatch(r"\s*\d+", prefix):
            return False

        return not is_abbreviation_period(text, index, sentence_start)

    if character == ";":
        return following.isupper() or following in "“‘\""

    if character == ":":
        previous = text[index - 1] if index else ""
        next_character = text[index + 1] if index + 1 < len(text) else ""
        if previous.isdigit() and next_character.isdigit():
            return False
        if following in "-–—“‘\"":
            return False
        return following.isupper()

    return False


def split_sentences(text, is_chinese=False, language=""):
    heading, text = separate_chapter_heading(text, language)
    text = join_wrapped_lines(text, is_chinese)
    sentences = [heading] if heading else []
    sentence_start = 0
    index = 0

    while index < len(text):
        if is_chinese:
            is_end = is_chinese_sentence_end(text, index)
        else:
            is_end = is_vietnamese_sentence_end(text, index, sentence_start)

        if not is_end:
            index += 1
            continue

        sentence_end = index + 1
        while sentence_end < len(text):
            quote_index = sentence_end
            while quote_index < len(text) and text[quote_index].isspace():
                quote_index += 1
            if quote_index < len(text) and text[quote_index] in CLOSING_QUOTES:
                sentence_end = quote_index + 1
            else:
                break

        sentence = text[sentence_start:sentence_end].strip()
        if sentence:
            sentences.append(sentence)
        sentence_start = sentence_end
        index = sentence_end

    remaining = text[sentence_start:].strip()
    if remaining:
        sentences.append(remaining)

    return sentences


def main():
    input_files = [
        (language, input_path)
        for language, input_dir in INPUT_DIRS.items()
        for input_path in sorted(input_dir.glob("*.txt"))
    ]
    if not input_files:
        input_dirs = ", ".join(str(path) for path in INPUT_DIRS.values())
        raise FileNotFoundError(f"Không tìm thấy file .txt trong: {input_dirs}")

    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    SENTENCE_OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    for language, input_path in input_files:
        text = input_path.read_text(encoding="utf-8")
        paragraph_blocks = parse_paragraph_blocks(text)
        cleaned_blocks = []
        sentence_blocks = []
        page_count = 0
        footnote_count = 0

        for paragraph_id, paragraph_text in paragraph_blocks:
            cleaned_text, removed_pages, removed_footnotes = clean_text(
                paragraph_text,
                remove_page_numbers=not bool(paragraph_id),
            )
            page_count += removed_pages
            footnote_count += removed_footnotes
            if not cleaned_text:
                continue

            cleaned_blocks.append((paragraph_id, cleaned_text))
            is_chinese = language == "han"
            sentences = split_sentences(
                cleaned_text,
                is_chinese=is_chinese,
                language=language,
            )
            sentence_blocks.append((paragraph_id, sentences))

        output_name = f"{input_path.stem}_{language}.txt"
        clean_path = OUTPUT_DIR / output_name
        clean_lines = []
        sentence_lines = []
        for paragraph_id, cleaned_text in cleaned_blocks:
            if paragraph_id:
                clean_lines.append(f"===đoạn {paragraph_id}===")
            clean_lines.extend([cleaned_text, ""])
        clean_path.write_text("\n".join(clean_lines).rstrip() + "\n", encoding="utf-8")

        for paragraph_id, sentences in sentence_blocks:
            if paragraph_id:
                sentence_lines.append(f"===đoạn {paragraph_id}===")
            sentence_lines.extend(sentences)
            sentence_lines.append("")

        sentence_path = SENTENCE_OUTPUT_DIR / output_name
        sentence_path.write_text("\n".join(sentence_lines).rstrip() + "\n", encoding="utf-8")
        print(
            f"{input_path.name}: bỏ {page_count} số trang, "
            f"{footnote_count} dòng chú thích; "
            f"tạo {sum(len(sentences) for _, sentences in sentence_blocks)} câu "
            f"-> {sentence_path}"
        )


if __name__ == "__main__":
    main()
