import re
from copy import deepcopy
from pathlib import Path

from docx import Document
from docx.enum.text import WD_COLOR_INDEX
from docx.text.run import Run


folder = Path(__file__).resolve().parent
word_pattern = re.compile(r"[^\W\d_]+(?:[’'-][^\W\d_]+)*", re.UNICODE)


def load_dictionary():
    text = (folder / "dictionary.txt").read_text(encoding="utf-8-sig")
    return {word.group().casefold() for word in word_pattern.finditer(text)}


def highlight_run(run, ranges):
    text = run.text

    if not text or not ranges:
        return

    points = {0, len(text)}

    for start, end in ranges:
        points.add(start)
        points.add(end)

    points = sorted(points)
    parts = []

    for start, end in zip(points, points[1:]):
        highlighted = any(
            start >= word_start and end <= word_end
            for word_start, word_end in ranges
        )
        parts.append((text[start:end], highlighted))

    original_xml = deepcopy(run._r)
    parent = run._r.getparent()
    paragraph = run._parent

    run.text = parts[0][0]

    if parts[0][1]:
        run.font.highlight_color = WD_COLOR_INDEX.YELLOW

    previous_xml = run._r

    for part, highlighted in parts[1:]:
        new_xml = deepcopy(original_xml)
        parent.insert(parent.index(previous_xml) + 1, new_xml)

        new_run = Run(new_xml, paragraph)
        new_run.text = part

        if highlighted:
            new_run.font.highlight_color = WD_COLOR_INDEX.YELLOW

        previous_xml = new_xml


def process_paragraph(paragraph, dictionary):
    runs = list(paragraph.runs)

    if not runs:
        return 0

    text = "".join(run.text for run in runs)

    unknown_words = [
        (match.start(), match.end())
        for match in word_pattern.finditer(text)
        if match.group().casefold() not in dictionary
    ]

    positions = []
    offset = 0

    for run in runs:
        start = offset
        end = start + len(run.text)
        positions.append((run, start, end))
        offset = end

    for run, run_start, run_end in positions:
        ranges = []

        for word_start, word_end in unknown_words:
            start = max(word_start, run_start)
            end = min(word_end, run_end)

            if start < end:
                ranges.append((start - run_start, end - run_start))

        highlight_run(run, ranges)

    return len(unknown_words)


def get_paragraphs(container, seen_paragraphs, seen_tables):
    for paragraph in getattr(container, "paragraphs", []):
        if paragraph._p not in seen_paragraphs:
            seen_paragraphs.add(paragraph._p)
            yield paragraph

    for table in getattr(container, "tables", []):
        if table._tbl in seen_tables:
            continue

        seen_tables.add(table._tbl)

        for row in table.rows:
            for cell in row.cells:
                yield from get_paragraphs(
                    cell, seen_paragraphs, seen_tables
                )


def main():
    input_file = folder / "input.docx"
    output_file = folder / "output.docx"

    dictionary = load_dictionary()
    document = Document(input_file)

    seen_paragraphs = set()
    seen_tables = set()
    total_words = 0

    for paragraph in get_paragraphs(
        document, seen_paragraphs, seen_tables
    ):
        total_words += process_paragraph(paragraph, dictionary)

    document.save(output_file)

    print(f"Highlighted words: {total_words}")
    print(f"Output file: {output_file}")


if __name__ == "__main__":
    main()