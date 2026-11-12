#!/usr/bin/env python3
from __future__ import annotations

import argparse
import unicodedata
from pathlib import Path

from docx import Document
from docx.enum.table import WD_CELL_VERTICAL_ALIGNMENT, WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt


HEADER_FILL = "1F4E78"
BANDED_FILL = "EAF2F8"
BODY_FILL = "FFFFFF"
BORDER_COLOR = "D9D9D9"


def _display_width(text: str) -> int:
    return sum(2 if unicodedata.east_asian_width(char) in {"W", "F"} else 1 for char in text)


def _get_or_add(parent, tag: str):
    child = parent.find(qn(tag))
    if child is None:
        child = OxmlElement(tag)
        parent.append(child)
    return child


def _set_cell_border(cell, color: str = BORDER_COLOR, size: str = "6") -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    borders = _get_or_add(tc_pr, "w:tcBorders")
    for edge in ("top", "left", "bottom", "right", "insideH", "insideV"):
        node = _get_or_add(borders, f"w:{edge}")
        node.set(qn("w:val"), "single")
        node.set(qn("w:sz"), size)
        node.set(qn("w:space"), "0")
        node.set(qn("w:color"), color)


def _set_cell_shading(cell, fill: str) -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    shd = _get_or_add(tc_pr, "w:shd")
    shd.set(qn("w:val"), "clear")
    shd.set(qn("w:color"), "auto")
    shd.set(qn("w:fill"), fill)


def _set_cell_margins(cell, top=80, start=100, bottom=80, end=100) -> None:
    tc_pr = cell._tc.get_or_add_tcPr()
    tc_mar = _get_or_add(tc_pr, "w:tcMar")
    for name, value in (("top", top), ("start", start), ("bottom", bottom), ("end", end)):
        node = _get_or_add(tc_mar, f"w:{name}")
        node.set(qn("w:w"), str(value))
        node.set(qn("w:type"), "dxa")


def _set_repeat_header(row) -> None:
    tr_pr = row._tr.get_or_add_trPr()
    header = _get_or_add(tr_pr, "w:tblHeader")
    header.set(qn("w:val"), "true")


def _set_run_format(cell, *, size_half_points: int = 18, bold: bool | None = None, color: str | None = None) -> None:
    for run in cell._tc.xpath(".//w:r"):
        r_pr = run.get_or_add_rPr()
        size = _get_or_add(r_pr, "w:sz")
        size.set(qn("w:val"), str(size_half_points))
        size_cs = _get_or_add(r_pr, "w:szCs")
        size_cs.set(qn("w:val"), str(size_half_points))
        if bold is not None:
            bold_node = _get_or_add(r_pr, "w:b")
            bold_node.set(qn("w:val"), "true" if bold else "false")
            bold_cs = _get_or_add(r_pr, "w:bCs")
            bold_cs.set(qn("w:val"), "true" if bold else "false")
        if color is not None:
            color_node = _get_or_add(r_pr, "w:color")
            color_node.set(qn("w:val"), color)
            for attr in ("themeColor", "themeTint", "themeShade"):
                color_node.attrib.pop(qn("w:" + attr), None)


def _column_widths(table, available_twips: int) -> list[int]:
    column_count = len(table.columns)
    scores: list[int] = []
    for column_index in range(column_count):
        widths = [_display_width(row.cells[column_index].text.strip()) for row in table.rows]
        scores.append(max(8, min(48, max(widths, default=8))))

    total = sum(scores) or column_count
    minimum = min(600, available_twips // column_count)
    remainder = available_twips - minimum * column_count
    widths = [minimum + int(remainder * score / total) for score in scores]
    widths[-1] += available_twips - sum(widths)
    return widths


def _set_table_widths(table, available_twips: int) -> None:
    if not table.rows or not table.columns:
        return
    widths = _column_widths(table, available_twips)
    table.autofit = False

    tbl_pr = table._tbl.tblPr
    layout = _get_or_add(tbl_pr, "w:tblLayout")
    layout.set(qn("w:type"), "fixed")
    tbl_w = _get_or_add(tbl_pr, "w:tblW")
    tbl_w.set(qn("w:w"), str(available_twips))
    tbl_w.set(qn("w:type"), "dxa")

    grid = table._tbl.tblGrid
    if grid is not None:
        for grid_col, width in zip(grid.gridCol_lst, widths):
            grid_col.set(qn("w:w"), str(width))

    for row in table.rows:
        logical_column = 0
        for tc in row._tr.tc_lst:
            tc_pr = tc.get_or_add_tcPr()
            grid_span = tc_pr.find(qn("w:gridSpan"))
            span = int(grid_span.get(qn("w:val"))) if grid_span is not None else 1
            width = sum(widths[logical_column : logical_column + span])
            tc_w = _get_or_add(tc_pr, "w:tcW")
            tc_w.set(qn("w:w"), str(width))
            tc_w.set(qn("w:type"), "dxa")
            logical_column += span


def beautify_tables(input_path: str | Path, output_path: str | Path | None = None) -> int:
    source = Path(input_path)
    destination = Path(output_path) if output_path else source
    document = Document(source)
    if not document.tables:
        if destination != source:
            document.save(destination)
        return 0

    section = document.sections[0]
    available_twips = int((section.page_width - section.left_margin - section.right_margin) / 635)

    for table in document.tables:
        table.alignment = WD_TABLE_ALIGNMENT.CENTER
        _set_table_widths(table, available_twips)
        _set_repeat_header(table.rows[0])

        column_max_widths = [
            max((_display_width(row.cells[index].text.strip()) for row in table.rows), default=0)
            for index in range(len(table.columns))
        ]

        for row_index, row in enumerate(table.rows):
            for height in row._tr.xpath('./w:trPr/w:trHeight'):
                if height.get(qn('w:hRule')) == 'exact':
                    height.set(qn('w:hRule'), 'atLeast')
            for column_index, cell in enumerate(row.cells):
                _set_cell_border(cell)
                _set_cell_margins(cell)
                cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER

                if row_index == 0:
                    _set_cell_shading(cell, HEADER_FILL)
                    _set_run_format(cell, bold=True, color="FFFFFF")
                else:
                    _set_cell_shading(cell, BANDED_FILL if row_index % 2 == 0 else BODY_FILL)
                    _set_run_format(cell)

                alignment = (
                    WD_ALIGN_PARAGRAPH.CENTER
                    if column_max_widths[column_index] <= 14
                    else WD_ALIGN_PARAGRAPH.LEFT
                )
                for paragraph in cell.paragraphs:
                    paragraph.alignment = alignment
                    paragraph.paragraph_format.space_before = Pt(0)
                    paragraph.paragraph_format.space_after = Pt(0)
                    paragraph.paragraph_format.line_spacing = 1.0

    document.save(destination)
    return len(document.tables)


def main() -> int:
    parser = argparse.ArgumentParser(description="Apply consistent professional formatting to DOCX tables.")
    parser.add_argument("input", help="Input DOCX path")
    parser.add_argument("-o", "--output", help="Output DOCX path. Default: update input in place")
    args = parser.parse_args()
    count = beautify_tables(args.input, args.output)
    print(f"Beautified {count} table(s).")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
