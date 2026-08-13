#!/usr/bin/env python3
"""Extract stable paragraph IDs and visible text from a DOCX for review planning."""

from __future__ import annotations

import argparse
import hashlib
import json
import zipfile
from pathlib import Path

from lxml import etree

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
NS = {"w": W_NS}


def w(local: str) -> str:
    return f"{{{W_NS}}}{local}"


def xml_root(zf: zipfile.ZipFile, member: str) -> etree._Element:
    return etree.fromstring(zf.read(member))


def text_in(el: etree._Element, include_deleted: bool = False) -> str:
    parts: list[str] = []
    for node in el.iter():
        if node.tag == w("t") or (include_deleted and node.tag == w("delText")):
            parts.append(node.text or "")
        elif node.tag == w("tab"):
            parts.append("\t")
        elif node.tag in {w("br"), w("cr")}:
            parts.append("\n")
    return "".join(parts)


def paragraph_style(p: etree._Element) -> str:
    style = p.find("w:pPr/w:pStyle", namespaces=NS)
    return "" if style is None else style.get(w("val"), "")


def table_cell_position(p: etree._Element) -> tuple[int | None, int | None]:
    cell = p.getparent()
    while cell is not None and cell.tag != w("tc"):
        cell = cell.getparent()
    if cell is None:
        return None, None
    row = cell.getparent()
    table = row.getparent() if row is not None else None
    if row is None or table is None:
        return None, None
    cells = [x for x in row if x.tag == w("tc")]
    rows = [x for x in table if x.tag == w("tr")]
    return rows.index(row), cells.index(cell)


def list_paragraphs(root: etree._Element, part: str, start: int) -> tuple[list[dict], int]:
    records: list[dict] = []
    counter = start
    for p in root.xpath(".//w:p", namespaces=NS):
        visible = text_in(p)
        all_text = text_in(p, include_deleted=True)
        row, col = table_cell_position(p)
        records.append(
            {
                "paragraph_id": f"p{counter:06d}",
                "part": part,
                "text": visible,
                "text_with_deleted": all_text,
                "style": paragraph_style(p),
                "table_row": row,
                "table_col": col,
            }
        )
        counter += 1
    return records, counter


def extract_docx(path: Path) -> dict:
    with zipfile.ZipFile(path, "r") as zf:
        members = set(zf.namelist())
        # Review comments are deliberately limited to the main story part. Script
        # titles, scene headings, dialogue, and table-based layouts all live in
        # document.xml, while excluding headers/footers prevents accidental notes
        # on repeated page furniture.
        parts = ["word/document.xml"]
        paragraphs: list[dict] = []
        counter = 1
        for part in parts:
            if part not in members:
                continue
            found, counter = list_paragraphs(xml_root(zf, part), part, counter)
            paragraphs.extend(found)

        comments = []
        if "word/comments.xml" in members:
            croot = xml_root(zf, "word/comments.xml")
            for c in croot.xpath(".//w:comment", namespaces=NS):
                comments.append(
                    {
                        "id": c.get(w("id"), ""),
                        "author": c.get(w("author"), ""),
                        "text": text_in(c),
                    }
                )

    source_text = "\n".join(f"{x['part']}\t{x['paragraph_id']}\t{x['text']}" for x in paragraphs)
    source_sha256 = hashlib.sha256(source_text.encode("utf-8")).hexdigest()
    return {
        "source_file": str(path.resolve()),
        "source_sha256": source_sha256,
        "paragraph_count": len(paragraphs),
        "existing_comment_count": len(comments),
        "paragraphs": paragraphs,
        "existing_comments": comments,
    }


def main() -> None:
    ap = argparse.ArgumentParser(description="Extract reviewable paragraph map from DOCX")
    ap.add_argument("docx")
    ap.add_argument("--out", required=True)
    args = ap.parse_args()
    source = Path(args.docx)
    if not source.is_file():
        raise SystemExit(f"Input DOCX not found: {source}")
    result = extract_docx(source)
    out = Path(args.out)
    out.parent.mkdir(parents=True, exist_ok=True)
    out.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")
    print(
        f"[OK] extracted paragraphs={result['paragraph_count']} "
        f"existing_comments={result['existing_comment_count']} sha256={result['source_sha256']} -> {out}"
    )


if __name__ == "__main__":
    main()
