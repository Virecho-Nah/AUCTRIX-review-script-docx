#!/usr/bin/env python3
"""Audit Word-comment wiring, anchors, bodies, and visible-text integrity."""

from __future__ import annotations

import argparse
import hashlib
import json
import zipfile
from pathlib import Path

from lxml import etree

from extract_docx import extract_docx

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
NS = {"w": W_NS, "pr": PKG_REL_NS, "ct": CT_NS}


def w(local: str) -> str:
    return f"{{{W_NS}}}{local}"


def node_text(el: etree._Element) -> str:
    parts: list[str] = []
    for node in el.iter():
        if node.tag == w("t"):
            parts.append(node.text or "")
        elif node.tag == w("tab"):
            parts.append("\t")
        elif node.tag in {w("br"), w("cr")}:
            parts.append("\n")
    return "".join(parts)


def paragraph_text(el: etree._Element) -> str:
    parts: list[str] = []
    for node in el.iter():
        if node.tag == w("t"):
            parts.append(node.text or "")
        elif node.tag == w("tab"):
            parts.append("\t")
        elif node.tag in {w("br"), w("cr")}:
            parts.append("\n")
    return "".join(parts)


def fail(errors: list[str], message: str) -> None:
    errors.append(message)


def audit(docx: Path, manifest_path: Path) -> dict:
    manifest = json.loads(manifest_path.read_text(encoding="utf-8-sig"))
    errors: list[str] = []
    output_map = extract_docx(docx)
    if output_map["source_sha256"] != manifest.get("source_sha256"):
        fail(errors, "Visible document text fingerprint differs from source")
    if manifest.get("output_sha256") != output_map["source_sha256"]:
        fail(errors, "Manifest output fingerprint does not match DOCX")

    with zipfile.ZipFile(docx, "r") as zf:
        corrupt = zf.testzip()
        if corrupt:
            fail(errors, f"Corrupt ZIP member: {corrupt}")
        names = set(zf.namelist())
        required = {"word/document.xml", "word/comments.xml", "word/_rels/document.xml.rels", "[Content_Types].xml"}
        for member in required - names:
            fail(errors, f"Missing required OOXML part: {member}")
        if errors:
            return {"ok": False, "errors": errors}

        doc_root = etree.fromstring(zf.read("word/document.xml"))
        comments_root = etree.fromstring(zf.read("word/comments.xml"))
        rels_root = etree.fromstring(zf.read("word/_rels/document.xml.rels"))
        ct_root = etree.fromstring(zf.read("[Content_Types].xml"))
        paragraphs = doc_root.xpath(".//w:p", namespaces=NS)
        paragraph_by_id = {
            meta["paragraph_id"]: paragraph
            for meta, paragraph in zip(output_map["paragraphs"], paragraphs)
        }

        rel_type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments"
        if not any(x.get("Type") == rel_type for x in rels_root.xpath(".//pr:Relationship", namespaces=NS)):
            fail(errors, "document.xml.rels lacks comments relationship")
        if not any(x.get("PartName") == "/word/comments.xml" for x in ct_root.xpath(".//ct:Override", namespaces=NS)):
            fail(errors, "[Content_Types].xml lacks comments override")

        all_comments = comments_root.xpath(".//w:comment", namespaces=NS)
        expected_total = manifest.get("existing_comment_count", 0) + manifest.get("new_comment_count", 0)
        if len(all_comments) != expected_total:
            fail(errors, f"Comment count mismatch: expected {expected_total}, found {len(all_comments)}")

        comment_by_id = {x.get(w("id")): x for x in all_comments}
        for item in manifest.get("comments", []):
            cid = item["comment_id"]
            comment = comment_by_id.get(cid)
            if comment is None:
                fail(errors, f"Missing comment body id={cid}")
                continue
            if comment.get(w("author"), "") != manifest.get("reviewer"):
                fail(errors, f"Wrong author for comment id={cid}")
            if node_text(comment) != item["comment_text"]:
                fail(errors, f"Comment text mismatch id={cid}")

            starts = doc_root.xpath(f".//w:commentRangeStart[@w:id='{cid}']", namespaces=NS)
            ends = doc_root.xpath(f".//w:commentRangeEnd[@w:id='{cid}']", namespaces=NS)
            refs = doc_root.xpath(f".//w:commentReference[@w:id='{cid}']", namespaces=NS)
            if (len(starts), len(ends), len(refs)) != (1, 1, 1):
                fail(errors, f"Anchor wiring mismatch id={cid}: start/end/ref={len(starts)}/{len(ends)}/{len(refs)}")
                continue

            target = paragraph_by_id.get(item["paragraph_id"])
            if target is None:
                fail(errors, f"Target paragraph missing for id={cid}")
                continue
            if not all(target is node.getparent() or target in node.iterancestors() for node in (starts[0], ends[0], refs[0])):
                fail(errors, f"Comment id={cid} is not anchored in {item['paragraph_id']}")
            target_text = paragraph_text(target)
            if hashlib.sha256(target_text.encode("utf-8")).hexdigest() != item["paragraph_text_sha256"]:
                fail(errors, f"Target paragraph text changed for comment id={cid}")
            if target_text.count(item["anchor_quote"]) != 1:
                fail(errors, f"Anchor quote no longer unique for comment id={cid}")

    return {
        "ok": not errors,
        "errors": errors,
        "new_comment_count": manifest.get("new_comment_count", 0),
        "visible_text_unchanged": output_map["source_sha256"] == manifest.get("source_sha256"),
    }


def main() -> None:
    ap = argparse.ArgumentParser(description="Audit a reviewed DOCX and its comment manifest")
    ap.add_argument("docx")
    ap.add_argument("--manifest", required=True)
    args = ap.parse_args()
    result = audit(Path(args.docx), Path(args.manifest))
    if not result["ok"]:
        for error in result["errors"]:
            print(f"[ERROR] {error}")
        raise SystemExit(2)
    print(
        f"[OK] audit passed new_comments={result['new_comment_count']} "
        f"visible_text_unchanged={str(result['visible_text_unchanged']).lower()}"
    )


if __name__ == "__main__":
    main()
