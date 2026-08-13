#!/usr/bin/env python3
"""Apply a validated screenplay review plan as true Word comments."""

from __future__ import annotations

import argparse
import datetime as dt
import hashlib
import json
import re
import zipfile
from pathlib import Path

from lxml import etree

from extract_docx import extract_docx

W_NS = "http://schemas.openxmlformats.org/wordprocessingml/2006/main"
PKG_REL_NS = "http://schemas.openxmlformats.org/package/2006/relationships"
CT_NS = "http://schemas.openxmlformats.org/package/2006/content-types"
NS = {"w": W_NS, "pr": PKG_REL_NS, "ct": CT_NS}

PRIORITIES = {"P0", "P1", "P2", "P3"}
CATEGORIES = {
    "剧情逻辑", "跨集因果", "单集发动", "卡点铺垫", "人物动机", "人物能动性",
    "人物关系", "情绪节奏", "信息释放", "设定一致性", "台词", "动作画面",
    "制作可行性", "连续性", "格式规范", "保护项",
}


def w(local: str) -> str:
    return f"{{{W_NS}}}{local}"


def pr(local: str) -> str:
    return f"{{{PKG_REL_NS}}}{local}"


def ct(local: str) -> str:
    return f"{{{CT_NS}}}{local}"


def xml_bytes(root: etree._Element) -> bytes:
    return etree.tostring(root, xml_declaration=True, encoding="UTF-8", standalone="yes")


def read_xml(zf: zipfile.ZipFile, member: str) -> etree._Element:
    return etree.fromstring(zf.read(member))


def paragraph_text(p: etree._Element) -> str:
    parts: list[str] = []
    for node in p.iter():
        if node.tag == w("t"):
            parts.append(node.text or "")
        elif node.tag == w("tab"):
            parts.append("\t")
        elif node.tag in {w("br"), w("cr")}:
            parts.append("\n")
    return "".join(parts)


def parse_plan(path: Path) -> dict:
    try:
        data = json.loads(path.read_text(encoding="utf-8-sig"))
    except Exception as exc:
        raise ValueError(f"Cannot parse review plan JSON: {exc}") from exc
    if not isinstance(data, dict):
        raise ValueError("Review plan must be a JSON object")
    if not isinstance(data.get("source_sha256"), str) or not data["source_sha256"]:
        raise ValueError("Review plan requires source_sha256")
    if not isinstance(data.get("comments"), list):
        raise ValueError("Review plan requires comments array")
    reviewer = data.get("reviewer", "AUCTRIX 剧本审稿")
    if not isinstance(reviewer, str) or not reviewer.strip():
        raise ValueError("reviewer must be a non-empty string")
    data["reviewer"] = reviewer.strip()
    return data


def validate_item(item: object, seen: set[str]) -> dict:
    if not isinstance(item, dict):
        raise ValueError("Each comment must be an object")
    required = {"paragraph_id", "anchor_quote", "priority", "category", "diagnosis", "suggestion"}
    missing = required - set(item)
    if missing:
        raise ValueError(f"Comment missing fields: {sorted(missing)}")
    values = {key: item[key].strip() if isinstance(item[key], str) else item[key] for key in required}
    if any(not isinstance(values[key], str) or not values[key] for key in required):
        raise ValueError("All comment fields must be non-empty strings")
    pid = values["paragraph_id"]
    if not re.fullmatch(r"p\d{6}", pid):
        raise ValueError(f"Invalid paragraph_id: {pid}")
    if pid in seen:
        raise ValueError(f"Only one new comment is allowed per paragraph: {pid}")
    seen.add(pid)
    if values["priority"] not in PRIORITIES:
        raise ValueError(f"Invalid priority at {pid}: {values['priority']}")
    if values["category"] not in CATEGORIES:
        raise ValueError(f"Invalid category at {pid}: {values['category']}")
    if len(values["anchor_quote"]) > 240:
        raise ValueError(f"anchor_quote is too long at {pid}")
    if len(values["diagnosis"]) > 600 or len(values["suggestion"]) > 600:
        raise ValueError(f"Comment is too long at {pid}; keep Word comments concise")
    return values


def ensure_comments_root(existing: bytes | None) -> etree._Element:
    if existing is not None:
        return etree.fromstring(existing)
    return etree.Element(w("comments"), nsmap={"w": W_NS})


def next_comment_id(doc_root: etree._Element, comments_root: etree._Element) -> int:
    used: set[int] = set()
    for el in doc_root.xpath(".//*[@w:id]", namespaces=NS):
        try:
            used.add(int(el.get(w("id"))))
        except (TypeError, ValueError):
            pass
    for el in comments_root.xpath(".//w:comment", namespaces=NS):
        try:
            used.add(int(el.get(w("id"))))
        except (TypeError, ValueError):
            pass
    candidate = 0
    while candidate in used:
        candidate += 1
    return candidate


def ensure_relationship(rels_root: etree._Element) -> None:
    rel_type = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments"
    for rel in rels_root.xpath(".//pr:Relationship", namespaces=NS):
        if rel.get("Type") == rel_type:
            return
    used: set[int] = set()
    for rel in rels_root.xpath(".//pr:Relationship", namespaces=NS):
        match = re.fullmatch(r"rId(\d+)", rel.get("Id", ""))
        if match:
            used.add(int(match.group(1)))
    candidate = 1
    while candidate in used:
        candidate += 1
    rel = etree.SubElement(rels_root, pr("Relationship"))
    rel.set("Id", f"rId{candidate}")
    rel.set("Type", rel_type)
    rel.set("Target", "comments.xml")


def ensure_content_type(ct_root: etree._Element) -> None:
    for override in ct_root.xpath(".//ct:Override", namespaces=NS):
        if override.get("PartName") == "/word/comments.xml":
            return
    override = etree.SubElement(ct_root, ct("Override"))
    override.set("PartName", "/word/comments.xml")
    override.set("ContentType", "application/vnd.openxmlformats-officedocument.wordprocessingml.comments+xml")


def add_anchor(p: etree._Element, comment_id: int) -> None:
    start = etree.Element(w("commentRangeStart"))
    start.set(w("id"), str(comment_id))
    insert_at = 1 if len(p) and p[0].tag == w("pPr") else 0
    p.insert(insert_at, start)

    end = etree.Element(w("commentRangeEnd"))
    end.set(w("id"), str(comment_id))
    p.append(end)
    run = etree.SubElement(p, w("r"))
    reference = etree.SubElement(run, w("commentReference"))
    reference.set(w("id"), str(comment_id))


def add_comment_body(root: etree._Element, comment_id: int, author: str, text: str) -> None:
    comment = etree.SubElement(root, w("comment"))
    comment.set(w("id"), str(comment_id))
    comment.set(w("author"), author)
    comment.set(w("initials"), "AX")
    comment.set(w("date"), dt.datetime.now(dt.timezone.utc).replace(microsecond=0).isoformat().replace("+00:00", "Z"))
    p = etree.SubElement(comment, w("p"))
    for i, line in enumerate(text.split("\n")):
        if i:
            break_run = etree.SubElement(p, w("r"))
            etree.SubElement(break_run, w("br"))
        run = etree.SubElement(p, w("r"))
        t = etree.SubElement(run, w("t"))
        if line[:1].isspace() or line[-1:].isspace():
            t.set("{http://www.w3.org/XML/1998/namespace}space", "preserve")
        t.text = line


def apply(source: Path, plan_path: Path, output: Path, manifest_path: Path) -> dict:
    plan = parse_plan(plan_path)
    source_map = extract_docx(source)
    if plan["source_sha256"] != source_map["source_sha256"]:
        raise ValueError("source_sha256 mismatch: the review plan targets a different DOCX version")

    seen: set[str] = set()
    items = [validate_item(x, seen) for x in plan["comments"]]
    map_by_id = {p["paragraph_id"]: p for p in source_map["paragraphs"]}

    with zipfile.ZipFile(source, "r") as zin:
        if zin.testzip() is not None:
            raise ValueError("Source DOCX ZIP is corrupt")
        names = set(zin.namelist())
        doc_root = read_xml(zin, "word/document.xml")
        paragraphs = doc_root.xpath(".//w:p", namespaces=NS)
        if len(paragraphs) != source_map["paragraph_count"]:
            raise ValueError("Paragraph enumeration changed unexpectedly")
        paragraph_by_id = {
            meta["paragraph_id"]: paragraph
            for meta, paragraph in zip(source_map["paragraphs"], paragraphs)
        }

        comments_bytes = zin.read("word/comments.xml") if "word/comments.xml" in names else None
        comments_root = ensure_comments_root(comments_bytes)
        rels_member = "word/_rels/document.xml.rels"
        if rels_member in names:
            rels_root = read_xml(zin, rels_member)
        else:
            rels_root = etree.Element(pr("Relationships"), nsmap={None: PKG_REL_NS})
        ct_root = read_xml(zin, "[Content_Types].xml")
        ensure_relationship(rels_root)
        ensure_content_type(ct_root)

        next_id = next_comment_id(doc_root, comments_root)
        manifest_items: list[dict] = []
        for item in items:
            meta = map_by_id.get(item["paragraph_id"])
            if meta is None or meta["part"] != "word/document.xml":
                raise ValueError(f"Unknown or unsupported paragraph_id: {item['paragraph_id']}")
            p = paragraph_by_id[item["paragraph_id"]]
            current_text = paragraph_text(p)
            quote = item["anchor_quote"]
            if current_text.count(quote) != 1:
                raise ValueError(
                    f"anchor_quote must match exactly once in {item['paragraph_id']}; "
                    f"found {current_text.count(quote)}"
                )
            comment_id = next_id
            next_id += 1
            body = f"【{item['priority']}｜{item['category']}】{item['diagnosis']}\n建议：{item['suggestion']}"
            add_anchor(p, comment_id)
            add_comment_body(comments_root, comment_id, plan["reviewer"], body)
            manifest_items.append(
                {
                    **item,
                    "comment_id": str(comment_id),
                    "comment_text": body,
                    "paragraph_text_sha256": hashlib.sha256(current_text.encode("utf-8")).hexdigest(),
                }
            )

        overrides = {
            "word/document.xml": xml_bytes(doc_root),
            "word/comments.xml": xml_bytes(comments_root),
            rels_member: xml_bytes(rels_root),
            "[Content_Types].xml": xml_bytes(ct_root),
        }
        output.parent.mkdir(parents=True, exist_ok=True)
        with zipfile.ZipFile(output, "w", zipfile.ZIP_DEFLATED) as zout:
            for info in zin.infolist():
                if info.filename in overrides:
                    zout.writestr(info, overrides[info.filename])
                else:
                    zout.writestr(info, zin.read(info.filename))
            for member in ("word/comments.xml", rels_member):
                if member not in names:
                    zout.writestr(member, overrides[member])

    output_map = extract_docx(output)
    if output_map["source_sha256"] != source_map["source_sha256"]:
        output.unlink(missing_ok=True)
        raise ValueError("Visible body text changed while adding comments; output discarded")

    manifest = {
        "source_file": str(source.resolve()),
        "output_file": str(output.resolve()),
        "source_sha256": source_map["source_sha256"],
        "output_sha256": output_map["source_sha256"],
        "existing_comment_count": source_map["existing_comment_count"],
        "new_comment_count": len(manifest_items),
        "reviewer": plan["reviewer"],
        "comments": manifest_items,
    }
    manifest_path.parent.mkdir(parents=True, exist_ok=True)
    manifest_path.write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding="utf-8")
    return manifest


def main() -> None:
    ap = argparse.ArgumentParser(description="Apply review-plan JSON as Word comments")
    ap.add_argument("docx")
    ap.add_argument("--plan", required=True)
    ap.add_argument("--out", required=True)
    ap.add_argument("--manifest", required=True)
    args = ap.parse_args()
    source = Path(args.docx)
    output = Path(args.out)
    if source.resolve() == output.resolve():
        raise SystemExit("Refusing to overwrite source DOCX; choose a different --out path")
    manifest = apply(source, Path(args.plan), output, Path(args.manifest))
    print(
        f"[OK] wrote {output} new_comments={manifest['new_comment_count']} "
        f"visible_text_unchanged=true manifest={args.manifest}"
    )


if __name__ == "__main__":
    main()
