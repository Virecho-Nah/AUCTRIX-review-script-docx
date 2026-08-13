#!/usr/bin/env python3
"""Create a minimal DOCX fixture and verify the full comment round trip."""

from __future__ import annotations

import json
import subprocess
import sys
import tempfile
import zipfile
from pathlib import Path


CONTENT_TYPES = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
  <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
  <Default Extension="xml" ContentType="application/xml"/>
  <Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>
</Types>"""

ROOT_RELS = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
  <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>
</Relationships>"""

DOCUMENT = """<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p><w:r><w:t>第1集</w:t></w:r></w:p>
    <w:p><w:r><w:t>林澈没有任何犹豫，立刻把唯一的证据交给刚刚背叛她的人。</w:t></w:r></w:p>
    <w:sectPr/>
  </w:body>
</w:document>"""


def run(command: list[str]) -> None:
    subprocess.run(command, check=True)


def main() -> None:
    scripts = Path(__file__).resolve().parent
    with tempfile.TemporaryDirectory(prefix="auctrix_review_test_") as temp_dir:
        temp = Path(temp_dir)
        source = temp / "fixture.docx"
        mapping = temp / "map.json"
        plan = temp / "plan.json"
        output = temp / "reviewed.docx"
        manifest = temp / "manifest.json"

        with zipfile.ZipFile(source, "w", zipfile.ZIP_DEFLATED) as archive:
            archive.writestr("[Content_Types].xml", CONTENT_TYPES)
            archive.writestr("_rels/.rels", ROOT_RELS)
            archive.writestr("word/document.xml", DOCUMENT)

        run([sys.executable, str(scripts / "extract_docx.py"), str(source), "--out", str(mapping)])
        mapped = json.loads(mapping.read_text(encoding="utf-8"))
        target = next(item for item in mapped["paragraphs"] if "唯一的证据" in item["text"])
        plan.write_text(
            json.dumps(
                {
                    "reviewer": "AUCTRIX 剧本审稿",
                    "source_sha256": mapped["source_sha256"],
                    "comments": [
                        {
                            "paragraph_id": target["paragraph_id"],
                            "anchor_quote": "把唯一的证据交给刚刚背叛她的人",
                            "priority": "P1",
                            "category": "人物动机",
                            "diagnosis": "这次关键选择没有可见触发或交换条件，人物像在替剧情递交证据。",
                            "suggestion": "补出迫使她改选的代价，或让她先提出能保护自身目标的交换条件。",
                        }
                    ],
                },
                ensure_ascii=False,
                indent=2,
            ),
            encoding="utf-8",
        )

        run(
            [
                sys.executable,
                str(scripts / "apply_review_comments.py"),
                str(source),
                "--plan",
                str(plan),
                "--out",
                str(output),
                "--manifest",
                str(manifest),
            ]
        )
        run([sys.executable, str(scripts / "audit_review_comments.py"), str(output), "--manifest", str(manifest)])
        result = json.loads(manifest.read_text(encoding="utf-8"))
        if result["new_comment_count"] != 1 or result["source_sha256"] != result["output_sha256"]:
            raise SystemExit("Round-trip assertions failed")
        print("[OK] AUCTRIX review DOCX round trip passed")


if __name__ == "__main__":
    main()
