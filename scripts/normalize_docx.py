#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Safely normalize measurable DOCX formatting into a new file.

The source is never opened for writing. The output is created only if it does not
already exist; document body text and embedded image bytes are compared before and after the OOXML round trip.
Scientific data, claims, references, DOI values, labels, and image bytes are not edited.
"""

from __future__ import annotations

import argparse
import hashlib
import re
import sys
from pathlib import Path

try:
    from docx import Document
    from docx.enum.style import WD_STYLE_TYPE
    from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
    from docx.enum.section import WD_ORIENT
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    from docx.shared import Cm, Mm, Pt
except Exception:
    Document = None

CHAPTER_RE = re.compile(r"^第\s*\d+\s*章")
HEADING_RE = re.compile(r"^(\d+(?:\.\d+){1,6})\s+.+$")
FIG_RE = re.compile(r"^图\s*\d+\s*[-－—]\s*\d+")
TABLE_RE = re.compile(r"^表\s*\d+\s*[-－—]\s*\d+")
REF_RE = re.compile(r"^\[\d+\]\s*")


def sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def body_text_signature(doc) -> tuple[str, ...]:
    return tuple(node.text or "" for node in doc.element.body.xpath(".//w:t"))


def image_media_signature(doc) -> tuple[tuple[str, str], ...]:
    """Hash embedded image parts so safe formatting can prove pixels were preserved."""
    items = []
    for part in doc.part.package.parts:
        partname = str(getattr(part, "partname", ""))
        content_type = str(getattr(part, "content_type", ""))
        if partname.startswith("/word/media/") or content_type.startswith("image/"):
            items.append((partname, hashlib.sha256(part.blob).hexdigest()))
    return tuple(sorted(items))


def set_run_fonts(run, east_asia: str, latin: str, size: float, bold: bool | None = None):
    run.font.name = latin
    run.font.size = Pt(size)
    if bold is not None:
        run.bold = bold
    rpr = run._element.get_or_add_rPr()
    rfonts = rpr.rFonts
    if rfonts is None:
        rfonts = OxmlElement("w:rFonts")
        rpr.insert(0, rfonts)
    rfonts.set(qn("w:eastAsia"), east_asia)
    rfonts.set(qn("w:ascii"), latin)
    rfonts.set(qn("w:hAnsi"), latin)


def add_page_field(paragraph):
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = paragraph.add_run()
    begin = OxmlElement("w:fldChar")
    begin.set(qn("w:fldCharType"), "begin")
    instruction = OxmlElement("w:instrText")
    instruction.set(qn("xml:space"), "preserve")
    instruction.text = " PAGE "
    separate = OxmlElement("w:fldChar")
    separate.set(qn("w:fldCharType"), "separate")
    cached = OxmlElement("w:t")
    cached.text = "1"
    end = OxmlElement("w:fldChar")
    end.set(qn("w:fldCharType"), "end")
    run._r.append(begin)
    run._r.append(instruction)
    run._r.append(separate)
    run._r.append(cached)
    run._r.append(end)
    set_run_fonts(run, "宋体", "Times New Roman", 10.5)


def has_page_field(footer) -> bool:
    try:
        instructions = " ".join(node.text or "" for node in footer._element.xpath(".//w:instrText"))
        return bool(re.search(r"(?<![A-Z])PAGE(?![A-Z])", instructions, re.I))
    except Exception:
        return False


def normalize_document(source: Path, output: Path) -> dict:
    if Document is None:
        raise RuntimeError("缺少 python-docx；请安装 requirements.txt 中的依赖。")
    source = Path(source).resolve()
    output = Path(output).resolve()
    if not source.is_file() or source.suffix.lower() != ".docx":
        raise ValueError("输入必须是存在的 .docx 文件。")
    if source == output:
        raise ValueError("输出路径不能覆盖原稿。")
    if output.exists():
        raise FileExistsError(f"修订版已存在，为保护既有文件，不覆盖：{output}")
    output.parent.mkdir(parents=True, exist_ok=True)
    source_hash_before = sha256(source)
    doc = Document(str(source))
    content_before = body_text_signature(doc)
    image_signature_before = image_media_signature(doc)
    fixes = []

    # A4 dimensions and house-standard margins. Keep an existing landscape section
    # landscape because it may contain an intentionally wide table.
    for section in doc.sections:
        landscape = section.page_width > section.page_height or section.orientation == WD_ORIENT.LANDSCAPE
        section.orientation = WD_ORIENT.LANDSCAPE if landscape else WD_ORIENT.PORTRAIT
        if landscape:
            section.page_width, section.page_height = Mm(297), Mm(210)
        else:
            section.page_width, section.page_height = Mm(210), Mm(297)
        section.top_margin = Mm(25)
        section.bottom_margin = Mm(25)
        section.left_margin = Mm(26)
        section.right_margin = Mm(24)
    fixes.append("各 section 统一为 A4 与规定页边距（保留现有横向 section 方向）")

    style_targets = {
        "Normal": (12, "宋体", "Times New Roman", False),
        "Heading 1": (14, "黑体", "Arial", True),
        "Heading 2": (12, "黑体", "Arial", True),
        "Heading 3": (10.5, "黑体", "Arial", True),
        "Caption": (10.5, "宋体", "Times New Roman", False),
    }
    for name, (size, east_asia, latin, bold) in style_targets.items():
        try:
            style = doc.styles[name]
        except KeyError:
            style = doc.styles.add_style(name, WD_STYLE_TYPE.PARAGRAPH)
        style.font.name = latin
        style.font.size = Pt(size)
        style.font.bold = bold
        rpr = style._element.get_or_add_rPr()
        rfonts = rpr.rFonts
        if rfonts is None:
            rfonts = OxmlElement("w:rFonts")
            rpr.insert(0, rfonts)
        rfonts.set(qn("w:eastAsia"), east_asia)
        rfonts.set(qn("w:ascii"), latin)
        rfonts.set(qn("w:hAnsi"), latin)

    in_references = False
    for paragraph in doc.paragraphs:
        text = paragraph.text.strip()
        if not text:
            continue
        if re.fullmatch(r"参考文献\s*[:：]?", text):
            in_references = True
            paragraph.style = doc.styles["Heading 1"]
            paragraph.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT
            paragraph.paragraph_format.first_line_indent = Cm(0)
            for run in paragraph.runs:
                set_run_fonts(run, "黑体", "Arial", 14, True)
            continue
        if CHAPTER_RE.match(text):
            paragraph.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
            paragraph.paragraph_format.space_before = Pt(24)
            paragraph.paragraph_format.space_after = Pt(18)
            paragraph.paragraph_format.first_line_indent = Cm(0)
            for run in paragraph.runs:
                set_run_fonts(run, "黑体", "Arial", 18, True)
        elif (match := HEADING_RE.match(text)):
            level = match.group(1).count(".")
            target_style = doc.styles[f"Heading {min(level, 3)}"]
            paragraph.style = target_style
            paragraph.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.LEFT
            paragraph.paragraph_format.first_line_indent = Cm(0)
            paragraph.paragraph_format.space_before = Pt({1: 18, 2: 12, 3: 6}.get(level, 6))
            paragraph.paragraph_format.space_after = Pt({1: 6, 2: 4, 3: 3}.get(level, 3))
            for run in paragraph.runs:
                set_run_fonts(run, "黑体", "Arial", {1: 14, 2: 12, 3: 10.5}.get(level, 10.5), True)
        elif FIG_RE.match(text) or TABLE_RE.match(text):
            paragraph.style = doc.styles["Caption"]
            paragraph.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
            paragraph.paragraph_format.first_line_indent = Cm(0)
            paragraph.paragraph_format.line_spacing = 1.0
            for run in paragraph.runs:
                set_run_fonts(run, "宋体", "Times New Roman", 10.5)
        elif in_references and REF_RE.match(text):
            paragraph.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            paragraph.paragraph_format.left_indent = Cm(0.74)
            paragraph.paragraph_format.first_line_indent = Cm(-0.74)
            paragraph.paragraph_format.line_spacing = 1.0
            paragraph.paragraph_format.space_before = Pt(0)
            paragraph.paragraph_format.space_after = Pt(2)
            for run in paragraph.runs:
                set_run_fonts(run, "宋体", "Times New Roman", 10.5)
        elif paragraph.style.name.startswith("List") or paragraph.style.name.startswith("Heading"):
            # Keep list indents and unnumbered semantic headings as authored.
            continue
        else:
            paragraph.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
            paragraph.paragraph_format.line_spacing = 1.5
            paragraph.paragraph_format.first_line_indent = Cm(0.74)
            paragraph.paragraph_format.space_before = Pt(0)
            paragraph.paragraph_format.space_after = Pt(0)
            for run in paragraph.runs:
                set_run_fonts(run, "宋体", "Times New Roman", 12)
    fixes.append("正文、章/节标题、图表题和参考文献的对齐、行距、缩进及字体字号已规范化")

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for paragraph in cell.paragraphs:
                    paragraph.paragraph_format.line_spacing = 1.0
                    paragraph.paragraph_format.space_before = Pt(0)
                    paragraph.paragraph_format.space_after = Pt(0)
                    for run in paragraph.runs:
                        set_run_fonts(run, "宋体", "Times New Roman", 10.5)
    if doc.tables:
        fixes.append("表格文字字体和字号统一为宋体/Times New Roman 10.5 pt，保留表格数据与边框")

    has_any_page = False
    for section in doc.sections:
        footer = section.footer
        if has_page_field(footer):
            has_any_page = True
            for paragraph in footer.paragraphs:
                instructions = " ".join(node.text or "" for node in paragraph._p.xpath(".//w:instrText"))
                if re.search(r"(?<![A-Z])PAGE(?![A-Z])", instructions, re.I):
                    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
                    break
            continue
        if section.footer.is_linked_to_previous:
            continue
        paragraph = footer.paragraphs[0] if footer.paragraphs else footer.add_paragraph()
        add_page_field(paragraph)
        has_any_page = True
    if not has_any_page and doc.sections:
        # A linked footer chain may have no concrete footer part. Add the field at
        # the first unlinked footer and let subsequent sections inherit it.
        section = doc.sections[0]
        paragraph = section.footer.paragraphs[0] if section.footer.paragraphs else section.footer.add_paragraph()
        add_page_field(paragraph)
    fixes.append("页脚页码字段已添加/居中；字段结果由 Word 打开后更新")

    temporary = output.with_name(output.stem + ".normalizing.tmp.docx")
    if temporary.exists():
        raise FileExistsError(f"临时输出文件已存在，为保护既有文件，不覆盖：{temporary}")
    try:
        doc.save(temporary)
        saved = Document(str(temporary))
        if body_text_signature(saved) != content_before:
            raise RuntimeError("格式修复后正文 OOXML 文本签名改变，已拒绝发布修订版。")
        if image_media_signature(saved) != image_signature_before:
            raise RuntimeError("格式修复后嵌入图片内容发生改变，已拒绝发布修订版。")
        if sha256(source) != source_hash_before:
            raise RuntimeError("原稿 SHA256 发生变化；已中止。")
        if output.exists():
            raise FileExistsError(f"输出文件在处理期间出现，未覆盖：{output}")
        temporary.rename(output)
    except Exception:
        if temporary.exists():
            temporary.unlink()
        raise
    return {"source": str(source), "output": str(output), "source_sha256": source_hash_before,
            "source_sha256_unchanged": sha256(source) == source_hash_before,
            "body_text_unchanged": body_text_signature(Document(str(output))) == content_before,
            "image_pixels_unchanged": image_media_signature(Document(str(output))) == image_signature_before,
            "fixes": fixes}


def main():
    parser = argparse.ArgumentParser(description="Create a safe formatted copy of a DOCX; never overwrite source or output.")
    parser.add_argument("input", type=Path)
    parser.add_argument("--output", required=True, type=Path)
    args = parser.parse_args()
    try:
        result = normalize_document(args.input, args.output)
        print("安全格式修复完成；原稿未覆盖，SHA256保持不变。")
        print("输出：" + result["output"])
        print("修复项目：" + "；".join(result["fixes"]))
    except Exception as exc:
        print(f"安全格式修复未完成：{exc}", file=sys.stderr)
        return 2
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
