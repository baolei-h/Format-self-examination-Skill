"""Create synthetic DOCX fixtures at test time; no manuscript binaries are committed."""

from __future__ import annotations

from pathlib import Path

from docx import Document
from docx.enum.section import WD_SECTION
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Mm, Pt
from PIL import Image


def set_run(run, *, cjk="宋体", latin="Times New Roman", size=12, bold=False):
    run.font.name = latin
    run.font.size = Pt(size)
    run.bold = bold
    rpr = run._element.get_or_add_rPr()
    fonts = rpr.rFonts
    if fonts is None:
        fonts = OxmlElement("w:rFonts")
        rpr.insert(0, fonts)
    fonts.set(qn("w:eastAsia"), cjk)
    fonts.set(qn("w:ascii"), latin)
    fonts.set(qn("w:hAnsi"), latin)


def add_page_field(paragraph):
    paragraph.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = paragraph.add_run()
    begin = OxmlElement("w:fldChar")
    begin.set(qn("w:fldCharType"), "begin")
    instruction = OxmlElement("w:instrText")
    instruction.set(qn("xml:space"), "preserve")
    instruction.text = " PAGE "
    end = OxmlElement("w:fldChar")
    end.set(qn("w:fldCharType"), "end")
    run._r.extend([begin, instruction, end])


def add_paragraph(doc, text, *, kind="body", size=12, cjk="宋体", latin="Times New Roman", bold=False):
    p = doc.add_paragraph()
    run = p.add_run(text)
    set_run(run, cjk=cjk, latin=latin, size=size, bold=bold)
    if kind == "chapter":
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.first_line_indent = Mm(0)
    elif kind == "heading":
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p.paragraph_format.first_line_indent = Mm(0)
    elif kind == "caption":
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.paragraph_format.first_line_indent = Mm(0)
    else:
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        p.paragraph_format.first_line_indent = Mm(7.4)
        p.paragraph_format.line_spacing = 1.5
        p.paragraph_format.space_before = Pt(0)
        p.paragraph_format.space_after = Pt(0)
    return p


def create_test_documents(directory: Path) -> tuple[Path, Path]:
    directory.mkdir(parents=True, exist_ok=True)
    image_path = directory / "synthetic-test-image.png"
    Image.new("RGB", (1800, 1200), color=(180, 200, 220)).save(image_path, dpi=(300, 300))

    good = Document()
    section = good.sections[0]
    section.page_width, section.page_height = Mm(210), Mm(297)
    section.top_margin, section.bottom_margin = Mm(25), Mm(25)
    section.left_margin, section.right_margin = Mm(26), Mm(24)
    add_page_field(section.footer.paragraphs[0])
    add_paragraph(good, "第5章 骨与硬组织植入材料", kind="chapter", size=18, cjk="黑体", latin="Arial", bold=True)
    add_paragraph(good, "本章按照科学问题、构建原理、代表性研究和应用组织内容。正文引用[1-6]并先于图表出现。", size=12)
    add_paragraph(good, "5.1 表面构建原理", kind="heading", size=14, cjk="黑体", latin="Arial", bold=True)
    add_paragraph(good, "每组 n=3，包含独立生物学重复和技术重复。数据以 mean ± SD 表示，图注说明误差棒为 SD；统计检验采用 ANOVA。P < 0.05；* 表示 P < 0.05。", size=12)
    add_paragraph(good, "见图5-1，图像为合成测试色块，不代表实验结果；表5-1列出合成测试字段。", size=12)
    good.add_picture(str(image_path), width=Inches(6), height=Inches(4))
    add_paragraph(good, "图5-1 合成测试图（非实验数据）", kind="caption", size=10.5)
    add_paragraph(good, "表5-1 合成测试表（非实验数据）", kind="caption", size=10.5)
    table = good.add_table(rows=2, cols=2)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    table.cell(0, 0).text, table.cell(0, 1).text = "测试变量", "单位"
    table.cell(1, 0).text, table.cell(1, 1).text = "合成参数", "mm"
    add_paragraph(good, "5.2 构建方法", kind="heading", size=14, cjk="黑体", latin="Arial", bold=True)
    add_paragraph(good, "方法部分只用于测试格式检查，10 mm 和 300 μm 的单位间距符合规范。", size=12)
    add_paragraph(good, "5.3 个性化骨修复体与骨再生界面的发展方向", kind="heading", size=14, cjk="黑体", latin="Arial", bold=True)
    add_paragraph(good, "近期专题引用[1-6]。文献用于检查年份窗口，不表示真实出版记录。", size=12)
    add_paragraph(good, "参考文献", kind="heading", size=14, cjk="黑体", latin="Arial", bold=True)
    for text in [
        "[1] Synthetic Author. Synthetic test title one[J]. Test Journal, 2025, 2(1): 1-2. DOI: 10.0000/test-fixture-one",
        "[2] Synthetic Author. Synthetic test title two[J]. Test Journal, 2022, 2(1): 3-4. DOI: 10.0000/test-fixture-two",
        "[3] Synthetic Author. Synthetic test title three[J]. Test Journal, 2017, 1(1): 5-6. DOI: 10.0000/test-fixture-three",
        "[4] Synthetic Author. Synthetic test title four[J]. Test Journal, 2021, 1(1): 7-8. DOI: 10.0000/test-fixture-four",
        "[5] Synthetic Author. Synthetic test title five[J]. Test Journal, 2016, 1(1): 9-10. DOI: 10.0000/test-fixture-five",
        "[6] Synthetic Author. Synthetic test title six[J]. Test Journal, 2027, 3(1): 11-12. DOI: 10.0000/test-fixture-six",
    ]:
        p = good.add_paragraph()
        p.alignment = WD_ALIGN_PARAGRAPH.JUSTIFY
        p.paragraph_format.left_indent = Mm(7.4)
        p.paragraph_format.first_line_indent = Mm(-7.4)
        p.paragraph_format.line_spacing = 1.0
        p.paragraph_format.space_after = Pt(2)
        set_run(p.add_run(text), size=10.5)
    compliant = directory / "compliant.docx"
    good.save(compliant)

    bad = Document()
    bad_section = bad.sections[0]
    bad_section.page_width, bad_section.page_height = Inches(8.5), Inches(11)
    bad_section.top_margin, bad_section.bottom_margin = Inches(0.5), Inches(1)
    bad_section.left_margin, bad_section.right_margin = Inches(1), Inches(1)
    add_paragraph(bad, "第5章 故意不合规的合成测试稿", kind="chapter", size=10, cjk="仿宋", latin="Calibri")
    add_paragraph(bad, "正文先引用图5-2，但引用[1]、[3]和缺失的[99]。同时包含单位10mm、转速300 rpm、离心条件1000xg和混用范围3–5 mm与6~8 mm。", size=9)
    for index in range(1, 17):
        add_paragraph(bad, f"5.{index} 故意设置的一级测试小节{index}", kind="heading", size=9, cjk="宋体", latin="Calibri")
    add_paragraph(bad, "图5-2 低分辨率合成测试图（非实验数据）", kind="caption", size=9)
    bad_image_path = directory / "screenshot-test.png"
    Image.new("RGB", (300, 200), color=(80, 80, 80)).save(bad_image_path)
    bad.add_picture(str(bad_image_path), width=Inches(6), height=Inches(4))
    add_paragraph(bad, "参考文献", kind="heading", size=9, cjk="宋体", latin="Calibri")
    for text in [
        "[1] Synthetic Author. Duplicate synthetic title[J]. Test Journal, 2021, 1(1): 1-2. DOI: 10.0000/test-only-duplicate",
        "[3] Synthetic Author. Duplicate synthetic title[J]. Test Journal, 2021, 1(1): 1-2. DOI: 10.0000/test-only-duplicate",
    ]:
        p = bad.add_paragraph(text)
        p.alignment = WD_ALIGN_PARAGRAPH.LEFT
        p.paragraph_format.line_spacing = 1.5
        for run in p.runs:
            set_run(run, cjk="仿宋", latin="Calibri", size=9)
    broken = directory / "broken.docx"
    bad.save(broken)
    return compliant, broken
