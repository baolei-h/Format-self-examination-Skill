from __future__ import annotations

from pathlib import Path

from audit_docx import Auditor
from audit_text import TextAuditor
from docx import Document
from docx.enum.section import WD_SECTION
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Inches, Pt


def categories(result):
    return {issue.category for issue in result.issues}


def test_compliant_docx_has_no_blocking_findings(documents):
    compliant, _ = documents
    result = Auditor(compliant, current_year=2026, expected_chapter=5).run()
    assert not [issue for issue in result.issues if issue.severity == "P0"]
    assert not [issue for issue in result.issues if issue.severity == "P1"], result.issues
    assert result.meta["chapter_no"] == 5
    assert result.metrics["first_level_section_count"] == 3
    assert result.metrics["figure_caption_count"] == 1
    assert result.metrics["image_object_count"] == 1
    assert result.metrics["editable_word_tables"] == 1
    assert result.metrics["page_number_field_detected"] is True
    assert result.metrics["reference_stats"]["within_5_years"] == 2  # 2022 and 2025
    assert result.metrics["reference_stats"]["within_10_years"] == 4  # 2017-2025, inclusive
    assert result.metrics["reference_stats"]["classic"] == 1  # 2016 is outside the inclusive 10-year window
    assert result.metrics["reference_stats"]["future_years"] == [{"reference": 6, "year": 2027}]
    assert "近年进展" not in categories(result)


def test_broken_docx_detects_required_regressions(documents):
    _, broken = documents
    result = Auditor(broken, current_year=2026, expected_chapter=5).run()
    found = categories(result)
    assert "页面版式" in found
    assert "字体与段落" in found
    assert "章节结构" in found
    assert "图表编号" in found
    assert "插图" in found
    assert "参考文献" in found
    assert "疑似重复文献" in found
    assert "量和单位" in found
    assert "近年进展" in found
    assert any(issue.severity == "P0" for issue in result.issues)
    assert result.metrics["first_level_section_count"] == 16
    assert result.metrics["missing_reference_numbers"] == [99]
    assert result.metrics["suspected_duplicate_references"]
    assert result.metrics["images"][0]["effective_dpi"] < 150
    assert result.metrics["unit_candidates"]["x_g_candidates"]
    assert len(result.metrics["unit_candidates"]["range_markers"]) > 1
    assert any(issue.category == "章节结构" and not issue.auto_fixable for issue in result.issues)
    assert any(issue.category == "图表编号" and not issue.auto_fixable for issue in result.issues)


def test_recent_year_window_boundaries_are_inclusive():
    assert Auditor.recent_year_bounds(2026, 5) == (2022, 2026)
    assert Auditor.recent_year_bounds(2026, 10) == (2017, 2026)


def test_all_word_sections_are_checked(documents, tmp_path):
    compliant, _ = documents
    doc = Document(str(compliant))
    second = doc.add_section(WD_SECTION.NEW_PAGE)
    second.page_width, second.page_height = Inches(8.5), Inches(11)
    second.top_margin = Inches(0.5)
    path = tmp_path / "two-sections.docx"
    doc.save(path)
    result = Auditor(path, current_year=2026, expected_chapter=5).run()
    assert len(result.metrics["sections"]) == 2
    assert any(issue.category == "页面版式" and issue.location.startswith("第2节") for issue in result.issues)
    assert result.metrics["sections"][1]["page_field"] is True


def test_page_number_alignment_uses_the_paragraph_containing_field(documents, tmp_path):
    compliant, _ = documents
    doc = Document(str(compliant))
    footer = doc.sections[0].footer
    first = footer.paragraphs[0]
    for node in list(first._p):
        first._p.remove(node)
    first.alignment = 0
    number = footer.add_paragraph()
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    number.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = number.add_run()
    begin = OxmlElement("w:fldChar")
    begin.set(qn("w:fldCharType"), "begin")
    instruction = OxmlElement("w:instrText")
    instruction.text = " PAGE "
    end = OxmlElement("w:fldChar")
    end.set(qn("w:fldCharType"), "end")
    run._r.extend([begin, instruction, end])
    path = tmp_path / "page-field-second-paragraph.docx"
    doc.save(path)
    result = Auditor(path, current_year=2026, expected_chapter=5).run()
    assert result.metrics["page_number_field_detected"] is True
    assert not any(issue.category == "页码" for issue in result.issues)


def test_style_inheritance_resolves_run_without_direct_font(documents, tmp_path):
    compliant, _ = documents
    doc = Document(str(compliant))
    normal = doc.styles["Normal"]
    normal.font.name = "Times New Roman"
    normal.font.size = Pt(12)
    rpr = normal._element.get_or_add_rPr()
    fonts = rpr.rFonts
    if fonts is None:
        fonts = OxmlElement("w:rFonts")
        rpr.insert(0, fonts)
    fonts.set(qn("w:eastAsia"), "宋体")
    body_run = doc.paragraphs[1].runs[0]
    if body_run._element.rPr is not None:
        body_run._element.remove(body_run._element.rPr)
    path = tmp_path / "inherited-style.docx"
    doc.save(path)
    result = Auditor(path, current_year=2026, expected_chapter=5).run()
    assert result.metrics["formatting_unresolved_count"] == 0
    assert not any(issue.category == "字体与段落" and "正文" in issue.location for issue in result.issues)


def test_markdown_audit_marks_layout_as_unmeasured(tmp_path):
    chapter = tmp_path / "Ch05.md"
    chapter.write_text(
        "# 第5章 测试\n\n5.1 原理\n\n正文引文[1]。\n\n5.2 个性化发展方向\n\n近期研究[1]。\n\n参考文献\n[1] Synthetic Author. Synthetic title[J]. Test Journal, 2025, 1(1): 1-2.\n",
        encoding="utf-8",
    )
    result = TextAuditor(chapter, current_year=2026, expected_chapter=5).run().as_json()
    assert result["metrics"]["format_checks"]["page_size"] == "未自动测量"
    assert result["metrics"]["reference_stats"]["within_5_years"] == 1
    assert not result["metrics"]["missing_reference_numbers"]


def test_markdown_audit_reports_duplicate_references_units_and_package(tmp_path):
    chapter = tmp_path / "Ch05.md"
    chapter.write_text(
        "# 第5章 测试\n\n5.1 原理\n\n条件10mm，转速300 rpm，n=3。\n\n5.2 个性化发展方向\n\n近期研究[1-2]。\n\n参考文献\n"
        "[1] Synthetic Author. Same synthetic title[J]. Test Journal, 2025, 1(1): 1-2. DOI: 10.0000/test-only\n"
        "[2] Synthetic Author. Same synthetic title[J]. Test Journal, 2025, 1(1): 1-2. DOI: 10.0000/test-only\n",
        encoding="utf-8",
    )
    for name in ["01_Chapter", "02_Figures", "03_FigureSources", "04_Tables_Data", "05_References", "06_Permissions", "07_Checklist"]:
        (tmp_path / name).mkdir()
    (tmp_path / "02_Figures" / "figure-source.txt").write_text("synthetic fixture", encoding="utf-8")
    result = TextAuditor(chapter, current_year=2026, package_dir=tmp_path, expected_chapter=5).run().as_json()
    found = {issue["category"] for issue in result["issues"]}
    assert "疑似重复文献" in found
    assert "量和单位" in found
    assert result["metrics"]["package_inventory"]["02_Figures"]["file_count"] == 1
