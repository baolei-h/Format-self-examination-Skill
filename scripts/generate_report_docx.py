#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Generate the fixed Markdown and editable Word chapter-audit reports."""

from __future__ import annotations

import argparse
import json
import re
import sys
from collections import Counter
from datetime import datetime
from pathlib import Path
from typing import Any

from review_contract import review_completeness

try:
    from docx import Document
    from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn
    from docx.shared import Cm, Pt, RGBColor
except Exception:
    Document = None

VERDICTS = ("可进入下一轮编辑", "修改后可进入下一轮编辑", "暂不建议交稿")
SEVERITY_ORDER = {"P0": 0, "P1": 1, "P2": 2, "P3": 3}


def safe(value: Any, fallback: str = "未提供") -> str:
    if value is None or value == "":
        return fallback
    if isinstance(value, (dict, list)):
        return json.dumps(value, ensure_ascii=False)
    return str(value)


def issue_counts(issues: list[dict[str, Any]]) -> Counter:
    counts = Counter(item.get("severity", "P3") for item in issues)
    for severity in ("P0", "P1", "P2", "P3"):
        counts.setdefault(severity, 0)
    return counts


def combined_issues(audit: dict[str, Any], review: dict[str, Any]) -> list[dict[str, Any]]:
    issues = list(audit.get("issues", [])) + list(review.get("findings", []))
    normalized = []
    for item in issues:
        if not isinstance(item, dict):
            continue
        candidate = dict(item)
        if candidate.get("severity") not in SEVERITY_ORDER:
            candidate["severity"] = "P2"
        candidate.setdefault("category", "内容审核")
        candidate.setdefault("location", "全文")
        candidate.setdefault("message", "未提供问题说明")
        candidate.setdefault("rule", "按书稿规范核对")
        candidate.setdefault("suggested_fix", "由作者确认")
        candidate.setdefault("auto_fixable", False)
        candidate.setdefault("human_review", False)
        normalized.append(candidate)
    return sorted(normalized, key=lambda x: (SEVERITY_ORDER[x["severity"]], x["category"], x["location"]))


def verdict_for(issues: list[dict[str, Any]]) -> str:
    counts = issue_counts(issues)
    if counts["P0"] > 0 or counts["P1"] >= 10:
        return "暂不建议交稿"
    if counts["P1"] > 0:
        return "修改后可进入下一轮编辑"
    return "可进入下一轮编辑"


def check_status(category_terms: tuple[str, ...], issues: list[dict[str, Any]], manual_terms: tuple[str, ...] = ()) -> str:
    matching = [item for item in issues if any(term in item.get("category", "") for term in category_terms)]
    if any(not item.get("human_review") for item in matching):
        return "☐ 未通过"
    if any(item.get("human_review") for item in matching):
        return "△ 人工确认"
    if manual_terms and any(any(term in item.get("category", "") for term in manual_terms) for item in issues):
        return "△ 人工确认"
    return "☑ 未发现可测偏差"


def report_data(audit: dict[str, Any], review: dict[str, Any] | None = None) -> dict[str, Any]:
    review = review or {}
    meta = audit.get("meta", {})
    metrics = audit.get("metrics", {})
    issues = combined_issues(audit, review)
    counts = issue_counts(issues)
    review_sections = review.get("sections", {}) if isinstance(review.get("sections", {}), dict) else {}
    ref_stats = metrics.get("reference_stats", {})
    text_est = metrics.get("text_estimate", {})
    recent_topics = metrics.get("recent_reference_stats", [])
    review_complete, review_gaps = review_completeness(review)
    audit_hash = metrics.get("source_sha256")
    if review_complete and audit_hash and review.get("source_sha256") != audit_hash:
        review_complete = False
        review_gaps.append("review.source_sha256 与检查稿件哈希不一致或缺失")
    conclusion = verdict_for(issues)
    chapter_no = meta.get("chapter_no")
    chapter_label = f"第{chapter_no}章" if chapter_no else safe(meta.get("chapter_title"), "章节未识别")
    filename = Path(str(audit.get("file", "稿件"))).name
    report = {
        "title": "《生物活性表面可控构建技术》章节自查报告",
        "chapter": chapter_label,
        "filename": filename,
        "checked_at": safe(audit.get("audited_at"), datetime.now().isoformat(timespec="seconds")),
        "body_scale": safe(text_est.get("estimated_words_chars"), "未自动统计"),
        "first_level_count": safe(metrics.get("first_level_section_count"), "未识别"),
        "figure_count": safe(metrics.get("figure_caption_count", len(metrics.get("figure_captions", []))), "未自动测量"),
        "image_object_count": safe(metrics.get("image_object_count"), "未自动测量"),
        "table_count": safe(metrics.get("editable_word_tables"), "未自动测量"),
        "equation_count": safe(metrics.get("editable_equation_objects"), "未自动测量"),
        "reference_count": safe(ref_stats.get("total", len(metrics.get("references", []))), "未自动统计"),
        "known_reference_years": safe(ref_stats.get("year_identified"), "未自动统计"),
        "unknown_reference_years": safe(ref_stats.get("year_unknown"), "未自动统计"),
        "recent5": f"{safe(ref_stats.get('within_5_years'), '0')} / {safe(ref_stats.get('within_5_years_ratio'), '未计算')}",
        "recent10": f"{safe(ref_stats.get('within_10_years'), '0')} / {safe(ref_stats.get('within_10_years_ratio'), '未计算')}",
        "issue_counts": counts,
        "verdict": conclusion,
        "review_complete": review_complete,
        "review_gaps": review_gaps,
        "issues": issues,
        "review": review,
        "review_sections": review_sections,
        "metrics": metrics,
        "audit": audit,
        "references_stats": ref_stats,
        "recent_topics": recent_topics,
    }
    return report


def format_checks(report: dict[str, Any]) -> list[tuple[str, str, str]]:
    metrics = report["metrics"]
    issues = report["issues"]
    text_checks = metrics.get("format_checks", {})
    is_text = audit_input_type(report["audit"]) in {"md", "markdown", "txt"}
    checks = [
        ("页面尺寸与方向", ("页面版式", "页面方向"), ("页面方向人工确认",)),
        ("页边距与各 section", ("页面版式",), ()),
        ("字体与字号", ("字体与段落", "格式人工确认"), ("格式人工确认",)),
        ("标题层级与编号", ("章节结构",), ()),
        ("正文对齐、行距、缩进", ("字体与段落",), ()),
        ("目录字段、层级与页码", ("目录", "目录人工确认"), ("目录人工确认",)),
        ("页眉、页脚与页码", ("页码",), ("页码",)),
    ]
    result = []
    for label, failed_terms, manual_terms in checks:
        if is_text and label in {"页面尺寸与方向", "页边距与各 section", "字体与字号"}:
            status = "未自动测量"
        else:
            status = check_status(failed_terms, issues, manual_terms)
        note = safe(text_checks.get(label), "")
        if is_text and label in {"页面尺寸与方向", "页边距与各 section", "字体与字号"}:
            note = "Markdown/TXT 无法测量此项。"
        result.append((label, status, note))
    return result


def audit_input_type(audit: dict[str, Any]) -> str:
    return str(audit.get("metrics", {}).get("input_format", Path(str(audit.get("file", ""))).suffix.lower().lstrip("."))).lower()


def markdown_table(headers: list[str], rows: list[list[Any]]) -> str:
    def cell(value: Any) -> str:
        return safe(value, "").replace("|", "/").replace("\n", " ")
    output = ["| " + " | ".join(headers) + " |", "|" + "|".join(["---"] * len(headers)) + "|"]
    output.extend("| " + " | ".join(cell(item) for item in row) + " |" for row in rows)
    return "\n".join(output)


def render_markdown(report: dict[str, Any]) -> str:
    counts = report["issue_counts"]
    metrics = report["metrics"]
    review = report["review"]
    sections = report["review_sections"]
    lines = [f"# {report['title']}", "", "## 1. 总体结果", "",
             f"- 稿件名称：{report['filename']}", f"- 章节：{report['chapter']}",
             f"- 检查时间：{report['checked_at']}", f"- 正文规模估算：{report['body_scale']}",
             f"- 一级小节数量：{report['first_level_count']}", f"- 图题 / Word图片对象：{report['figure_count']} / {report['image_object_count']}",
             f"- 可编辑Word表格：{report['table_count']}", f"- 可编辑公式对象：{report['equation_count']}",
             f"- 参考文献：总数 {report['reference_count']}；可识别年份 {report['known_reference_years']}；年份未知 {report['unknown_reference_years']}",
             f"- 近5年文献：{report['recent5']}；近10年文献：{report['recent10']}",
             f"- 问题数：P0 {counts['P0']} / P1 {counts['P1']} / P2 {counts['P2']} / P3 {counts['P3']}",
             f"- 结论：**{report['verdict']}**",]
    if not report["review_complete"]:
        lines.append("- 状态：**确定性检查完成，AI全文语义审核尚未完成或覆盖不足；本报告是草稿，不可作为完整自查结论。**")
        lines.append(f"- AI审核待补项：{safe(report['review_gaps'])}")
    else:
        lines.append("- AI全文语义审核：已完成")
    lines.extend(["", "该结论仅代表本书内部撰写规范自查，不代表出版社最终审核结果。", "", "## 2. 必须修改的问题", ""])
    must_fix = [item for item in report["issues"] if item["severity"] in {"P0", "P1"}]
    rows = [[f"{index}", item["severity"], item["location"], item["message"], item["rule"], item["suggested_fix"], "是" if item.get("auto_fixable") else "否"] for index, item in enumerate(must_fix, start=1)]
    lines.append(markdown_table(["编号", "优先级", "位置", "问题", "规范要求", "修改建议", "可否自动修复"], rows) if rows else "未发现 P0/P1 问题。")
    lines.extend(["", "## 3. 格式检查结果", "", markdown_table(["检查项", "结果", "备注"], [[a, b, c] for a, b, c in format_checks(report)]), ""])
    lines.extend(["## 4. 图、表、公式和统计", ""])
    lines.append(f"- 图题 {report['figure_count']} 个；Word图片对象 {report['image_object_count']} 个；有效DPI已测量 {len(metrics.get('images', []))} 个对象。")
    lines.append(f"- 可编辑Word表格 {report['table_count']} 个；标题位置、表头单位和跨页重复表头需按 `table_layout` 与 Word 页面效果核验。")
    lines.append(f"- 可编辑公式对象 {report['equation_count']} 个；公式编号候选 {len(metrics.get('equation_labels', []))} 个。")
    stats = metrics.get("statistics", {})
    lines.append(f"- 统计报告启发式：{'检测到统计相关内容' if stats.get('detected') else '未识别到统计相关内容'}；统计检验是否适当始终需作者/统计人员确认。")
    for key, label in [("figure_visual_review", "AI视觉复核"), ("statistics_review", "统计语义复核")]:
        data = sections.get(key, {}) if isinstance(sections.get(key, {}), dict) else {}
        lines.append(f"- {label}：{safe(data.get('status'), '△需人工确认')}；{safe(data.get('summary'), '当前审核未提供结论。')}")
    for item in [x for x in report["issues"] if any(k in x["category"] for k in ("插图", "图题", "表格", "公式", "统计", "比例尺"))]:
        lines.append(f"- {item['severity']} {item['category']}（{item['location']}）：{item['message']}")
    lines.extend(["", "## 5. 参考文献", ""])
    lines.append(f"- 编号：正文缺失 {metrics.get('missing_reference_numbers', [])}；未被正文引用 {metrics.get('uncited_reference_numbers', [])}。")
    lines.append(f"- 年份分布：近5年 {report['recent5']}；近10年 {report['recent10']}；经典文献 {safe(report['references_stats'].get('classic'), '未统计')}；未知年份 {report['unknown_reference_years']}。")
    lines.append(f"- 疑似重复候选：{safe(metrics.get('suspected_duplicate_references', []), '无')}（需人工确认，不自动删除）。")
    lines.append(f"- 著录字段待核对：{safe(metrics.get('reference_metadata_gaps', []), '未检测到启发式缺项')}；DOI缺失不自动判错，也不自动编造。")
    for item in [x for x in report["issues"] if "参考文献" in x["category"] or "疑似重复" in x["category"]]:
        lines.append(f"- {item['severity']} {item['location']}：{item['message']}")
    lines.extend(["", "## 6. 内容和专著叙事", ""])
    for key, label in [("monograph_logic", "专著逻辑"), ("thesis_integration", "学位论文整合"), ("team_research_continuity", "团队研究连续性"), ("content_repetition", "内容重复"), ("mechanism_evidence", "机制证据"), ("clinical_extrapolation", "临床外推"), ("terminology_consistency", "术语与缩写一致性")]:
        data = sections.get(key, {}) if isinstance(sections.get(key, {}), dict) else {}
        lines.append(f"- {label}：{safe(data.get('status'), '△待AI审核')}；{safe(data.get('summary'), '需通读全文并结合上下文判断。')}")
    lines.extend(["", "## 7. 近5–10年进展", ""])
    if report["recent_topics"]:
        for item in report["recent_topics"]:
            lines.append(f"- `{item.get('title', '未命名专题')}`：专题引用 {item.get('reference_count', 0)} 条；有年份 {item.get('citation_n_with_year', 0)} 条；近5年 {item.get('within_5y', 0)}；近10年 {item.get('within_10y', 0)}。")
    else:
        recent_sections = metrics.get("recent_sections", [])
        lines.append("- 识别到的专题：" + ("、".join(str(x[3] if isinstance(x, (list, tuple)) else x.get("title")) for x in recent_sections) if recent_sections else "未识别到。"))
    recent_review = review.get("recent_progress", {}) if isinstance(review.get("recent_progress", {}), dict) else {}
    lines.append(f"- 主题覆盖与可能缺失：{safe(recent_review.get('summary'), '需结合章节主题及可核验的近年研究人工判断。')}")
    lines.append(f"- 文献类型比例：{safe(recent_review.get('review_types'), '未人工统计')}。未联网核验时不得编造文献或声称已覆盖最新研究。")
    lines.extend(["", "## 8. 版权、溯源、科研诚信", ""])
    integrity = sections.get("copyright_integrity", {}) if isinstance(sections.get("copyright_integrity", {}), dict) else {}
    lines.append(f"- 状态：{safe(integrity.get('status'), '△需作者人工确认')}。{safe(integrity.get('summary'), '核对图源、转载/改编许可、团队已发表图重用、比例尺和图像处理记录。')}")
    lines.append("- 本报告只列风险线索，不构成对版权侵权或科研不端的认定。")
    lines.extend(["", "## 9. 交稿包", ""])
    if metrics.get("package_provided"):
        inventory = metrics.get("package_inventory", {})
        for folder in ["01_Chapter", "02_Figures", "03_FigureSources", "04_Tables_Data", "05_References", "06_Permissions", "07_Checklist"]:
            item = inventory.get(folder, {})
            lines.append(f"- {'☑' if item.get('exists') else '☐'} `{folder}`：{'存在' if item.get('exists') else '缺失'}，文件 {item.get('file_count', 0)} 个。")
    else:
        lines.append("未提供完整交稿目录，无法核验01–07文件夹及源文件/权限材料。")
    lines.extend(["", "## 10. 推荐修改顺序", "", "1. 第一轮：先处理 P0 和 P1，修正文献映射、结构和确定性格式问题。", "2. 第二轮：重整专著叙事、重复内容和机制/临床证据链。", "3. 第三轮：统一 P2 单位、缩写、图表和元数据。", "4. 第四轮：P3 和出版社终版版式。", "", "## 11. 最终自查清单", ""])
    checklist = [
        "正文引文与文后参考文献逐条对应，并确认编号连续",
        "图、表、公式按章连续编号，正文先引用后出现",
        "显微图比例尺、图像源文件与图像处理记录已核验",
        "统计 n、重复类型、误差棒、检验和显著性说明齐全",
        "术语、缩写、单位、有效数字经作者核对",
        "近5–10年进展已由可核验来源复核，未杜撰文献或 DOI",
        "学位论文素材已重组为专著逻辑，团队研究脉络清楚",
        "第三方材料来源和许可记录已由作者/版权负责人确认",
        "复杂表格、公式、化学结构保留可编辑源文件",
        "交稿包目录、Word、图源、数据、文献库和清单齐全",
    ]
    lines.extend(f"- [ ] {item}" for item in checklist)
    lines.extend(["", "---", "本报告由确定性工具与AI/人工复核信息合并生成。自动检查不能替代作者科学审核、版权审核、统计专家复核或出版社终审。"])
    return "\n".join(lines) + "\n"


def set_cell_shading(cell, fill: str):
    tc_pr = cell._tc.get_or_add_tcPr()
    shading = OxmlElement("w:shd")
    shading.set(qn("w:fill"), fill)
    tc_pr.append(shading)


def set_run_font(run, east_asia: str = "宋体", latin: str = "Times New Roman", size: float = 10.5, bold: bool | None = None):
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


def set_cell_text(cell, value: Any, header: bool = False):
    cell.text = ""
    cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER
    paragraph = cell.paragraphs[0]
    paragraph.paragraph_format.space_after = Pt(0)
    run = paragraph.add_run(safe(value, ""))
    set_run_font(run, size=9, bold=header)
    if header:
        set_cell_shading(cell, "D9EAF7")


def add_docx_table(doc, headers: list[str], rows: list[list[Any]]):
    table = doc.add_table(rows=1, cols=len(headers))
    table.style = "Table Grid"
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    for index, title in enumerate(headers):
        set_cell_text(table.rows[0].cells[index], title, header=True)
    for row in rows:
        cells = table.add_row().cells
        for index, value in enumerate(row):
            set_cell_text(cells[index], value)
    doc.add_paragraph()
    return table


def render_docx(report: dict[str, Any], output_path: Path):
    if Document is None:
        raise RuntimeError("缺少 python-docx，无法生成Word自查报告")
    doc = Document()
    section = doc.sections[0]
    section.top_margin = Cm(2.2)
    section.bottom_margin = Cm(2.2)
    section.left_margin = Cm(2.3)
    section.right_margin = Cm(2.3)
    normal = doc.styles["Normal"]
    normal.font.name = "Times New Roman"
    normal.font.size = Pt(10.5)
    normal._element.rPr.rFonts.set(qn("w:eastAsia"), "宋体")
    for style_name, size in [("Title", 18), ("Heading 1", 14), ("Heading 2", 12)]:
        style = doc.styles[style_name]
        style.font.name = "Arial"
        style.font.size = Pt(size)
        style.font.bold = True
        style._element.rPr.rFonts.set(qn("w:eastAsia"), "黑体")

    title = doc.add_paragraph()
    title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = title.add_run(report["title"])
    set_run_font(run, east_asia="黑体", latin="Arial", size=18, bold=True)
    doc.add_paragraph()
    doc.add_heading("1. 总体结果", level=1)
    summary_rows = [
        ["稿件名称", report["filename"], "章节", report["chapter"]],
        ["检查时间", report["checked_at"], "正文规模估算", report["body_scale"]],
        ["一级小节数量", report["first_level_count"], "图题/图片对象", f"{report['figure_count']} / {report['image_object_count']}"],
        ["可编辑表格/公式", f"{report['table_count']} / {report['equation_count']}", "参考文献总数", report["reference_count"]],
        ["近5年文献", report["recent5"], "近10年文献", report["recent10"]],
        ["P0/P1/P2/P3", " / ".join(str(report["issue_counts"][x]) for x in ("P0", "P1", "P2", "P3")), "结论", report["verdict"]],
    ]
    add_docx_table(doc, ["项目", "结果", "项目", "结果"], summary_rows)
    if not report["review_complete"]:
        p = doc.add_paragraph("确定性检查完成，AI全文语义审核尚未完成或覆盖不足；本报告为草稿，不可作为完整自查结论。")
        for run in p.runs:
            run.font.color.rgb = RGBColor(192, 0, 0)
            run.bold = True
        doc.add_paragraph("AI审核待补项：" + safe(report["review_gaps"]))
    doc.add_paragraph("该结论仅代表本书内部撰写规范自查，不代表出版社最终审核结果。")

    doc.add_heading("2. 必须修改的问题", level=1)
    must_fix = [item for item in report["issues"] if item["severity"] in {"P0", "P1"}]
    issue_rows = [[str(index), item["severity"], item["location"], item["message"], item["rule"], item["suggested_fix"], "是" if item.get("auto_fixable") else "否"] for index, item in enumerate(must_fix, start=1)]
    add_docx_table(doc, ["编号", "优先级", "位置", "问题", "规范要求", "修改建议", "自动修复"], issue_rows or [["—", "—", "全文", "未发现 P0/P1 问题", "—", "继续完成人工复核", "否"]])

    doc.add_heading("3. 格式检查结果", level=1)
    add_docx_table(doc, ["检查项", "结果", "备注"], [[a, b, c] for a, b, c in format_checks(report)])

    doc.add_heading("4. 图、表、公式和统计", level=1)
    for text in [
        f"图题 {report['figure_count']} 个；Word图片对象 {report['image_object_count']} 个。图像比例尺、面板、坐标轴、标签、颜色、亮度/对比度、拉伸和科研图像完整性需视觉核验。",
        f"可编辑Word表格 {report['table_count']} 个；表题位置、三线表、表头单位、跨页重复表头和表注需核验。",
        f"可编辑公式对象 {report['equation_count']} 个；编号候选 {len(report['metrics'].get('equation_labels', []))} 个；公式截图与符号含义需核验。",
        f"统计启发式状态：{'检测到统计内容' if report['metrics'].get('statistics', {}).get('detected') else '未识别统计内容'}。统计方法是否适当需作者/统计人员判断。",
    ]:
        doc.add_paragraph(text, style="List Bullet")
    for key, label in [("figure_visual_review", "AI视觉复核"), ("statistics_review", "统计语义复核")]:
        data = report["review_sections"].get(key, {}) if isinstance(report["review_sections"].get(key, {}), dict) else {}
        doc.add_paragraph(f"{label}：{safe(data.get('status'), '△需人工确认')}；{safe(data.get('summary'), '当前审核未提供结论。')}", style="List Bullet")

    doc.add_heading("5. 参考文献", level=1)
    stats = report["references_stats"]
    ref_texts = [
        f"正文引用但文后缺失：{safe(report['metrics'].get('missing_reference_numbers'), '无')}；文后未被正文引用：{safe(report['metrics'].get('uncited_reference_numbers'), '无')}。",
        f"年份统计：可识别 {safe(stats.get('year_identified'), '0')}；未知 {safe(stats.get('year_unknown'), '0')}；近5年窗口 {safe(stats.get('window_5'), '未计算')}；近10年窗口 {safe(stats.get('window_10'), '未计算')}；经典文献 {safe(stats.get('classic'), '未计算')}。",
        f"疑似重复候选：{safe(report['metrics'].get('suspected_duplicate_references'), '无')}。候选需人工核实，不自动删除。",
        f"元数据待核候选：{safe(report['metrics'].get('reference_metadata_gaps'), '未检测到启发式缺项')}。旧文献缺 DOI 不判错，禁止编造 DOI。",
    ]
    for text in ref_texts:
        doc.add_paragraph(text, style="List Bullet")

    doc.add_heading("6. 内容和专著叙事", level=1)
    sections = report["review_sections"]
    for key, label in [("monograph_logic", "专著逻辑"), ("thesis_integration", "学位论文整合"), ("team_research_continuity", "团队研究连续性"), ("content_repetition", "内容重复"), ("mechanism_evidence", "机制证据"), ("clinical_extrapolation", "临床外推"), ("terminology_consistency", "术语与缩写一致性")]:
        data = sections.get(key, {}) if isinstance(sections.get(key, {}), dict) else {}
        doc.add_paragraph(f"{label}：{safe(data.get('status'), '△待AI审核')}；{safe(data.get('summary'), '需阅读全文并结合上下文判断。')}", style="List Bullet")

    doc.add_heading("7. 近5–10年进展", level=1)
    topics = report["recent_topics"]
    if topics:
        add_docx_table(doc, ["前沿专题", "专题引用", "可识别年份", "近5年", "近10年"], [[x.get("title", "未命名专题"), x.get("reference_count", 0), x.get("citation_n_with_year", 0), x.get("within_5y", 0), x.get("within_10y", 0)] for x in topics])
    recent_review = report["review"].get("recent_progress", {}) if isinstance(report["review"].get("recent_progress", {}), dict) else {}
    doc.add_paragraph(safe(recent_review.get("summary"), "需结合章节主题和可核验的近年研究判断。"))
    doc.add_paragraph("未联网核验时不得编造文献、DOI或声称已覆盖最新研究。")

    doc.add_heading("8. 版权、溯源、科研诚信", level=1)
    integrity = sections.get("copyright_integrity", {}) if isinstance(sections.get("copyright_integrity", {}), dict) else {}
    doc.add_paragraph(f"{safe(integrity.get('status'), '△需作者人工确认')}：{safe(integrity.get('summary'), '核对第三方图源、转载/改编许可、团队已发表图复用和图像处理记录。')}")
    doc.add_paragraph("本报告列出待核风险，不构成对版权侵权或科研不端的认定。")

    doc.add_heading("9. 交稿包", level=1)
    package = report["metrics"].get("package_inventory", {})
    if report["metrics"].get("package_provided"):
        add_docx_table(doc, ["目录", "状态", "文件数", "文件清单（节选）"], [[name, "存在" if data.get("exists") else "缺失", data.get("file_count", 0), ", ".join(data.get("files", [])[:8])] for name, data in package.items()])
    else:
        doc.add_paragraph("未提供交稿目录，无法核验01–07文件夹、图源、数据、文献库和许可材料。")

    doc.add_heading("10. 推荐修改顺序", level=1)
    for item in ["第一轮：处理P0/P1，修正文献映射、结构和基础格式。", "第二轮：重整专著叙事、重复内容和机制/临床证据链。", "第三轮：统一P2单位、缩写、图表和元数据。", "第四轮：处理P3和出版社终版版式。"]:
        doc.add_paragraph(item, style="List Number")

    doc.add_heading("11. 最终自查清单", level=1)
    for item in [
        "正文引文与文后参考文献逐条对应，并确认编号连续",
        "图、表、公式按章连续编号，正文先引用后出现",
        "显微图比例尺、图像源文件与图像处理记录已核验",
        "统计n、重复类型、误差棒、检验和显著性说明齐全",
        "术语、缩写、单位、有效数字经作者核对",
        "近5–10年进展已由可核验来源复核，未杜撰文献或DOI",
        "学位论文素材已重组为专著逻辑，团队研究脉络清楚",
        "第三方材料来源和许可记录已由作者/版权负责人确认",
        "复杂表格、公式、化学结构保留可编辑源文件",
        "交稿包目录、Word、图源、数据、文献库和清单齐全",
    ]:
        doc.add_paragraph("☐ " + item)
    doc.add_paragraph("自动检查不能替代作者科学审核、版权审核、统计专家复核或出版社终审。")
    output_path.parent.mkdir(parents=True, exist_ok=True)
    doc.save(output_path)
    return output_path


def main():
    parser = argparse.ArgumentParser(description="Generate a complete Markdown and Word self-examination report.")
    parser.add_argument("--audit-json", required=True, type=Path)
    parser.add_argument("--review-json", type=Path, help="AI semantic review JSON; when absent the report is clearly marked draft.")
    parser.add_argument("--output-md", required=True, type=Path)
    parser.add_argument("--output-docx", required=True, type=Path)
    args = parser.parse_args()
    try:
        audit = json.loads(args.audit_json.read_text(encoding="utf-8"))
        review = json.loads(args.review_json.read_text(encoding="utf-8")) if args.review_json else {}
        report = report_data(audit, review)
        markdown = render_markdown(report)
        args.output_md.parent.mkdir(parents=True, exist_ok=True)
        args.output_md.write_text(markdown, encoding="utf-8")
        try:
            render_docx(report, args.output_docx)
            print(f"已生成 Markdown 与 Word 报告：{args.output_md.name} / {args.output_docx.name}")
        except Exception as exc:
            print(f"Word报告生成失败（Markdown已保留）：{exc}", file=sys.stderr)
            return 3
    except Exception as exc:
        print(f"报告生成失败：{exc}", file=sys.stderr)
        return 2
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
