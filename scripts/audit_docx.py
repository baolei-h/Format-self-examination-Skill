#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Deterministic audit helper for 《生物活性表面可控构建技术》 chapter DOCX files.

This script performs measurable checks only. It intentionally does not judge scientific
correctness, copyright ownership, or evidence strength; the calling agent must perform
those semantic checks using the bundled audit rubric.
"""

from __future__ import annotations

import argparse
import difflib
import json
import math
import os
import re
import sys
import unicodedata
from collections import Counter, defaultdict
from dataclasses import dataclass, asdict
from datetime import datetime
from io import BytesIO
from pathlib import Path
from typing import Any, Iterable

try:
    from docx import Document
    from docx.enum.section import WD_ORIENT
    from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_LINE_SPACING
    from docx.oxml.ns import qn
except Exception as e:  # pragma: no cover
    print("缺少运行依赖 python-docx。请由调用环境安装仓库 requirements.txt 中的依赖后重试。", file=sys.stderr)
    sys.exit(2)

try:
    from PIL import Image
except Exception:
    Image = None

EMU_PER_INCH = 914400
MM_PER_INCH = 25.4
PT_PER_INCH = 72.0

CHAPTER_RE = re.compile(r"^第\s*(\d+)\s*章\s*(.+)?$")
H_RE = re.compile(r"^(\d+(?:\.\d+){1,6})\s+(.+)$")
FIG_CAP_RE = re.compile(r"^图\s*(\d+)\s*[-－—]\s*(\d+)\s*(.*)$")
TAB_CAP_RE = re.compile(r"^表\s*(\d+)\s*[-－—]\s*(\d+)\s*(.*)$")
REF_ENTRY_RE = re.compile(r"^\[(\d+)\]\s*(.+)$")
CIT_GROUP_RE = re.compile(r"\[\s*\d+(?:\s*[-–—,，]\s*\d+)*\s*\]")
YEAR_RE = re.compile(r"\b(19\d{2}|20\d{2})\b")
CHINESE_RE = re.compile(r"[\u4e00-\u9fff]")
LATIN_WORD_RE = re.compile(r"[A-Za-z]+(?:[-'][A-Za-z]+)*")
NUMBER_TOKEN_RE = re.compile(r"\d+(?:\.\d+)?")
RECENT_TITLE_RE = re.compile(r"(最新进展|研究进展|发展方向|前沿|新技术|先进制造|个性化|精准|智能|未来发展|发展趋势|趋势展望|发展展望)", re.I)
DOI_RE = re.compile(r"\b10\.\d{4,9}/[^\s<>\]\[{}]+", re.I)
STRONG_CLAIM_RE = re.compile(r"首次实现|首次发现|首次提出|证实|证明|决定性|直接调控|完全|彻底解决|显著优于|临床有效|实现临床应用|巨大临床价值")
THESIS_STACKING_RE = re.compile(r"某某研究|某某同学|某同学|另一项研究|另一研究|某研究团队")

A4_W_MM = 210.0
A4_H_MM = 297.0
TARGET_MARGINS = {"top": 25.0, "bottom": 25.0, "left": 26.0, "right": 24.0}

SEV_ORDER = {"P0": 0, "P1": 1, "P2": 2, "P3": 3}


@dataclass
class Issue:
    severity: str
    category: str
    location: str
    message: str
    rule: str
    suggested_fix: str
    auto_fixable: bool = False
    human_review: bool = False


class Auditor:
    def __init__(self, path: Path, current_year: int | None = None, package_dir: Path | None = None, expected_chapter: int | None = None):
        self.path = path
        self.current_year = current_year if current_year is not None else datetime.now().year
        self.package_dir = package_dir
        self.expected_chapter = expected_chapter
        self.doc = Document(str(path))
        self.paragraphs = list(self.doc.paragraphs)
        self.issues: list[Issue] = []
        self.meta: dict[str, Any] = {}
        self.metrics: dict[str, Any] = {}
        self._para_class: dict[int, str] = {}
        self._default_rfonts = self.document_default_fonts()
        self._default_size = self.document_default_size()

    @staticmethod
    def recent_year_bounds(current_year: int, years: int) -> tuple[int, int]:
        """Return an inclusive calendar-year window (e.g. 2022-2026 for 5 years)."""
        return current_year - years + 1, current_year

    def document_default_fonts(self) -> dict[str, str | None]:
        out = {"eastAsia": None, "ascii": None, "hAnsi": None}
        try:
            nodes = self.doc.styles.element.xpath(".//w:docDefaults/w:rPrDefault/w:rPr/w:rFonts")
            if nodes:
                for key in out:
                    out[key] = nodes[0].get(qn("w:" + key))
        except Exception:
            pass
        return out

    def document_default_size(self) -> float | None:
        try:
            nodes = self.doc.styles.element.xpath(".//w:docDefaults/w:rPrDefault/w:rPr/w:sz")
            if nodes and nodes[0].get(qn("w:val")):
                return int(nodes[0].get(qn("w:val"))) / 2
        except Exception:
            pass
        return None

    def all_text(self, include_tables: bool = True) -> str:
        parts = [p.text for p in self.paragraphs]
        if include_tables:
            seen = {id(p._p) for p in self.paragraphs}
            for table in self.doc.tables:
                for row in table.rows:
                    for cell in row.cells:
                        for paragraph in cell.paragraphs:
                            if id(paragraph._p) not in seen:
                                seen.add(id(paragraph._p))
                                parts.append(paragraph.text)
        return "\n".join(parts)

    def add(self, severity: str, category: str, location: str, message: str, rule: str, suggested_fix: str, auto=False, human=False):
        self.issues.append(Issue(severity, category, location, message, rule, suggested_fix, auto, human))

    @staticmethod
    def mm(length) -> float | None:
        if length is None:
            return None
        return float(length.inches) * MM_PER_INCH

    @staticmethod
    def pt(length) -> float | None:
        if length is None:
            return None
        return float(length.pt)

    @staticmethod
    def approx(a: float | None, b: float, tol: float) -> bool:
        return a is not None and abs(a - b) <= tol

    @staticmethod
    def text_estimate(text: str) -> dict[str, int]:
        # Planning-oriented estimate, not publisher's typesetting character count.
        zh = len(CHINESE_RE.findall(text))
        latin = len(LATIN_WORD_RE.findall(text))
        nums = len(NUMBER_TOKEN_RE.findall(text))
        return {"chinese_chars": zh, "latin_words": latin, "number_tokens": nums, "estimated_words_chars": zh + latin + nums}

    def style_chain(self, style):
        seen = set()
        while style is not None and getattr(style, "style_id", None) not in seen:
            seen.add(getattr(style, "style_id", None))
            yield style
            style = getattr(style, "base_style", None)

    def pformat_value(self, p, attr: str):
        v = getattr(p.paragraph_format, attr, None)
        if v is not None:
            return v
        for st in self.style_chain(p.style):
            v = getattr(st.paragraph_format, attr, None)
            if v is not None:
                return v
        return None

    def effective_alignment(self, p):
        a = p.alignment
        if a is not None:
            return a
        for st in self.style_chain(p.style):
            a = st.paragraph_format.alignment
            if a is not None:
                return a
        return None

    def rfonts_from_element(self, el) -> dict[str, str | None]:
        out = {"eastAsia": None, "ascii": None, "hAnsi": None}
        try:
            rpr = el.rPr
            if rpr is None or rpr.rFonts is None:
                return out
            for k in out:
                out[k] = rpr.rFonts.get(qn("w:" + k))
        except Exception:
            pass
        return out

    def effective_run_fonts(self, run, p_style=None) -> dict[str, str | None]:
        out = self.rfonts_from_element(run._element)
        if run.font.name:
            out["ascii"] = out["ascii"] or run.font.name
            out["hAnsi"] = out["hAnsi"] or run.font.name
        if p_style is not None:
            for st in self.style_chain(p_style):
                rf = self.rfonts_from_element(st._element)
                for k in out:
                    out[k] = out[k] or rf[k]
                if st.font.name:
                    out["ascii"] = out["ascii"] or st.font.name
                    out["hAnsi"] = out["hAnsi"] or st.font.name
        for key in out:
            out[key] = out[key] or self._default_rfonts.get(key)
        return out

    def effective_run_size(self, run, p_style=None) -> float | None:
        if run.font.size is not None:
            return run.font.size.pt
        if p_style is not None:
            for st in self.style_chain(p_style):
                if st.font.size is not None:
                    return st.font.size.pt
        return self._default_size

    def effective_run_bold(self, run, p_style=None) -> bool | None:
        if run.bold is not None:
            return bool(run.bold)
        if p_style is not None:
            for st in self.style_chain(p_style):
                if st.font.bold is not None:
                    return bool(st.font.bold)
        return None

    def paragraph_representative_run(self, p):
        for r in p.runs:
            if r.text.strip():
                return r
        return None

    def classify_paragraphs(self):
        chapter_no = None
        chapter_title = None
        chapter_titles = []
        headings = []
        fig_caps = []
        tab_caps = []
        refs_heading_idx = None
        refs = []
        for i, p in enumerate(self.doc.paragraphs):
            t = p.text.strip()
            if not t:
                self._para_class[i] = "empty"
                continue
            m = CHAPTER_RE.match(t)
            if m:
                self._para_class[i] = "chapter"
                chapter_titles.append((i, int(m.group(1)), t))
                if chapter_no is None:
                    chapter_no = int(m.group(1))
                    chapter_title = t
                continue
            if re.fullmatch(r"参考文献\s*[:：]?", t):
                self._para_class[i] = "refs_heading"
                if refs_heading_idx is None:
                    refs_heading_idx = i
                continue
            m = REF_ENTRY_RE.match(t)
            if m:
                self._para_class[i] = "ref_entry"
                refs.append((i, int(m.group(1)), m.group(2)))
                continue
            m = FIG_CAP_RE.match(t)
            if m:
                self._para_class[i] = "fig_caption"
                fig_caps.append((i, int(m.group(1)), int(m.group(2)), t))
                continue
            m = TAB_CAP_RE.match(t)
            if m:
                self._para_class[i] = "table_caption"
                tab_caps.append((i, int(m.group(1)), int(m.group(2)), t))
                continue
            m = H_RE.match(t)
            if m:
                level = m.group(1).count(".")
                self._para_class[i] = f"heading_{level}"
                headings.append((i, m.group(1), level, m.group(2), t))
                continue
            self._para_class[i] = "body"
        self.meta.update({
            "chapter_no": chapter_no,
            "chapter_title": chapter_title,
            "chapter_titles": chapter_titles,
            "refs_heading_idx": refs_heading_idx,
            "current_year": self.current_year,
        })
        self.metrics.update({
            "headings": headings,
            "figure_captions": fig_caps,
            "table_captions": tab_caps,
            "references": refs,
        })

    def audit_page(self):
        section_metrics = []
        footer_page_fields = []
        for idx, sec in enumerate(self.doc.sections, start=1):
            w, h = self.mm(sec.page_width), self.mm(sec.page_height)
            loc = f"第{idx}节(Word section)"
            portrait = h is not None and w is not None and h >= w
            a4_page = w is not None and h is not None and self.approx(min(w, h), A4_W_MM, 1.5) and self.approx(max(w, h), A4_H_MM, 1.5)
            if not a4_page:
                dims = f"{w:.1f} x {h:.1f} mm" if w is not None and h is not None else "尺寸无法解析"
                self.add("P1", "页面版式", loc, f"页面约为 {dims}，非统一A4作者交稿版。", "A4 210 x 297 mm", "改为A4；若出版社模板另有要求则按出版社模板。", True)
            if not portrait:
                self.add("P2", "页面方向人工确认", loc, "检测到横向 section；标准默认纵向，超宽表可单独横向。", "默认纵向，仅超宽表单独横向", "确认横向范围仅覆盖确有需要的宽表，不误判为格式错误。", human=True)
            for key, val in TARGET_MARGINS.items():
                got = self.mm(getattr(sec, f"{key}_margin"))
                if got is None or not self.approx(got, val, 1.2):
                    got_text = f"{got:.1f} mm" if got is not None else "未设置"
                    self.add("P1", "页面版式", loc, f"{key}边距约 {got_text}，目标 {val:.0f} mm。", "上25/下25/左26/右24 mm", "统一章节页边距。", True)

            footer = sec.footer
            instructions = []
            try:
                instructions = [node.text or "" for node in footer._element.xpath(".//w:instrText")]
            except Exception:
                pass
            has_page_field = any(re.search(r"(?<![A-Z])PAGE(?![A-Z])", item, re.I) for item in instructions)
            footer_page_fields.append(has_page_field)
            footer_paragraphs = list(footer.paragraphs)
            field_alignment = None
            if has_page_field:
                for footer_paragraph in footer_paragraphs:
                    paragraph_instructions = [node.text or "" for node in footer_paragraph._p.xpath(".//w:instrText")]
                    if any(re.search(r"(?<![A-Z])PAGE(?![A-Z])", item, re.I) for item in paragraph_instructions):
                        field_alignment = str(footer_paragraph.alignment)
                        break
            section_metrics.append({
                "section": idx,
                "page_mm": [round(w, 2) if w is not None else None, round(h, 2) if h is not None else None],
                "orientation": "portrait" if portrait else "landscape_or_invalid",
                "header_text": " ".join(p.text.strip() for p in sec.header.paragraphs if p.text.strip()),
                "footer_text": " ".join(p.text.strip() for p in footer_paragraphs if p.text.strip()),
                "page_field": has_page_field,
                "footer_alignment": field_alignment,
            })
        self.metrics["sections"] = section_metrics
        self.metrics["page_number_field_detected"] = bool(footer_page_fields) and all(footer_page_fields)
        if not footer_page_fields or not all(footer_page_fields):
            self.add("P2", "页码", "页脚", "至少一个 Word section 的页脚未检测到 PAGE 自动页码字段；手工页码或出版社后排版请人工确认。", "页码置于页脚居中", "确认每个 section 均有正确页码字段；本项无法仅凭文本可靠判断人工页码。", human=True)
        elif any(item.get("footer_alignment") != str(WD_ALIGN_PARAGRAPH.CENTER) for item in section_metrics):
            self.add("P2", "页码", "页脚", "检测到页码字段，但至少一个 section 页脚段落未居中。", "页码置于页脚居中", "将页码段落居中并检查不同 section 的页脚链接。", True)

    def audit_structure(self):
        chapter_no = self.meta.get("chapter_no")
        headings = self.metrics["headings"]
        if chapter_no is None:
            self.add("P1", "章节结构", "全文", "未识别到“第X章 ……”章标题。", "章标题格式：第1章……", "确认章标题文字与编号格式。")
            return
        chapter_titles = self.meta.get("chapter_titles", [])
        mismatched_chapters = [item for item in chapter_titles if item[1] != chapter_no]
        if mismatched_chapters:
            self.add("P1", "章节结构", mismatched_chapters[0][2], "文档中出现多个不同章号。", "单章文件只包含一个章号", "核对是否误合并了其他章节。")
        if self.expected_chapter is not None and chapter_no != self.expected_chapter:
            self.add("P1", "章节结构", self.meta.get("chapter_title", "全文"), f"文件识别为第{chapter_no}章，与本次请求的第{self.expected_chapter}章不一致。", "检查请求章节与稿件一致", "确认选择了正确版本的章节文件。")
        # First-level sections are numbers like 1.1 => level == 1
        h1s = [x for x in headings if x[2] == 1]
        self.metrics["first_level_section_count"] = len(h1s)
        if len(h1s) > 15:
            self.add("P1", "章节结构", "全文", f"一级小节共 {len(h1s)} 个，超过15个。", "每个大章一级小节不超过15个", "合并相近小节，保留主线。")
        too_deep = [x for x in headings if x[2] > 3]  # 1.1.1.1 is the deepest supported section level.
        if too_deep:
            self.add("P1", "章节结构", too_deep[0][4], "检测到超过四级标题的编号。", "标题层级原则上不超过四级", "上提或合并标题层级。")
        wrong_chapter = [x for x in headings if int(x[1].split('.')[0]) != chapter_no]
        if wrong_chapter:
            ex = "; ".join(x[4] for x in wrong_chapter[:3])
            self.add("P1", "章节结构", ex, "节编号与章号不一致。", "节编号首位应与章号一致", "统一章节编号并更新交叉引用。")

        # Check each heading level for skipped/duplicate numbers and skipped levels.
        sequences: dict[tuple[int, ...], list[tuple[int, int, str]]] = defaultdict(list)
        last_level = 0
        level_jumps = []
        for paragraph_index, number, level, _, title in headings:
            parts = tuple(int(x) for x in number.split("."))
            parent = parts[:-1]
            sequences[parent].append((paragraph_index, parts[-1], title))
            if last_level and level > last_level + 1:
                level_jumps.append(title)
            if last_level == 0 and level > 1:
                level_jumps.append(title)
            last_level = level
        for parent, items in sequences.items():
            actual = [item[1] for item in items]
            expected = list(range(1, max(actual) + 1)) if actual else []
            if actual != expected:
                prefix = ".".join(str(x) for x in parent) or "一级小节"
                label = f"{prefix}下级" if parent else "一级小节"
                self.add("P1", "章节结构", label, f"编号序列为 {actual[:30]}{'...' if len(actual)>30 else ''}，存在缺号、重号或未从1开始。", "同层级标题从1开始连续编号", "核对标题增删后的层级编号。")
        if level_jumps:
            self.add("P1", "章节结构", "; ".join(level_jumps[:5]), "标题层级发生跳级。", "标题层级不得跳过直接下钻", "在缺失的上级标题下组织内容或调整层级。")

        # Length estimate excluding references.
        ref_idx = self.meta.get("refs_heading_idx")
        paras = self.doc.paragraphs[:ref_idx] if ref_idx is not None else self.doc.paragraphs
        text = "\n".join(p.text for p in paras)
        est = self.text_estimate(text)
        self.metrics["text_estimate"] = est
        n = est["estimated_words_chars"]
        if chapter_no == 1:
            if n < 5000 or n > 10000:
                self.add("P2", "章节规模", "全文", f"估算正文规模约 {n:,}（中文字符+英文词+数字token），规划目标5,000-10,000。", "第1章约5,000-10,000字", "根据概论与导航功能增删内容；最终字数以编辑部口径为准。")
        elif 2 <= chapter_no <= 10:
            if n < 45000 or n > 65000:
                self.add("P2", "章节规模", "全文", f"估算正文规模约 {n:,}，规划目标约50,000-60,000。", "第2-10章约5-6万字", "按章节策划目标调整；这是规划性要求，不作为机械退稿标准。")

        # Recent section presence for chapters 2-10. The topic is not limited to one fixed title.
        if 2 <= chapter_no <= 10:
            recent = [x for x in headings if RECENT_TITLE_RE.search(x[3])]
            self.metrics["recent_sections"] = recent
            if not recent:
                self.add("P1", "近年进展", "全文", "未识别到专门的近5-10年进展/发展方向一级小节。", "第2-10章均设置前沿专题", "增加专门一级小节，并以近5-10年高质量研究为主体更新。")

    def audit_formatting(self):
        # Expected: class -> (size, alignment, bold, cjk, latin)
        expectations = {
            "chapter": (18.0, WD_ALIGN_PARAGRAPH.CENTER, True, "黑体", "Arial"),
            "heading_1": (14.0, WD_ALIGN_PARAGRAPH.LEFT, True, "黑体", "Arial"),
            "heading_2": (12.0, WD_ALIGN_PARAGRAPH.LEFT, True, "黑体", "Arial"),
            "heading_3": (10.5, WD_ALIGN_PARAGRAPH.LEFT, True, "黑体", "Arial"),
            "body": (12.0, WD_ALIGN_PARAGRAPH.JUSTIFY, None, "宋体", "Times New Roman"),
            "fig_caption": (10.5, WD_ALIGN_PARAGRAPH.CENTER, None, "宋体", "Times New Roman"),
            "table_caption": (10.5, WD_ALIGN_PARAGRAPH.CENTER, None, "宋体", "Times New Roman"),
            "refs_heading": (14.0, WD_ALIGN_PARAGRAPH.LEFT, True, "黑体", "Arial"),
            "ref_entry": (10.5, WD_ALIGN_PARAGRAPH.JUSTIFY, None, "宋体", "Times New Roman"),
        }
        deviations = defaultdict(list)
        unresolved = defaultdict(list)
        for i, p in enumerate(self.paragraphs):
            cls = self._para_class.get(i)
            if cls not in expectations:
                continue
            t = p.text.strip()
            if not t:
                continue
            size_exp, align_exp, bold_exp, cjk_exp, latin_exp = expectations[cls]
            runs = [r for r in p.runs if r.text.strip()]
            if not runs:
                continue
            align = self.effective_alignment(p)
            loc = f"段落{i+1}: {t[:45]}"
            if align is None:
                # Word's default paragraph alignment is left; treat it as measurable.
                align = WD_ALIGN_PARAGRAPH.LEFT
            if align != align_exp:
                deviations[(cls, "对齐")].append((loc, str(align), str(align_exp)))
            for run in runs:
                size = self.effective_run_size(run, p.style)
                if size is None:
                    unresolved[(cls, "字号")].append(loc)
                elif not (run.font.superscript or run.font.subscript) and abs(size - size_exp) > 0.6:
                    deviations[(cls, "字号")].append((loc, size, size_exp))
                bold = self.effective_run_bold(run, p.style)
                if bold_exp is not None and (bold if bold is not None else False) != bold_exp:
                    deviations[(cls, "加粗")].append((loc, bold, bold_exp))
                fonts = self.effective_run_fonts(run, p.style)
                if CHINESE_RE.search(run.text):
                    cjk = fonts.get("eastAsia")
                    if cjk is None:
                        unresolved[(cls, "中文字体")].append(loc)
                    elif cjk_exp.casefold() not in cjk.casefold():
                        deviations[(cls, "中文字体")].append((loc, cjk, cjk_exp))
                if re.search(r"[A-Za-z0-9]", run.text):
                    lat = fonts.get("ascii") or fonts.get("hAnsi")
                    if lat is None:
                        unresolved[(cls, "西文字体")].append(loc)
                    elif latin_exp.casefold() not in lat.casefold():
                        deviations[(cls, "西文字体")].append((loc, lat, latin_exp))

            if cls == "body":
                first = self.pformat_value(p, "first_line_indent")
                cm = first.cm if first is not None else 0.0
                if abs(cm - 0.74) > 0.18:
                    deviations[(cls, "首行缩进")].append((loc, f"{cm:.2f} cm", "约0.74 cm"))
                ls = self.pformat_value(p, "line_spacing")
                rule = self.pformat_value(p, "line_spacing_rule")
                ok = False
                if isinstance(ls, float) and abs(ls - 1.5) <= 0.08:
                    ok = True
                if rule == WD_LINE_SPACING.ONE_POINT_FIVE:
                    ok = True
                if not ok:
                    deviations[(cls, "行距")].append((loc, str(ls) if ls is not None else "默认/未设置", "1.5倍"))
                before = self.pformat_value(p, "space_before")
                after = self.pformat_value(p, "space_after")
                if before is not None and abs(before.pt) > 0.6:
                    deviations[(cls, "段前")].append((loc, f"{before.pt:.1f} pt", "0 pt"))
                if after is not None and abs(after.pt) > 0.6:
                    deviations[(cls, "段后")].append((loc, f"{after.pt:.1f} pt", "0 pt"))
            elif cls == "ref_entry":
                ls = self.pformat_value(p, "line_spacing")
                rule = self.pformat_value(p, "line_spacing_rule")
                if (isinstance(ls, float) and abs(ls - 1.0) > 0.12) or (ls is None and rule not in (WD_LINE_SPACING.SINGLE,)):
                    deviations[(cls, "行距")].append((loc, str(ls), "单倍"))
                left = self.pformat_value(p, "left_indent")
                first = self.pformat_value(p, "first_line_indent")
                left_cm = left.cm if left is not None else 0.0
                first_cm = first.cm if first is not None else 0.0
                if abs(left_cm - 0.74) > 0.18 or abs(first_cm + 0.74) > 0.18:
                    deviations[(cls, "悬挂缩进")].append((loc, f"左缩进{left_cm:.2f} cm/首行{first_cm:.2f} cm", "约0.74 cm/首行-0.74 cm"))

        label_map = {
            "chapter": "章标题", "heading_1": "一级节", "heading_2": "二级节", "heading_3": "三级节",
            "body": "正文", "fig_caption": "图题", "table_caption": "表题", "refs_heading": "参考文献标题", "ref_entry": "参考文献条目",
        }
        for (cls, what), vals in deviations.items():
            examples = "; ".join(f"{v[0]}（{v[1]} -> {v[2]}）" for v in vals[:3])
            sev = "P1" if cls in {"chapter", "heading_1", "heading_2", "body", "ref_entry"} else "P2"
            self.add(sev, "字体与段落", examples, f"{label_map.get(cls, cls)}{what}存在 {len(vals)} 处可测偏差。", "统一字体字号与段落表", "批量通过Word样式统一，不逐段手工修改。", True)
        unresolved_count = sum(len(values) for values in unresolved.values())
        if unresolved_count:
            details = "; ".join(f"{kind}:{len(values)}处" for kind, values in sorted(unresolved.items()))
            self.add("P2", "格式人工确认", "全文", f"Word样式链/文档默认值仍无法可靠解析 {unresolved_count} 处格式（{details}）。", "按 run → 段落样式 → 样式继承链 → Normal/文档默认值判断", "在 Word 中检查对应段落；脚本不把未知格式误报为确定错误。", human=True)
        self.metrics["formatting_unresolved_count"] = unresolved_count

    def citation_numbers(self, text: str) -> set[int]:
        out = set()
        for g in CIT_GROUP_RE.findall(text):
            s = g.strip("[] ")
            # normalize punctuation
            s = s.replace("，", ",").replace("–", "-").replace("—", "-")
            for part in s.split(","):
                part = part.strip()
                if not part:
                    continue
                if "-" in part:
                    a, b = [x.strip() for x in part.split("-", 1)]
                    if a.isdigit() and b.isdigit():
                        a, b = int(a), int(b)
                        if 0 < a <= b <= 10000:
                            out.update(range(a, b + 1))
                elif part.isdigit():
                    out.add(int(part))
        return out

    def audit_references(self):
        refs = self.metrics["references"]
        ref_idx = self.meta.get("refs_heading_idx")
        lower5, _ = self.recent_year_bounds(self.current_year, 5)
        lower10, _ = self.recent_year_bounds(self.current_year, 10)
        ref_stats = {
            "total": len(refs), "year_identified": 0, "year_unknown": 0,
            "within_5_years": 0, "within_5_years_ratio": None,
            "within_10_years": 0, "within_10_years_ratio": None,
            "classic": 0, "future_years": [], "window_5": [lower5, self.current_year],
            "window_10": [lower10, self.current_year],
        }
        if not refs:
            self.metrics["reference_stats"] = ref_stats
            self.add("P1", "参考文献", "文末", "未识别到形如[1]的参考文献条目。", "GB/T 7714-2025顺序编码制", "确认存在独立参考文献列表并从[1]开始编号。")
            return
        nums = [n for _, n, _ in refs]
        expected = list(range(1, len(nums) + 1))
        if nums != expected:
            self.add("P1", "参考文献", "文末", f"参考文献编号为 {nums[:20]}{'...' if len(nums)>20 else ''}，并非从[1]连续编号。", "每章参考文献独立顺序编码", "使用文献管理软件重新编号。")
        body_paras = self.paragraphs[:ref_idx] if ref_idx is not None else self.paragraphs
        table_text = "\n".join(cell.text for table in self.doc.tables for row in table.rows for cell in row.cells)
        body_text = "\n".join([p.text for p in body_paras] + [table_text])
        cited = self.citation_numbers(body_text)
        refs_set = set(nums)
        missing = sorted(cited - refs_set)
        uncited = sorted(refs_set - cited)
        self.metrics["citation_count_unique"] = len(cited)
        self.metrics["missing_reference_numbers"] = missing
        self.metrics["uncited_reference_numbers"] = uncited
        if missing:
            self.add("P0", "参考文献", "正文/文末", f"正文引用了但文后缺失的编号：{missing[:30]}", "正文引文与文后文献一一对应", "补齐真实文献或修正文中编号；不得猜测文献。")
        if uncited:
            sev = "P1" if len(uncited) > max(3, len(refs)//10) else "P2"
            self.add(sev, "参考文献", "文末", f"文后存在未在正文识别到的文献编号：{uncited[:30]}", "文献应与正文引用对应", "确认是否漏引、应删除或引用格式未被识别。")

        bad_type = []
        ref_years = {}
        metadata_gaps = []
        for idx, n, txt in refs:
            type_match = re.search(r"\[(J|M|D|C|P|S|EB/OL|DS)(?:/OL)?\]", txt, re.I)
            if not type_match:
                bad_type.append(n)
            yrs = [int(y) for y in YEAR_RE.findall(txt)]
            if yrs:
                year = max(yrs)
                ref_years[n] = year
                if year > self.current_year:
                    ref_stats["future_years"].append({"reference": n, "year": year})
                elif year >= lower5:
                    ref_stats["within_5_years"] += 1
                if year <= self.current_year and year >= lower10:
                    ref_stats["within_10_years"] += 1
                elif year < lower10:
                    ref_stats["classic"] += 1
            else:
                ref_stats["year_unknown"] += 1

            # A conservative field-presence check by reference type. It never invents or
            # validates bibliographic facts and leaves ambiguous records for human review.
            if type_match:
                kind = type_match.group(1).upper()
                prefix = txt[:type_match.start()].strip(" .。;；")
                suffix = txt[type_match.end():].strip(" .。;；")
                gap = []
                if not prefix or not re.search(r"[A-Za-z\u4e00-\u9fff]", prefix):
                    gap.append("作者/责任者")
                if not prefix or len(re.sub(r"^.*?[.。]\s*", "", prefix).strip()) < 2:
                    gap.append("題名")
                if not yrs:
                    gap.append("年份")
                if kind == "J":
                    if not suffix:
                        gap.append("期刊名")
                    if not re.search(r"\b\d{1,4}(?:\s*\(\s*\d+\s*\))?\s*[:：,，]|文章号|article\s*(?:no\.?|number)", suffix, re.I):
                        gap.append("卷期及页码/文章号（需人工核对）")
                elif kind in {"M", "D", "C"} and not re.search(r"[:：]", suffix):
                    gap.append("出版地/单位或出版项")
                elif kind == "P" and not re.search(r"(?:专利号|[A-Z]{2}\s*\d{5,}|\d{7,})", txt, re.I):
                    gap.append("专利号")
                elif kind == "S" and not re.search(r"(?:GB/?T|ISO|ASTM|YY/?T|标准号)", txt, re.I):
                    gap.append("标准编号")
                elif kind in {"EB/OL", "DS"} and not re.search(r"https?://|doi\s*:\s*10\.", txt, re.I):
                    gap.append("URL/DOI")
                if gap:
                    metadata_gaps.append({"reference": n, "fields": gap})
        self.metrics["reference_years"] = ref_years
        ref_stats["year_identified"] = len(ref_years)
        denominator = len(ref_years)
        if denominator:
            ref_stats["within_5_years_ratio"] = round(ref_stats["within_5_years"] / denominator, 4)
            ref_stats["within_10_years_ratio"] = round(ref_stats["within_10_years"] / denominator, 4)
        self.metrics["reference_stats"] = ref_stats
        if bad_type:
            self.add("P2", "参考文献", "文末", f"以下条目未检测到常见文献类型标识[J]/[M]/[D]等：{bad_type[:30]}", "GB/T 7714-2025类型标识框架", "逐条核对文献类型和著录字段。", human=True)
        if metadata_gaps:
            self.metrics["reference_metadata_gaps"] = metadata_gaps
            examples = "; ".join(f"[{x['reference']}]缺{'、'.join(x['fields'])}" for x in metadata_gaps[:10])
            self.add("P2", "参考文献字段", "文后条目", f"启发式发现字段待核对：{examples}", "按文献类型核对 GB/T 7714-2025 核心著录项", "对照原始出版记录逐条核实； DOI 仅在原始来源可确认时补录。", human=True)
        if ref_stats["future_years"]:
            future = ", ".join(f"[{x['reference']}] {x['year']}年" for x in ref_stats["future_years"][:10])
            self.add("P2", "参考文献年份", "文后条目", f"发现晚于当前年份 {self.current_year} 的疑似年份：{future}。", "参考文献年份不得晚于当前年份", "检查是否为录入错误、在线优先日期或未来出版信息。", human=True)

        self.metrics["suspected_duplicate_references"] = self.find_duplicate_references(refs)
        if self.metrics["suspected_duplicate_references"]:
            pairs = self.metrics["suspected_duplicate_references"][:12]
            summary = "; ".join(f"[{p['references'][0]}] ↔ [{p['references'][1]}] ({'/'.join(p['reasons'])})" for p in pairs)
            self.add("P2", "疑似重复文献", "文后条目", f"发现疑似重复候选：{summary}。候选不等同于确认重复。", "检查完全相同条目、DOI、标题及作者年份", "人工核对原始来源；不得自动删除文献。", human=True)

        # Recent section citation freshness.
        recent_sections = self.metrics.get("recent_sections", [])
        if recent_sections:
            recent_stats = []
            headings = self.metrics["headings"]
            for rs in recent_sections:
                start = rs[0]
                end = ref_idx if ref_idx is not None else len(self.doc.paragraphs)
                for h in headings:
                    if h[2] <= rs[2] and h[0] > start:
                        end = min(end, h[0])
                        break
                sec_text = "\n".join(p.text for p in self.paragraphs[start+1:end])
                sec_cites = sorted(self.citation_numbers(sec_text))
                years = [ref_years[n] for n in sec_cites if n in ref_years]
                n5 = sum(self.current_year - 4 <= y <= self.current_year for y in years)
                n10 = sum(self.current_year - 9 <= y <= self.current_year for y in years)
                recent_stats.append({
                    "title": rs[4], "reference_count": len(sec_cites), "citation_n_with_year": len(years),
                    "year_unknown": len([n for n in sec_cites if n not in ref_years]),
                    "within_10y": n10, "within_10y_ratio": round(n10/len(years), 4) if years else None,
                    "within_5y": n5, "within_5y_ratio": round(n5/len(years), 4) if years else None,
                    "years": years,
                })
            self.metrics["recent_reference_stats"] = recent_stats

    def find_duplicate_references(self, refs):
        candidates: dict[tuple[int, int], set[str]] = defaultdict(set)
        normalized_entries = {}
        dois = defaultdict(list)
        titles = {}

        def normalized(value: str) -> str:
            value = unicodedata.normalize("NFKC", value).casefold()
            return re.sub(r"[^\w\u4e00-\u9fff]", "", value)

        def title_part(value: str) -> str:
            marker = re.search(r"\[(?:J|M|D|C|P|S|EB/OL|DS)\]", value, re.I)
            head = value[:marker.start()] if marker else value
            split = re.split(r"[.。]\s+", head.strip())
            return normalized(split[-1] if split else head)

        for _, number, text in refs:
            key = normalized(text)
            normalized_entries.setdefault(key, []).append(number)
            titles[number] = title_part(text)
            doi_matches = DOI_RE.findall(text)
            if doi_matches:
                dois[normalized(doi_matches[-1].rstrip(".,;:)")).casefold()].append(number)
        for numbers in normalized_entries.values():
            for left, right in zip(numbers, numbers[1:]):
                candidates[tuple(sorted((left, right)))].add("完全相同条目")
        for numbers in dois.values():
            for left, right in zip(numbers, numbers[1:]):
                candidates[tuple(sorted((left, right)))].add("相同 DOI")
        for index, (_, left, _) in enumerate(refs):
            left_title = titles.get(left, "")
            if len(left_title) < 8:
                continue
            for _, right, _ in refs[index + 1:]:
                right_title = titles.get(right, "")
                if len(right_title) < 8:
                    continue
                ratio = difflib.SequenceMatcher(None, left_title, right_title).ratio()
                if ratio >= 0.92:
                    candidates[tuple(sorted((left, right)))].add(f"標題相似 {ratio:.0%}")
        return [
            {"references": list(pair), "reasons": sorted(reasons)}
            for pair, reasons in sorted(candidates.items())
            if pair[0] != pair[1]
        ]

    def audit_figures_tables(self):
        chap = self.meta.get("chapter_no")
        fig_caps = self.metrics["figure_captions"]
        tab_caps = self.metrics["table_captions"]
        body_blocks = []
        table_index = 0
        last_nonempty = ""
        try:
            from docx.text.paragraph import Paragraph
            for child in self.doc.element.body.iterchildren():
                if child.tag == qn("w:p"):
                    para = Paragraph(child, self.doc._body)
                    body_blocks.append({"kind": "paragraph", "text": para.text.strip()})
                    if para.text.strip():
                        last_nonempty = para.text.strip()
                elif child.tag == qn("w:tbl"):
                    table_index += 1
                    table = self.doc.tables[table_index - 1] if table_index <= len(self.doc.tables) else None
                    repeat_header = False
                    if table is not None and table.rows:
                        try:
                            repeat_header = table.rows[0]._tr.trPr is not None and table.rows[0]._tr.trPr.find(qn("w:tblHeader")) is not None
                        except Exception:
                            pass
                    body_blocks.append({"kind": "table", "text": last_nonempty, "table": table_index, "caption_above": bool(TAB_CAP_RE.match(last_nonempty)), "repeat_header": repeat_header})
                    last_nonempty = ""
        except Exception:
            body_blocks = []

        table_metrics = [item for item in body_blocks if item.get("kind") == "table"]
        self.metrics["editable_word_tables"] = len(self.doc.tables)
        self.metrics["table_layout"] = table_metrics
        uncaptained_tables = [item["table"] for item in table_metrics if not item["caption_above"]]
        if uncaptained_tables:
            self.add("P2", "表格", f"Word表格 {uncaptained_tables[:10]}", "未在部分可编辑表格前识别到表题。", "表题位于表格上方", "检查是否缺表题或表题格式未被识别。", human=True)
        long_tables_without_header = [item["table"] for item in table_metrics if not item["repeat_header"] and item.get("table", 0) <= len(self.doc.tables) and len(self.doc.tables[item["table"]-1].rows) >= 6]
        self.metrics["tables_without_repeat_header"] = [item["table"] for item in table_metrics if not item["repeat_header"]]
        if long_tables_without_header:
            self.add("P2", "表格跨页", f"Word表格 {long_tables_without_header[:10]}", "较长表格未标记重复表头；是否跨页需结合 Word 页面布局确认。", "跨页表重复显示表头", "若表格跨页，在 Word 表格属性中设置标题行重复。", human=True)

        # Caption chapter consistency and sequence.
        for kind, caps, label in [("figure", fig_caps, "图"), ("table", tab_caps, "表")]:
            if chap is not None:
                wrong = [x for x in caps if x[1] != chap]
                if wrong:
                    self.add("P1", "图表编号", wrong[0][3], f"{label}编号章号与当前第{chap}章不一致。", "按章编号", "统一编号并更新正文交叉引用。")
            nums = [x[2] for x in caps]
            if nums and nums != list(range(1, len(nums)+1)):
                self.add("P1", "图表编号", f"{label}题", f"{label}序号为 {nums}，存在跳号/重复风险。", "按出现顺序连续编号", "检查编号并更新交叉引用。")

        # First citation before caption.
        full_paras = self.paragraphs
        for i, chn, n, caption in fig_caps + tab_caps:
            marker_re = re.compile(rf"(?:图|表)\s*{chn}\s*[-－—]\s*{n}(?!\d)")
            prior = any(marker_re.search(full_paras[j].text) for j in range(0, i) if j != i)
            if not prior:
                self.add("P2", "图表引用", caption, "未在图/表题之前识别到正文首次引用。", "正文先引用后出现图表", "在正文相应位置先引用该图/表；若交叉引用格式特殊请人工确认。")

        # Editable Word tables vs captions.
        if len(self.doc.tables) < len(tab_caps):
            self.add("P1", "表格", "全文", f"识别到 {len(tab_caps)} 个表题，但仅 {len(self.doc.tables)} 个可编辑Word表格。", "表格必须可编辑", "检查是否有表格被作为图片粘贴，并补交Excel源文件。")

        # Direct border check for obvious full-grid tables.
        gridish = []
        for ti, table in enumerate(self.doc.tables, start=1):
            try:
                borders = table._tbl.tblPr.first_child_found_in("w:tblBorders")
                if borders is None:
                    continue
                bad = []
                for tag in ["left", "right", "insideV"]:
                    el = borders.find(qn("w:" + tag))
                    if el is not None and el.get(qn("w:val")) not in (None, "nil", "none"):
                        bad.append(tag)
                if bad:
                    gridish.append(ti)
            except Exception:
                pass
        if gridish:
            self.add("P2", "表格", f"Word表格 {gridish[:10]}", "检测到显式竖线/内部竖线，可能不是三线表。", "优先三线表，一般不使用竖线", "人工确认复杂表是否确需辅助线；一般表统一为三线表。")
        self.metrics["tables_with_explicit_vertical_borders"] = gridish
        self.metrics["table_header_unit_review_required"] = bool(self.doc.tables)

        # Image effective DPI.
        image_metrics = []
        try:
            image_object_count = len(self.doc.element.body.xpath(".//w:drawing"))
        except Exception:
            image_object_count = len(self.doc.inline_shapes)
        self.metrics["image_object_count"] = image_object_count
        self.metrics["figure_caption_count"] = len(fig_caps)
        self.metrics["figure_caption_image_difference"] = len(fig_caps) - image_object_count
        self.metrics["figure_visual_review_required"] = image_object_count > 0
        if (fig_caps and image_object_count == 0) or (image_object_count > 0 and len(fig_caps) == 0) or abs(len(fig_caps) - image_object_count) >= 2:
            self.add("P2", "图题与图片对象", "全文", f"检测到 {len(fig_caps)} 个图题和 {image_object_count} 个 Word 图片对象，数量差异需核对（多联图可能包含多个对象）。", "图题与图片对象对应", "逐图核对图题、面板和浮动图片；该数量检查不能替代视觉检查。", human=True)
        if Image is not None:
            for si, shape in enumerate(self.doc.inline_shapes, start=1):
                try:
                    svg_nodes = [node for node in shape._inline.iter() if str(node.tag).endswith("}svgBlip")]
                    if svg_nodes:
                        svg_rid = svg_nodes[0].get(qn("r:embed"))
                        svg_part = self.doc.part.related_parts.get(svg_rid) if svg_rid else None
                        svg_name = os.path.basename(str(getattr(svg_part, "partname", f"image{si}.svg")))
                        image_metrics.append({"index": si, "name": svg_name, "type": getattr(svg_part, "content_type", "image/svg+xml"), "vector": True, "effective_dpi": None})
                        continue
                    blips = shape._inline.xpath('.//a:blip')
                    if not blips:
                        continue
                    rid = blips[0].get(qn('r:embed'))
                    part = self.doc.part.related_parts[rid]
                    ctype = getattr(part, 'content_type', '')
                    name = os.path.basename(str(getattr(part, 'partname', f'image{si}')))
                    if any(x in ctype.lower() for x in ['emf', 'wmf', 'svg']) or name.lower().endswith(('.emf','.wmf','.svg')):
                        image_metrics.append({"index": si, "name": name, "type": ctype, "vector": True})
                        continue
                    img = Image.open(BytesIO(part.blob))
                    pxw, pxh = img.size
                    win = shape.width / EMU_PER_INCH
                    hin = shape.height / EMU_PER_INCH
                    dpi_x = pxw / win if win else 0
                    dpi_y = pxh / hin if hin else 0
                    eff = min(dpi_x, dpi_y)
                    image_metrics.append({"index": si, "name": name, "type": ctype, "px": [pxw, pxh], "display_in": [round(win,2), round(hin,2)], "effective_dpi": round(eff,1), "vector": False})
                    if re.search(r"screenshot|screen.?shot|wechat|weixin|qqimage|capture", name, re.I) or min(pxw, pxh) < 500:
                        self.add("P2", "低分辨率/截图风险", f"内嵌图{si} {name}", f"图片尺寸为 {pxw}×{pxh} px 或文件名提示截图来源，建议核对原图与科研图像完整性。", "优先使用可追溯原始图/高质量导出文件", "人工检查来源、比例尺、面板标签、裁切和处理记录；该提示不是不端指控。", human=True)
                    if eff < 150:
                        self.add("P1", "插图", f"内嵌图{si} {name}", f"按Word显示尺寸估算有效分辨率约 {eff:.0f} dpi，明显低于出版素材最低要求。", "连续色调图最终尺寸至少300 dpi", "替换为高分辨率原图；线稿/混合图建议更高。")
                    elif eff < 300:
                        self.add("P1", "插图", f"内嵌图{si} {name}", f"按Word显示尺寸估算有效分辨率约 {eff:.0f} dpi，低于300 dpi。", "连续色调图最终尺寸至少300 dpi", "使用原始高分辨率文件重新置入。")
                except Exception as e:
                    image_metrics.append({"index": si, "error": str(e)})
        else:
            self.add("P3", "插图", "全文", "Pillow不可用，未自动计算内嵌图有效DPI。", "图片需满足分辨率要求", "人工检查或安装Pillow后重新审计。")
        self.metrics["images"] = image_metrics

        microscopy_terms = re.compile(r"显微|共聚焦|扫描电镜|透射电镜|微观形貌|micrograph|microscopy|SEM|TEM", re.I)
        scale_terms = re.compile(r"比例尺|scale\s*bar|\b(?:μ|u)?m\b|\bnm\b", re.I)
        for paragraph_index, _, _, caption in fig_caps:
            nearby = "\n".join(p.text for p in self.paragraphs[max(0, paragraph_index-2):min(len(self.paragraphs), paragraph_index+3)])
            if microscopy_terms.search(nearby) and not scale_terms.search(nearby):
                self.add("P2", "比例尺人工确认", caption, "图题或相邻正文提到显微图，但未识别到比例尺说明。", "显微图需有真实比例尺", "在原图中人工确认比例尺存在且数值准确；脚本不能验证图像像素内容。", human=True)

        # Equations count (OMML).
        try:
            omath = len(self.doc.element.xpath('.//m:oMath'))
        except Exception:
            omath = 0
        self.metrics["editable_equation_objects"] = omath
        try:
            self.metrics["display_equation_objects"] = len(self.doc.element.xpath(".//m:oMathPara"))
        except Exception:
            self.metrics["display_equation_objects"] = 0
        equation_labels = []
        for match in re.finditer(r"(?:式|公式)\s*[（(]\s*(\d+)\s*[-－—]\s*(\d+)\s*[）)]", self.all_text()):
            equation_labels.append((int(match.group(1)), int(match.group(2)), match.group(0)))
        self.metrics["equation_labels"] = equation_labels
        if chap is not None:
            wrong_equations = [item for item in equation_labels if item[0] != chap]
            if wrong_equations:
                self.add("P1", "公式编号", wrong_equations[0][2], f"公式编号章号与当前第{chap}章不一致。", "按章编号式（X-1）", "核对公式编号和正文交叉引用。")
        eq_nums = [item[1] for item in equation_labels]
        if eq_nums and eq_nums != list(range(1, len(eq_nums) + 1)):
            self.add("P1", "公式编号", "全文", f"公式序号为 {eq_nums}，存在跳号/重号或未从1开始。", "公式按章连续编号", "人工核对公式及引用编号。")
        if self.metrics["display_equation_objects"] and not equation_labels:
            self.add("P2", "公式编号", "全文", "检测到 Word 显示公式对象，但未识别到式（X-X）编号。", "显示公式按章编号并在正文引用", "检查公式编号；内联公式不要求单独编号。", human=True)

    def audit_toc(self):
        text = self.all_text(include_tables=False)
        try:
            instructions = " ".join(node.text or "" for node in self.doc.element.xpath(".//w:instrText"))
        except Exception:
            instructions = ""
        heading_present = bool(re.search(r"(?im)^\s*(?:目\s*录|Contents)\s*$", text))
        field_present = bool(re.search(r"\bTOC\b", instructions, re.I))
        self.metrics["toc"] = {
            "heading_present": heading_present,
            "word_toc_field_present": field_present,
            "page_numbers_visually_verified": False if heading_present else None,
            "status": "automatic_field_detected" if field_present else ("manual_or_unverified" if heading_present else "not_present"),
        }
        if heading_present and not field_present:
            self.add("P2", "目录", "目录", "检测到目录标题，但未检测到 Word TOC 自动目录字段，可能为手工目录。", "目录使用 Word 标题样式自动生成", "在 Word 中检查目录字段、标题映射、层级和页码；单章文件无目录不判错。", human=True)
        elif field_present:
            self.add("P2", "目录人工确认", "目录", "检测到 Word TOC 字段；脚本无法确认字段已更新或页码与排版一致。", "目录层级、点线和页码正确", "在 Word 中更新整个目录并核对页码与 Heading 映射。", human=True)

    def audit_statistics(self):
        ref_idx = self.meta.get("refs_heading_idx")
        paragraphs = self.paragraphs[:ref_idx] if ref_idx is not None else self.paragraphs
        table_text = "\n".join(cell.text for table in self.doc.tables for row in table.rows for cell in row.cells)
        text = "\n".join([p.text for p in paragraphs] + [table_text])
        cues = re.compile(r"统计学|statistical|\bn\s*=|mean\s*[±+\-]|±\s*(?:SD|SEM)|标准差|标准误|median|误差棒|误差线|error\s*bars?|显著性|\bP\s*[<=>]|ANOVA|t\s*检验|t-test|方差分析|Mann.?Whitney|Kruskal", re.I)
        if not cues.search(text):
            self.metrics["statistics"] = {"detected": False, "manual_review_required": False}
            return
        checks = {
            "sample_size_n": bool(re.search(r"\bn\s*=\s*\d+|样本量|样本数|每组\s*\d+", text, re.I)),
            "replicate_type": bool(re.search(r"生物(?:学)?重复|biological replicates?", text, re.I) and re.search(r"技术重复|technical replicates?", text, re.I)),
            "summary_expression": bool(re.search(r"mean\s*[±+\-]\s*(?:SD|SEM)|均值\s*[±+\-]\s*(?:标准差|标准误)|median|中位数", text, re.I)),
            "error_bar_definition": bool(re.search(r"误差棒|误差线|error\s*bars?", text, re.I)),
            "statistical_test": bool(re.search(r"统计检验|t\s*检验|t-test|ANOVA|方差分析|Mann.?Whitney|Kruskal|卡方|\bχ2\b|\bchi.?square\b", text, re.I)),
            "p_value_or_significance": bool(re.search(r"\bP\s*[<=>]\s*0?\.?\d+|显著性|\*{1,3}", text, re.I)),
            "significance_legend": bool(re.search(r"\*+\s*(?:P\s*[<=>]|表示|代表)|星号|显著性标记", text, re.I)),
        }
        self.metrics["statistics"] = {"detected": True, "checks": checks, "manual_review_required": True}
        missing = []
        for key, label in [
            ("sample_size_n", "样本量 n"), ("replicate_type", "生物学与技术重复类型"),
            ("summary_expression", "mean±SD/SEM 或中位数等数据表达"),
            ("error_bar_definition", "误差棒含义"), ("statistical_test", "统计检验名称"),
        ]:
            if not checks[key]:
                missing.append(label)
        if checks["p_value_or_significance"] and not checks["significance_legend"]:
            missing.append("显著性符号说明")
        if missing:
            self.add("P2", "统计报告要素", "统计结果/图表", f"启发式未识别到：{'、'.join(missing)}。", "报告 n、重复类型、数据表达、误差棒、检验和显著性说明", "检查图注、方法与结果段；若信息在其他位置，人工确认并交叉引用。", human=True)

    def audit_semantic_candidates(self):
        ref_idx = self.meta.get("refs_heading_idx")
        paragraphs = self.paragraphs[:ref_idx] if ref_idx is not None else self.paragraphs
        body_text = "\n".join(p.text for p in paragraphs)
        manual = []

        claim_candidates = []
        thesis_candidates = []
        for index, para in enumerate(paragraphs, start=1):
            text = para.text.strip()
            if not text:
                continue
            for match in STRONG_CLAIM_RE.finditer(text):
                claim_candidates.append({"paragraph": index, "phrase": match.group(0), "context": text[max(0, match.start()-35):match.end()+60]})
            if THESIS_STACKING_RE.search(text):
                thesis_candidates.append({"paragraph": index, "context": text[:160]})
        self.metrics["strong_claim_candidates"] = claim_candidates[:50]
        self.metrics["thesis_stacking_candidates"] = thesis_candidates[:50]
        if claim_candidates:
            locations = "; ".join(f"段落{x['paragraph']}“{x['phrase']}”" for x in claim_candidates[:8])
            self.add("P2", "科学表述强度候选", locations, "发现需要与证据强度逐项核对的强结论措辞；关键词本身不代表结论错误。", "因果和临床表述应与证据等级匹配", "由 AI/作者对照本段数据与引用判断；不得由脚本自动改写科学结论。", human=True)
            manual.append("对强因果、首次、临床有效等表述核对证据等级和引用。")
        if len(thesis_candidates) >= 2:
            locations = "; ".join(f"段落{x['paragraph']}" for x in thesis_candidates[:8])
            self.add("P2", "专著叙事候选", locations, "多处出现逐篇研究叙述线索；需判断是否形成论文/学生逐篇堆叠。", "按科学问题—技术发展—机制—功能—应用组织", "重组为技术发展脉络；不能仅凭关键词认定为论文拼接。", human=True)
            manual.append("判断团队学位论文材料是否已按专著叙事整合，而非逐篇堆叠。")

        normalized_paragraphs = []
        for index, para in enumerate(paragraphs, start=1):
            raw = para.text.strip()
            if len(raw) < 50 or self._para_class.get(index-1) in {"chapter", "heading_1", "heading_2", "heading_3", "ref_entry", "fig_caption", "table_caption"}:
                continue
            norm = re.sub(r"\W", "", unicodedata.normalize("NFKC", raw).casefold())
            if len(norm) >= 40:
                normalized_paragraphs.append((index, norm, raw))
        repeated = []
        exact_groups: dict[str, list[int]] = defaultdict(list)
        fuzzy_buckets: dict[tuple[str, int], list[tuple[int, str]]] = defaultdict(list)
        for paragraph_index, normalized_text, _ in normalized_paragraphs:
            exact_groups[normalized_text].append(paragraph_index)
            if len(normalized_text) >= 80:
                # Near-duplicate comparisons are limited to similar-length paragraphs
                # with a shared prefix; full-document pairwise comparison is too costly
                # for long book chapters and adds little beyond the AI review.
                fuzzy_buckets[(normalized_text[:16], len(normalized_text) // 100)].append((paragraph_index, normalized_text))
        for paragraph_indexes in exact_groups.values():
            for pos, left_index in enumerate(paragraph_indexes):
                repeated.extend({"paragraphs": [left_index, right_index], "similarity": 1.0} for right_index in paragraph_indexes[pos+1:])
        compared_pairs = set()
        for bucket_items in fuzzy_buckets.values():
            for pos, (left_index, left_norm) in enumerate(bucket_items):
                for right_index, right_norm in bucket_items[pos+1:]:
                    pair = tuple(sorted((left_index, right_index)))
                    if pair in compared_pairs or left_norm == right_norm:
                        continue
                    compared_pairs.add(pair)
                    ratio = difflib.SequenceMatcher(None, left_norm, right_norm).ratio()
                    if ratio >= 0.92:
                        repeated.append({"paragraphs": list(pair), "similarity": round(ratio, 3)})
        self.metrics["repeated_paragraph_candidates"] = repeated[:100]
        if repeated:
            pairs = "; ".join(f"段落{x['paragraphs'][0]}↔段落{x['paragraphs'][1]} ({x['similarity']:.0%})" for x in repeated[:10])
            self.add("P2", "内容重复候选", pairs, "发现文本相同或高度相似的段落候选；模板化方法描述可能合理。", "避免概念、背景、方法和结果重复展开", "作者对照上下文判断是否需要合并或交叉引用。", human=True)

        definitions = {}
        for match in re.finditer(r"\b([A-Z][A-Z0-9-]{1,11})\s*[）)]", body_text):
            acronym = match.group(1)
            start_full = body_text.rfind("（", 0, match.start())
            start_ascii = body_text.rfind("(", 0, match.start())
            start = max(start_full, start_ascii)
            segment = body_text[start+1:match.start()] if start >= 0 else ""
            english_words = re.findall(r"[A-Za-z]{2,}", segment)
            if len(english_words) >= 2 and ("," in segment or "，" in segment or len(english_words) >= 3):
                definitions.setdefault(acronym, match.start())
        acronym_candidates = []
        common = {"DNA", "RNA", "ATP", "SEM", "SD", "TEM", "AI", "DOI", "GB", "PDF", "DOCX", "URL", "ISO", "P", "N"}
        tokens = Counter(re.findall(r"\b[A-Z][A-Z0-9-]{2,10}\b", body_text))
        for acronym, count in tokens.items():
            if count < 3 or acronym in common:
                continue
            occurrences = [m.start() for m in re.finditer(rf"\b{re.escape(acronym)}\b", body_text)]
            definition_pos = definitions.get(acronym)
            if definition_pos is None or (occurrences and occurrences[0] < definition_pos):
                acronym_candidates.append({"term": acronym, "occurrences": count, "defined_at": definition_pos})
        self.metrics["abbreviation_candidates"] = acronym_candidates
        if acronym_candidates:
            summary = ", ".join(f"{item['term']}×{item['occurrences']}" for item in acronym_candidates[:15])
            self.add("P2", "缩写人工确认", "全文", f"疑似未在首次出现时定义的缩写：{summary}。", "首次出现提供中文名、英文全称和缩写", "逐项确认并统一术语；脚本不自动替换科学术语。", human=True)
            manual.append("核对缩写首次定义与同一材料/方法的中文名、英文名和缩写一致性。")

        self.metrics["semantic_manual_review_required"] = True
        self.metrics["semantic_manual_checks"] = manual

    def audit_units_abbrev(self):
        ref_idx = self.meta.get("refs_heading_idx")
        paras = self.paragraphs[:ref_idx] if ref_idx is not None else self.paragraphs
        table_text = "\n".join(cell.text for table in self.doc.tables for row in table.rows for cell in row.cells)
        text = "\n".join([p.text for p in paras] + [table_text])
        units = r"(?:mol\s*[·.]\s*L(?:\s*\^?\s*[−-]?1)?|mm|cm|μm|um|nm|mL|μL|uL|mg|μg|ug|kg|g|kPa|MPa|Pa|°C|℃|min|h|d|s)"
        bad_units = re.findall(rf"(?<![A-Za-z])\d+(?:\.\d+)?{units}(?![A-Za-z])", text, re.I)
        if bad_units:
            self.add("P2", "量和单位", "全文", f"检测到可能缺少数字-单位空格的表达，如：{', '.join(bad_units[:10])}", "数字与单位通常留半角空格", "人工核对并统一为如25 °C、10 mm、5 mL。", human=True)
        if re.search(r"\brpm\b", text, re.I):
            self.add("P2", "量和单位", "全文", "检测到 rpm。", "转速优先写 r/min", "作者核对后统一为 r/min；若离心实验可优先报告相对离心力 ×g。", human=True)
        centrifuge_x = re.findall(r"(?<!\d)\d+(?:\.\d+)?\s*[xX]\s*g(?![A-Za-z])", text)
        if centrifuge_x:
            self.add("P2", "量和单位", "全文", f"检测到疑似用 x 表示乘号的离心表达：{centrifuge_x[:10]}。", "离心力优先用 ×g", "按语义核对并统一乘号；不得误改普通字母x。", human=True)
        ranges = set(re.findall(r"\d+(?:\.\d+)?\s*([-–—~～至到])\s*\d+(?:\.\d+)?", text))
        if len(ranges) > 1:
            self.add("P2", "数字与范围", "全文", f"数值范围符号混用：{sorted(ranges)}。", "范围表达全章统一", "选择出版社或全书约定形式并统一；小数位/有效数字需结合数据精度人工判断。", human=True)
        if re.search(r"如下图|下图|下表|如下表", text):
            self.add("P2", "交叉引用", "全文", "检测到“如下图/下图/下表”等位置依赖表达。", "使用明确编号引用", "按上下文改为“见图X-X”“表X-X显示……”。", human=True)
        self.metrics["unit_candidates"] = {"missing_number_space": bad_units, "rpm_present": bool(re.search(r"\brpm\b", text, re.I)), "x_g_candidates": centrifuge_x, "range_markers": sorted(ranges), "significant_digits": "人工核对"}

    def audit_package(self):
        if not self.package_dir:
            self.metrics["package_provided"] = False
            return
        p = self.package_dir
        required = ["01_Chapter", "02_Figures", "03_FigureSources", "04_Tables_Data", "05_References", "06_Permissions", "07_Checklist"]
        missing = [x for x in required if not (p / x).exists()]
        inventory = {}
        for name in required:
            folder = p / name
            files = sorted(x for x in folder.rglob("*") if x.is_file()) if folder.exists() else []
            inventory[name] = {
                "exists": folder.exists(),
                "file_count": len(files),
                "files": [str(x.relative_to(p)) for x in files[:50]],
            }
        self.metrics["package_provided"] = True
        self.metrics["package_inventory"] = inventory
        self.metrics["package_missing_dirs"] = missing
        if missing:
            self.add("P1", "交稿包", str(p), f"缺少标准交稿目录：{', '.join(missing)}", "每章交稿包01-07目录", "补齐目录和对应材料；若本轮仅审正文，可注明尚未进入完整交稿阶段。")
        for d in required:
            if not inventory[d]["exists"] or inventory[d]["file_count"]:
                continue
            sev = "P2" if d in {"04_Tables_Data", "06_Permissions"} else "P1"
            self.add(sev, "交稿包", d, f"目录 {d} {'不存在' if d in missing else '为空'}。", "按实际需要提交章节、图件、源文件、数据、文献库、许可和清单", "补充对应材料，或由章节负责人说明该目录确不适用。", human=True)

    def run(self):
        self.classify_paragraphs()
        self.audit_page()
        self.audit_structure()
        self.audit_formatting()
        self.audit_references()
        self.audit_figures_tables()
        self.audit_toc()
        self.audit_statistics()
        self.audit_semantic_candidates()
        self.audit_units_abbrev()
        self.audit_package()
        self.issues.sort(key=lambda x: (SEV_ORDER.get(x.severity, 9), x.category, x.location))
        self.metrics["manual_review_required"] = True
        self.metrics["manual_review_categories"] = sorted({i.category for i in self.issues if i.human_review})
        return self

    def verdict(self) -> str:
        c = Counter(i.severity for i in self.issues)
        if c["P0"] > 0 or c["P1"] >= 10:
            return "暂不建议交稿"
        if c["P1"] > 0:
            return "修改后可进入下一轮编辑"
        return "可进入下一轮编辑"

    def as_json(self):
        return {
            "file": str(self.path),
            "audited_at": datetime.now().isoformat(timespec="seconds"),
            "verdict": self.verdict(),
            "meta": self.meta,
            "metrics": self.metrics,
            "issue_counts": dict(Counter(i.severity for i in self.issues)),
            "issues": [asdict(i) for i in self.issues],
        }

    def as_markdown(self) -> str:
        counts = Counter(i.severity for i in self.issues)
        est = self.metrics.get("text_estimate", {})
        lines = []
        lines.append("# DOCX自动自查结果（确定性检查）")
        lines.append("")
        lines.append(f"- 文件：`{self.path.name}`")
        lines.append(f"- 识别章节：{self.meta.get('chapter_title') or '未识别'}")
        lines.append(f"- 自动结论：**{self.verdict()}**（仍需模型完成语义、版权与科研诚信检查）")
        lines.append(f"- P0：{counts['P0']}；P1：{counts['P1']}；P2：{counts['P2']}；P3：{counts['P3']}")
        if est:
            lines.append(f"- 估算正文规模：{est.get('estimated_words_chars', 0):,}（中文字符+英文词+数字token）")
        lines.append(f"- 一级小节数：{self.metrics.get('first_level_section_count', '未识别')}")
        if self.metrics.get("recent_reference_stats"):
            for s in self.metrics["recent_reference_stats"]:
                lines.append(f"- 前沿专题 `{s['title']}`：可识别年份引用 {s['citation_n_with_year']} 条；近5年 {s['within_5y']} 条；近10年 {s['within_10y']} 条")
        lines.append("")
        if not self.issues:
            lines.append("未发现可自动测量的偏差。仍需完成语义、版权、科研诚信和源文件完整性检查。")
            return "\n".join(lines)
        lines.append("## 问题清单")
        lines.append("")
        lines.append("|级别|类别|位置|问题|规范/建议|自动修复|")
        lines.append("|---|---|---|---|---|---|")
        for i in self.issues:
            loc = i.location.replace('|','/').replace('\n',' ')
            msg = i.message.replace('|','/').replace('\n',' ')
            rule = i.rule.replace('|','/').replace('\n',' ')
            lines.append(f"|{i.severity}|{i.category}|{loc}|{msg}|{rule}|{'是' if i.auto_fixable else '否'}|")
        lines.append("")
        lines.append("## 自动指标")
        lines.append("")
        lines.append(f"- Word可编辑表格：{self.metrics.get('editable_word_tables', 0)}")
        lines.append(f"- 图题：{len(self.metrics.get('figure_captions', []))}；表题：{len(self.metrics.get('table_captions', []))}")
        lines.append(f"- Word可编辑公式对象：{self.metrics.get('editable_equation_objects', 0)}")
        imgs = self.metrics.get('images', [])
        if imgs:
            lines.append(f"- 内嵌图片对象：{len(imgs)}")
            for im in imgs[:20]:
                if im.get('vector'):
                    lines.append(f"  - 图{im['index']} {im.get('name','')}：矢量/可缩放类型")
                elif 'effective_dpi' in im:
                    lines.append(f"  - 图{im['index']} {im.get('name','')}：{im['px'][0]}x{im['px'][1]} px；Word显示约 {im['display_in'][0]}x{im['display_in'][1]} in；有效DPI约 {im['effective_dpi']}")
        return "\n".join(lines)


def main():
    ap = argparse.ArgumentParser(description="Audit a chapter DOCX against the book house style.")
    ap.add_argument("docx", type=Path)
    ap.add_argument("--current-year", type=int, default=datetime.now().year)
    ap.add_argument("--expected-chapter", type=int, default=None)
    ap.add_argument("--package-dir", type=Path, default=None)
    ap.add_argument("--json", dest="json_path", type=Path, default=None)
    ap.add_argument("--markdown", dest="md_path", type=Path, default=None)
    args = ap.parse_args()

    if not args.docx.exists():
        ap.error(f"DOCX not found: {args.docx}")
    if args.docx.suffix.lower() != ".docx":
        ap.error("This helper only accepts .docx. Use manual/content-only audit for other formats.")
    if args.package_dir and not args.package_dir.exists():
        ap.error(f"Package directory not found: {args.package_dir}")

    auditor = Auditor(args.docx, args.current_year, args.package_dir, args.expected_chapter).run()
    md = auditor.as_markdown()
    print(md)
    if args.md_path:
        args.md_path.parent.mkdir(parents=True, exist_ok=True)
        args.md_path.write_text(md, encoding="utf-8")
    if args.json_path:
        args.json_path.parent.mkdir(parents=True, exist_ok=True)
        args.json_path.write_text(json.dumps(auditor.as_json(), ensure_ascii=False, indent=2), encoding="utf-8")


if __name__ == "__main__":
    main()
