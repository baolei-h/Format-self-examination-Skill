#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Deterministic audit helper for 《生物活性表面可控构建技术》 chapter DOCX files.

This script performs measurable checks only. It intentionally does not judge scientific
correctness, copyright ownership, or evidence strength; the calling agent must perform
those semantic checks using the bundled audit rubric.
"""

from __future__ import annotations

import argparse
import json
import math
import os
import re
import sys
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
    print("ERROR: python-docx is required: %s" % e, file=sys.stderr)
    sys.exit(2)

try:
    from PIL import Image
except Exception:
    Image = None

EMU_PER_INCH = 914400
MM_PER_INCH = 25.4
PT_PER_INCH = 72.0

CHAPTER_RE = re.compile(r"^第\s*(\d+)\s*章\s*(.+)?$")
H_RE = re.compile(r"^(\d+(?:\.\d+){1,4})\s+(.+)$")
FIG_CAP_RE = re.compile(r"^图\s*(\d+)\s*[-－—]\s*(\d+)\s*(.*)$")
TAB_CAP_RE = re.compile(r"^表\s*(\d+)\s*[-－—]\s*(\d+)\s*(.*)$")
REF_ENTRY_RE = re.compile(r"^\[(\d+)\]\s*(.+)$")
CIT_GROUP_RE = re.compile(r"\[\s*\d+(?:\s*[-–—,，]\s*\d+)*\s*\]")
YEAR_RE = re.compile(r"\b(19\d{2}|20\d{2})\b")
CHINESE_RE = re.compile(r"[\u4e00-\u9fff]")
LATIN_WORD_RE = re.compile(r"[A-Za-z]+(?:[-'][A-Za-z]+)*")
NUMBER_TOKEN_RE = re.compile(r"\d+(?:\.\d+)?")
RECENT_TITLE_RE = re.compile(r"(最新进展|研究进展|发展方向|前沿|新技术|先进生物制造|精准.*再生|智能化|未来发展|发展趋势)")

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


class Auditor:
    def __init__(self, path: Path, current_year: int = 2026, package_dir: Path | None = None):
        self.path = path
        self.current_year = current_year
        self.package_dir = package_dir
        self.doc = Document(str(path))
        self.issues: list[Issue] = []
        self.meta: dict[str, Any] = {}
        self.metrics: dict[str, Any] = {}
        self._para_class: dict[int, str] = {}

    def add(self, severity: str, category: str, location: str, message: str, rule: str, suggested_fix: str, auto=False):
        self.issues.append(Issue(severity, category, location, message, rule, suggested_fix, auto))

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
        return out

    def effective_run_size(self, run, p_style=None) -> float | None:
        if run.font.size is not None:
            return run.font.size.pt
        if p_style is not None:
            for st in self.style_chain(p_style):
                if st.font.size is not None:
                    return st.font.size.pt
        return None

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
                if chapter_no is None:
                    chapter_no = int(m.group(1))
                    chapter_title = t
                continue
            if t == "参考文献":
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
            "refs_heading_idx": refs_heading_idx,
        })
        self.metrics.update({
            "headings": headings,
            "figure_captions": fig_caps,
            "table_captions": tab_caps,
            "references": refs,
        })

    def audit_page(self):
        for idx, sec in enumerate(self.doc.sections, start=1):
            w, h = self.mm(sec.page_width), self.mm(sec.page_height)
            loc = f"第{idx}节(Word section)"
            portrait = h is not None and w is not None and h >= w
            if not (self.approx(w, A4_W_MM, 1.5) and self.approx(h, A4_H_MM, 1.5)):
                self.add("P1", "页面版式", loc, f"页面约为 {w:.1f} x {h:.1f} mm，非统一A4作者交稿版。", "A4 210 x 297 mm", "改为A4；若出版社模板另有要求则按出版社模板。", True)
            if not portrait:
                self.add("P1", "页面版式", loc, "页面为横向或宽高异常。", "默认纵向", "正文恢复纵向；仅确有需要的超宽表单独横向。", True)
            for key, val in TARGET_MARGINS.items():
                got = self.mm(getattr(sec, f"{key}_margin"))
                if not self.approx(got, val, 1.2):
                    self.add("P1", "页面版式", loc, f"{key}边距约 {got:.1f} mm，目标 {val:.0f} mm。", "上25/下25/左26/右24 mm", "统一章节页边距。", True)

    def audit_structure(self):
        chapter_no = self.meta.get("chapter_no")
        headings = self.metrics["headings"]
        if chapter_no is None:
            self.add("P1", "章节结构", "全文", "未识别到“第X章 ……”章标题。", "章标题格式：第1章……", "确认章标题文字与编号格式。", True)
            return
        # First-level sections are numbers like 1.1 => level == 1
        h1s = [x for x in headings if x[2] == 1]
        self.metrics["first_level_section_count"] = len(h1s)
        if len(h1s) > 15:
            self.add("P1", "章节结构", "全文", f"一级小节共 {len(h1s)} 个，超过15个。", "每个大章一级小节不超过15个", "合并相近小节，保留主线。")
        too_deep = [x for x in headings if x[2] > 3]  # 1.1.1.1.1 would be level 4; four dot levels beyond section
        if too_deep:
            self.add("P1", "章节结构", too_deep[0][4], "检测到超过四级标题的编号。", "标题层级原则上不超过四级", "上提或合并标题层级。")
        wrong_chapter = [x for x in headings if int(x[1].split('.')[0]) != chapter_no]
        if wrong_chapter:
            ex = "; ".join(x[4] for x in wrong_chapter[:3])
            self.add("P1", "章节结构", ex, "节编号与章号不一致。", "节编号首位应与章号一致", "统一章节编号并更新交叉引用。", True)

        # Basic sequentiality for first-level headings.
        actual = []
        for x in h1s:
            parts = x[1].split('.')
            try:
                actual.append(int(parts[1]))
            except Exception:
                pass
        if actual and actual != list(range(actual[0], actual[0] + len(actual))):
            self.add("P1", "章节结构", "一级小节", f"一级小节编号序列为 {actual}，存在跳号/重复风险。", "一级小节连续编号", "检查删除/新增小节后的编号。", True)

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

        # Recent section presence for chapters 2-10.
        if 2 <= chapter_no <= 10:
            recent = [x for x in h1s if RECENT_TITLE_RE.search(x[3])]
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
        for i, p in enumerate(self.doc.paragraphs):
            cls = self._para_class.get(i)
            if cls not in expectations:
                continue
            t = p.text.strip()
            if not t:
                continue
            size_exp, align_exp, bold_exp, cjk_exp, latin_exp = expectations[cls]
            run = self.paragraph_representative_run(p)
            if run is None:
                continue
            size = self.effective_run_size(run, p.style)
            align = self.effective_alignment(p)
            bold = self.effective_run_bold(run, p.style)
            fonts = self.effective_run_fonts(run, p.style)
            loc = f"段落{i+1}: {t[:45]}"
            if size is not None and abs(size - size_exp) > 0.6:
                deviations[(cls, "字号")].append((loc, size, size_exp))
            if align is not None and align != align_exp:
                deviations[(cls, "对齐")].append((loc, str(align), str(align_exp)))
            if bold_exp is not None and bold is not None and bold != bold_exp:
                deviations[(cls, "加粗")].append((loc, bold, bold_exp))
            cjk = fonts.get("eastAsia")
            lat = fonts.get("ascii") or fonts.get("hAnsi")
            if cjk and cjk_exp not in cjk:
                deviations[(cls, "中文字体")].append((loc, cjk, cjk_exp))
            if lat and latin_exp.lower() not in lat.lower():
                deviations[(cls, "西文字体")].append((loc, lat, latin_exp))

            if cls == "body":
                first = self.pformat_value(p, "first_line_indent")
                if first is not None:
                    cm = first.cm
                    if abs(cm - 0.74) > 0.18:
                        deviations[(cls, "首行缩进")].append((loc, f"{cm:.2f} cm", "约0.74 cm"))
                ls = self.pformat_value(p, "line_spacing")
                rule = self.pformat_value(p, "line_spacing_rule")
                ok = False
                if isinstance(ls, float) and abs(ls - 1.5) <= 0.08:
                    ok = True
                if rule == WD_LINE_SPACING.ONE_POINT_FIVE:
                    ok = True
                if ls is not None and not ok:
                    deviations[(cls, "行距")].append((loc, str(ls), "1.5倍"))
                before = self.pformat_value(p, "space_before")
                after = self.pformat_value(p, "space_after")
                if before is not None and abs(before.pt) > 0.6:
                    deviations[(cls, "段前")].append((loc, f"{before.pt:.1f} pt", "0 pt"))
                if after is not None and abs(after.pt) > 0.6:
                    deviations[(cls, "段后")].append((loc, f"{after.pt:.1f} pt", "0 pt"))
            elif cls == "ref_entry":
                ls = self.pformat_value(p, "line_spacing")
                rule = self.pformat_value(p, "line_spacing_rule")
                if isinstance(ls, float) and abs(ls - 1.0) > 0.12 and rule != WD_LINE_SPACING.SINGLE:
                    deviations[(cls, "行距")].append((loc, str(ls), "单倍"))

        label_map = {
            "chapter": "章标题", "heading_1": "一级节", "heading_2": "二级节", "heading_3": "三级节",
            "body": "正文", "fig_caption": "图题", "table_caption": "表题", "refs_heading": "参考文献标题", "ref_entry": "参考文献条目",
        }
        for (cls, what), vals in deviations.items():
            examples = "; ".join(f"{v[0]}（{v[1]} -> {v[2]}）" for v in vals[:3])
            sev = "P1" if cls in {"chapter", "heading_1", "heading_2", "body", "ref_entry"} else "P2"
            self.add(sev, "字体与段落", examples, f"{label_map.get(cls, cls)}{what}存在 {len(vals)} 处可测偏差。", "统一字体字号与段落表", "批量通过Word样式统一，不逐段手工修改。", True)

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
        if not refs:
            self.add("P1", "参考文献", "文末", "未识别到形如[1]的参考文献条目。", "GB/T 7714-2025顺序编码制", "确认存在独立参考文献列表并从[1]开始编号。")
            return
        nums = [n for _, n, _ in refs]
        expected = list(range(1, len(nums) + 1))
        if nums != expected:
            self.add("P1", "参考文献", "文末", f"参考文献编号为 {nums[:20]}{'...' if len(nums)>20 else ''}，并非从[1]连续编号。", "每章参考文献独立顺序编码", "使用文献管理软件重新编号。", True)
        body_paras = self.doc.paragraphs[:ref_idx] if ref_idx is not None else self.doc.paragraphs
        body_text = "\n".join(p.text for p in body_paras)
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
        for idx, n, txt in refs:
            if not re.search(r"\[(J|M|D|C|P|S|EB/OL|DS)(?:/OL)?\]", txt, re.I):
                bad_type.append(n)
            yrs = [int(y) for y in YEAR_RE.findall(txt)]
            if yrs:
                # Prefer latest plausible year in entry.
                ref_years[n] = max(yrs)
        self.metrics["reference_years"] = ref_years
        if bad_type:
            self.add("P2", "参考文献", "文末", f"以下条目未检测到常见文献类型标识[J]/[M]/[D]等：{bad_type[:30]}", "GB/T 7714-2025类型标识框架", "逐条核对文献类型和著录字段。")

        # Recent section citation freshness.
        recent_sections = self.metrics.get("recent_sections", [])
        if recent_sections and ref_years:
            recent_stats = []
            headings = self.metrics["headings"]
            first_level_indices = [(i, h) for i, h in enumerate(headings) if h[2] == 1]
            for rs in recent_sections:
                start = rs[0]
                end = ref_idx if ref_idx is not None else len(self.doc.paragraphs)
                for h in headings:
                    if h[2] == 1 and h[0] > start:
                        end = min(end, h[0])
                        break
                sec_text = "\n".join(p.text for p in self.doc.paragraphs[start+1:end])
                sec_cites = sorted(self.citation_numbers(sec_text))
                years = [ref_years[n] for n in sec_cites if n in ref_years]
                if years:
                    n10 = sum(y >= self.current_year - 10 for y in years)
                    n5 = sum(y >= self.current_year - 5 for y in years)
                    recent_stats.append({"title": rs[4], "citation_n_with_year": len(years), "within_10y": n10, "within_5y": n5, "years": years})
                    if n10 / len(years) < 0.6:
                        self.add("P2", "近年进展", rs[4], f"可识别年份的引用中，近10年占比约 {n10}/{len(years)} ({n10/len(years):.0%})。", "前沿专题以近5-10年研究为主体", "补充代表性近年原始研究，同时保留必要经典文献。")
            self.metrics["recent_reference_stats"] = recent_stats

    def audit_figures_tables(self):
        chap = self.meta.get("chapter_no")
        fig_caps = self.metrics["figure_captions"]
        tab_caps = self.metrics["table_captions"]
        # Caption chapter consistency and sequence.
        for kind, caps, label in [("figure", fig_caps, "图"), ("table", tab_caps, "表")]:
            if chap is not None:
                wrong = [x for x in caps if x[1] != chap]
                if wrong:
                    self.add("P1", "图表编号", wrong[0][3], f"{label}编号章号与当前第{chap}章不一致。", "按章编号", "统一编号并更新正文交叉引用。", True)
            nums = [x[2] for x in caps]
            if nums and nums != list(range(nums[0], nums[0]+len(nums))):
                self.add("P1", "图表编号", f"{label}题", f"{label}序号为 {nums}，存在跳号/重复风险。", "按出现顺序连续编号", "检查编号并更新交叉引用。", True)

        # First citation before caption.
        full_paras = self.doc.paragraphs
        for i, chn, n, caption in fig_caps + tab_caps:
            marker_re = re.compile(rf"(?:图|表)\s*{chn}\s*[-－—]\s*{n}(?!\d)")
            prior = any(marker_re.search(full_paras[j].text) for j in range(0, i) if j != i)
            if not prior:
                self.add("P2", "图表引用", caption, "未在图/表题之前识别到正文首次引用。", "正文先引用后出现图表", "在正文相应位置先引用该图/表；若交叉引用格式特殊请人工确认。")

        # Editable Word tables vs captions.
        self.metrics["editable_word_tables"] = len(self.doc.tables)
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
            self.add("P2", "表格", f"Word表格 {gridish[:10]}", "检测到显式竖线/内部竖线，可能不是三线表。", "优先三线表，一般不使用竖线", "人工确认复杂表是否确需辅助线；一般表统一为三线表。", True)

        # Image effective DPI.
        image_metrics = []
        if Image is not None:
            for si, shape in enumerate(self.doc.inline_shapes, start=1):
                try:
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
                    if eff < 150:
                        self.add("P1", "插图", f"内嵌图{si} {name}", f"按Word显示尺寸估算有效分辨率约 {eff:.0f} dpi，明显低于出版素材最低要求。", "连续色调图最终尺寸至少300 dpi", "替换为高分辨率原图；线稿/混合图建议更高。")
                    elif eff < 300:
                        self.add("P1", "插图", f"内嵌图{si} {name}", f"按Word显示尺寸估算有效分辨率约 {eff:.0f} dpi，低于300 dpi。", "连续色调图最终尺寸至少300 dpi", "使用原始高分辨率文件重新置入。")
                except Exception as e:
                    image_metrics.append({"index": si, "error": str(e)})
        else:
            self.add("P3", "插图", "全文", "Pillow不可用，未自动计算内嵌图有效DPI。", "图片需满足分辨率要求", "人工检查或安装Pillow后重新审计。")
        self.metrics["images"] = image_metrics

        # Equations count (OMML).
        try:
            omath = len(self.doc.element.xpath('.//m:oMath'))
        except Exception:
            omath = 0
        self.metrics["editable_equation_objects"] = omath

    def audit_units_abbrev(self):
        ref_idx = self.meta.get("refs_heading_idx")
        paras = self.doc.paragraphs[:ref_idx] if ref_idx is not None else self.doc.paragraphs
        text = "\n".join(p.text for p in paras)
        # Number immediately followed by common SI unit without space. °C exception still house standard uses space.
        bad_units = re.findall(r"(?<![A-Za-z])\d+(?:\.\d+)?(?:mm|cm|mL|μL|uL|mg|μg|ug|g|kg|kPa|MPa|Pa|°C|min|h|d)\b", text)
        if bad_units:
            self.add("P2", "量和单位", "全文", f"检测到可能缺少数字-单位空格的表达，如：{', '.join(bad_units[:10])}", "数字与单位通常留半角空格", "批量核对并统一为如25 °C、10 mm、5 mL。", True)
        # Discouraged rpm.
        if re.search(r"\brpm\b", text, re.I):
            self.add("P2", "量和单位", "全文", "检测到 rpm。", "转速优先写 r/min", "统一为 r/min；若离心实验可优先报告相对离心力 ×g。", True)
        # Below/above figure wording.
        if re.search(r"如下图|下图|下表|如下表", text):
            self.add("P2", "交叉引用", "全文", "检测到“如下图/下图/下表”等位置依赖表达。", "使用明确编号引用", "改为“见图X-X”“表X-X显示……”。", True)

    def audit_package(self):
        if not self.package_dir:
            return
        p = self.package_dir
        required = ["01_Chapter", "02_Figures", "03_FigureSources", "04_Tables_Data", "05_References", "06_Permissions", "07_Checklist"]
        missing = [x for x in required if not (p / x).exists()]
        self.metrics["package_missing_dirs"] = missing
        if missing:
            self.add("P1", "交稿包", str(p), f"缺少标准交稿目录：{', '.join(missing)}", "每章交稿包01-07目录", "补齐目录和对应材料；若本轮仅审正文，可注明尚未进入完整交稿阶段。")
        else:
            # Empty-source warnings.
            for d in ["02_Figures", "03_FigureSources", "05_References", "06_Permissions", "07_Checklist"]:
                files = [x for x in (p/d).rglob('*') if x.is_file()]
                if not files:
                    sev = "P2" if d == "06_Permissions" else "P1"
                    self.add(sev, "交稿包", d, f"目录 {d} 为空。", "提交对应源文件/文献库/权限材料/自查表", "补交文件或明确说明该章确无此类材料。")

    def run(self):
        self.classify_paragraphs()
        self.audit_page()
        self.audit_structure()
        self.audit_formatting()
        self.audit_references()
        self.audit_figures_tables()
        self.audit_units_abbrev()
        self.audit_package()
        self.issues.sort(key=lambda x: (SEV_ORDER.get(x.severity, 9), x.category, x.location))
        return self

    def verdict(self) -> str:
        c = Counter(i.severity for i in self.issues)
        if c["P0"] > 0:
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
    ap.add_argument("--current-year", type=int, default=2026)
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

    auditor = Auditor(args.docx, args.current_year, args.package_dir).run()
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
