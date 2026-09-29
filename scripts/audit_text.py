#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Content-only checks for Markdown and TXT chapter manuscripts."""

from __future__ import annotations

import difflib
import json
import re
import unicodedata
from collections import Counter, defaultdict
from datetime import datetime
from pathlib import Path
from typing import Any

CHAPTER_RE = re.compile(r"^\s*第\s*(\d+)\s*章\s*(.*)$")
NUMBERED_HEADING_RE = re.compile(r"^\s*(?:#{1,6}\s*)?(\d+(?:\.\d+){1,6})\s+(.+?)\s*$")
REF_RE = re.compile(r"^\s*\[(\d+)\]\s*(.+)$")
CIT_RE = re.compile(r"\[\s*\d+(?:\s*[-–—,，]\s*\d+)*\s*\]")
YEAR_RE = re.compile(r"\b(19\d{2}|20\d{2})\b")
RECENT_RE = re.compile(r"最新进展|研究进展|发展方向|前沿|新技术|先进制造|个性化|精准|智能|未来发展|发展趋势|趋势展望", re.I)
DOI_RE = re.compile(r"\b10\.\d{4,9}/[^\s<>\]\[{}]+", re.I)
STRONG_RE = re.compile(r"首次实现|首次发现|首次提出|证实|证明|决定性|直接调控|完全|彻底解决|显著优于|临床有效|实现临床应用|巨大临床价值")


def text_estimate(text: str) -> dict[str, int]:
    cjk = len(re.findall(r"[\u4e00-\u9fff]", text))
    latin = len(re.findall(r"[A-Za-z]+(?:[-'][A-Za-z]+)*", text))
    numbers = len(re.findall(r"\d+(?:\.\d+)?", text))
    return {"chinese_chars": cjk, "latin_words": latin, "number_tokens": numbers, "estimated_words_chars": cjk + latin + numbers}


def cited_numbers(text: str) -> set[int]:
    result: set[int] = set()
    for group in CIT_RE.findall(text):
        for item in group.strip("[] ").replace("，", ",").replace("–", "-").replace("—", "-").split(","):
            item = item.strip()
            if "-" in item:
                parts = item.split("-", 1)
                if all(part.strip().isdigit() for part in parts):
                    low, high = (int(part.strip()) for part in parts)
                    if 0 < low <= high <= 10000:
                        result.update(range(low, high + 1))
            elif item.isdigit():
                result.add(int(item))
    return result


class TextAuditor:
    def __init__(self, path: Path, current_year: int | None = None, package_dir: Path | None = None, expected_chapter: int | None = None):
        self.path = Path(path)
        self.current_year = current_year if current_year is not None else datetime.now().year
        self.package_dir = package_dir
        self.expected_chapter = expected_chapter
        self.text = self.path.read_text(encoding="utf-8-sig", errors="replace")
        self.lines = self.text.splitlines()
        self.issues: list[dict[str, Any]] = []
        self.metrics: dict[str, Any] = {"input_format": self.path.suffix.lower().lstrip("."), "current_year": self.current_year}
        self.meta: dict[str, Any] = {"chapter_no": None, "chapter_title": None, "current_year": self.current_year}

    def add(self, severity: str, category: str, location: str, message: str, rule: str, suggested_fix: str, auto: bool = False, human: bool = False):
        self.issues.append({
            "severity": severity, "category": category, "location": location, "message": message,
            "rule": rule, "suggested_fix": suggested_fix, "auto_fixable": auto, "human_review": human,
        })

    def run(self):
        chapter = None
        chapter_titles = []
        headings = []
        references = []
        in_references = False
        for index, line in enumerate(self.lines, start=1):
            content = line.strip().strip("# ").strip()
            chapter_match = CHAPTER_RE.match(content)
            if chapter_match:
                number = int(chapter_match.group(1))
                chapter_titles.append({"line": index, "chapter": number, "text": content})
                if chapter is None:
                    chapter = number
                    self.meta.update({"chapter_no": number, "chapter_title": content})
            if re.fullmatch(r"参考文献\s*[:：]?", content):
                in_references = True
            ref_match = REF_RE.match(content)
            if ref_match:
                in_references = True
                references.append({"line": index, "number": int(ref_match.group(1)), "text": ref_match.group(2)})
                continue
            heading_match = NUMBERED_HEADING_RE.match(line)
            if heading_match:
                number_text = heading_match.group(1)
                headings.append({"line": index, "number": number_text, "level": number_text.count("."), "title": heading_match.group(2)})

        if chapter is None:
            self.add("P1", "章节结构", "全文", "未识别到“第X章”章标题。", "章标题格式：第1章……", "确认章标题及编号。", auto=True)
        elif self.expected_chapter is not None and chapter != self.expected_chapter:
            self.add("P1", "章节结构", self.meta.get("chapter_title", "全文"), f"识别为第{chapter}章，与请求的第{self.expected_chapter}章不一致。", "检查请求章节与稿件一致", "确认选择正确稿件。")
        if chapter_titles and len({item["chapter"] for item in chapter_titles}) > 1:
            self.add("P1", "章节结构", f"行{chapter_titles[0]['line']}", "文档中出现多个章号。", "单章文件只包含一个章号", "检查是否误合并其他章节。")

        first_level = [item for item in headings if item["level"] == 1]
        self.metrics["headings"] = headings
        self.metrics["first_level_section_count"] = len(first_level)
        if len(first_level) > 15:
            self.add("P1", "章节结构", "全文", f"一级小节共 {len(first_level)} 个，超过15个。", "一级小节不超过15个", "合并相近小节，保留主线。")
        if any(item["level"] > 3 for item in headings):
            item = next(item for item in headings if item["level"] > 3)
            self.add("P1", "章节结构", f"行{item['line']} {item['number']}", "标题层级超过四级。", "标题层级不超过四级", "上提或合并标题层级。")
        sequences: dict[tuple[int, ...], list[int]] = defaultdict(list)
        last_level = 0
        jumps = []
        for item in headings:
            parts = tuple(int(x) for x in item["number"].split("."))
            sequences[parts[:-1]].append(parts[-1])
            if (last_level == 0 and item["level"] > 1) or (last_level and item["level"] > last_level + 1):
                jumps.append(item)
            last_level = item["level"]
            if chapter is not None and parts[0] != chapter:
                self.add("P1", "章节结构", f"行{item['line']} {item['number']}", f"小节章号与第{chapter}章不一致。", "节编号首位应与章号一致", "统一编号并核对交叉引用。")
        for parent, values in sequences.items():
            if values != list(range(1, max(values) + 1)):
                self.add("P1", "章节结构", ".".join(map(str, parent)) or "一级小节", f"标题编号序列 {values[:25]} 存在跳号/重号/未从1开始。", "各级标题从1开始连续编号", "核对同层级标题编号。")
        if jumps:
            self.add("P1", "章节结构", f"行{jumps[0]['line']} {jumps[0]['number']}", "标题层级发生跳级。", "标题层级不得跳过", "补齐或调整上级标题。")

        body_lines = []
        for line in self.lines:
            content = line.strip().strip("# ").strip()
            if re.fullmatch(r"参考文献\s*[:：]?", content) or REF_RE.match(content):
                break
            body_lines.append(line)
        body_text = "\n".join(body_lines)
        self.metrics["text_estimate"] = text_estimate(body_text)
        if chapter == 1 and not 5000 <= self.metrics["text_estimate"]["estimated_words_chars"] <= 10000:
            self.add("P2", "章节规模", "全文", f"估算规模 {self.metrics['text_estimate']['estimated_words_chars']:,}，第1章规划目标为5,000–10,000。", "章节篇幅是规划目标，不是机械退稿条件", "结合章节功能和编辑策划判断篇幅。", human=True)
        elif chapter is not None and 2 <= chapter <= 10 and not 45000 <= self.metrics["text_estimate"]["estimated_words_chars"] <= 65000:
            self.add("P2", "章节规模", "全文", f"估算规模 {self.metrics['text_estimate']['estimated_words_chars']:,}，第2–10章规划目标约50,000–60,000。", "章节篇幅是规划目标，不是机械退稿条件", "结合章节功能和编辑策划判断篇幅。", human=True)

        recent_sections = [item for item in headings if RECENT_RE.search(item["title"])]
        self.metrics["recent_sections"] = recent_sections
        if chapter is not None and 2 <= chapter <= 10 and not recent_sections:
            self.add("P1", "近年进展", "全文", "未识别到专门的近5–10年进展/发展方向专题。", "第2–10章须设置前沿专题", "增加前沿/个性化/智能化/先进制造等专题并核对近年研究。")

        reference_nums = [item["number"] for item in references]
        cited = cited_numbers(body_text)
        self.metrics["references"] = references
        self.metrics["missing_reference_numbers"] = sorted(cited - set(reference_nums))
        self.metrics["uncited_reference_numbers"] = sorted(set(reference_nums) - cited)
        if references and reference_nums != list(range(1, len(reference_nums) + 1)):
            self.add("P1", "参考文献", "文末", f"参考文献编号为 {reference_nums[:25]}，未从[1]连续编号。", "GB/T 7714-2025顺序编码制", "核对编号，不自动重排文献。")
        if self.metrics["missing_reference_numbers"]:
            self.add("P0", "参考文献", "正文/文末", f"正文引用但文后缺少：{self.metrics['missing_reference_numbers'][:25]}。", "正文引文与文后文献一一对应", "核对真实文献与引文编号，不得猜造条目。")
        if self.metrics["uncited_reference_numbers"]:
            self.add("P2", "参考文献", "文末", f"未在正文识别到引用的条目：{self.metrics['uncited_reference_numbers'][:25]}。", "文献应与正文引用对应", "确认是否漏引、误引或格式未识别。", human=True)

        current = self.current_year
        years_by_num = {}
        duplicate_reasons: dict[tuple[int, int], set[str]] = defaultdict(set)
        exact_map: dict[str, list[int]] = defaultdict(list)
        doi_map: dict[str, list[int]] = defaultdict(list)
        titles = {}
        for item in references:
            text = item["text"]
            years = [int(year) for year in YEAR_RE.findall(text)]
            if years:
                years_by_num[item["number"]] = max(years)
            normalized = re.sub(r"[^\w\u4e00-\u9fff]", "", unicodedata.normalize("NFKC", text).casefold())
            exact_map[normalized].append(item["number"])
            marker = re.search(r"\[(?:J|M|D|C|P|S|EB/OL|DS)\]", text, re.I)
            head = text[:marker.start()] if marker else text
            titles[item["number"]] = re.sub(r"[^\w\u4e00-\u9fff]", "", re.split(r"[.。]\s+", head.strip())[-1].casefold())
            for doi in DOI_RE.findall(text):
                doi_map[re.sub(r"[.,;:]+$", "", doi).casefold()].append(item["number"])
        for nums, reason in [(values, "完全相同条目") for values in exact_map.values()] + [(values, "相同 DOI") for values in doi_map.values()]:
            for pos, left in enumerate(nums):
                for right in nums[pos+1:]:
                    duplicate_reasons[tuple(sorted((left, right)))].add(reason)
        for pos, left in enumerate(reference_nums):
            for right in reference_nums[pos+1:]:
                a, b = titles.get(left, ""), titles.get(right, "")
                if len(a) >= 8 and len(b) >= 8:
                    ratio = difflib.SequenceMatcher(None, a, b).ratio()
                    if ratio >= 0.92:
                        duplicate_reasons[tuple(sorted((left, right)))].add(f"標題相似 {ratio:.0%}")
        self.metrics["suspected_duplicate_references"] = [{"references": list(pair), "reasons": sorted(reasons)} for pair, reasons in sorted(duplicate_reasons.items())]
        if duplicate_reasons:
            self.add("P2", "疑似重复文献", "文末", "; ".join(f"[{a}]↔[{b}] {'/'.join(sorted(reason))}" for (a, b), reason in list(duplicate_reasons.items())[:10]), "检查完全相同条目、DOI和相似标题", "对照原始来源人工核验，不自动删除。", human=True)

        metadata_gaps = []
        for item in references:
            text = item["text"]
            marker = re.search(r"\[(J|M|D|C|P|S|EB/OL|DS)\]", text, re.I)
            if not marker:
                metadata_gaps.append({"reference": item["number"], "fields": ["文献类型标识"]})
                continue
            kind = marker.group(1).upper()
            prefix, suffix = text[:marker.start()].strip(" .。;；"), text[marker.end():].strip(" .。;；")
            missing = []
            if not prefix or not re.search(r"[A-Za-z\u4e00-\u9fff]", prefix):
                missing.append("作者/责任者")
            if not prefix or len(re.sub(r"^.*?[.。]\s*", "", prefix).strip()) < 2:
                missing.append("题名")
            if not YEAR_RE.search(text):
                missing.append("年份")
            if kind == "J":
                if not suffix:
                    missing.append("期刊名")
                if suffix and not re.search(r"\b\d{1,4}(?:\s*\(\s*\d+\s*\))?\s*[:：,，]|文章号|article\s*(?:no\.?|number)", suffix, re.I):
                    missing.append("卷期及页码/文章号（需人工核对）")
            elif kind in {"M", "D", "C"} and suffix and not re.search(r"[:：]", suffix):
                missing.append("出版地/单位或出版项")
            elif kind == "P" and not re.search(r"(?:专利号|[A-Z]{2}\s*\d{5,}|\d{7,})", text, re.I):
                missing.append("专利号")
            elif kind == "S" and not re.search(r"(?:GB/?T|ISO|ASTM|YY/?T|标准号)", text, re.I):
                missing.append("标准编号")
            elif kind in {"EB/OL", "DS"} and not re.search(r"https?://|doi\s*:\s*10\.", text, re.I):
                missing.append("URL/DOI")
            if missing:
                metadata_gaps.append({"reference": item["number"], "fields": missing})
        self.metrics["reference_metadata_gaps"] = metadata_gaps
        if metadata_gaps:
            self.add("P2", "参考文献字段", "文后条目", f"启发式发现字段待核对：{metadata_gaps[:12]}。", "按GB/T 7714-2025和文献类型核对核心著录项", "对照原始出版记录逐条核实；不自动补 DOI。", human=True)
        if duplicate_reasons:
            self.add("P2", "疑似重複文獻", "文末", "; ".join(f"[{a}]↔[{b}] {'/'.join(sorted(r))}" for (a, b), r in list(duplicate_reasons.items())[:10]), "核查完全相同條目、DOI、標題", "人工核对原始来源；不得自动删除。", human=True)

        lower5, lower10 = current - 4, current - 9
        recognized = len(years_by_num)
        n5 = sum(lower5 <= year <= current for year in years_by_num.values())
        n10 = sum(lower10 <= year <= current for year in years_by_num.values())
        future = [{"reference": num, "year": year} for num, year in years_by_num.items() if year > current]
        stats = {"total": len(references), "year_identified": recognized, "year_unknown": len(references)-recognized,
                 "within_5_years": n5, "within_5_years_ratio": round(n5/recognized, 4) if recognized else None,
                 "within_10_years": n10, "within_10_years_ratio": round(n10/recognized, 4) if recognized else None,
                 "classic": sum(year < lower10 for year in years_by_num.values() if year <= current), "future_years": future,
                 "window_5": [lower5, current], "window_10": [lower10, current]}
        self.metrics["reference_stats"] = stats
        if future:
            self.add("P2", "参考文献年份", "文末", f"发现晚于当前年份 {current} 的疑似年份：{future[:10]}。", "核对在线优先或未来日期", "对照原始出版记录人工核验。", human=True)
        recent_counts = []
        for section in recent_sections:
            following = next((item["line"] for item in headings if item["line"] > section["line"] and item["level"] <= section["level"]), len(self.lines)+1)
            segment = "\n".join(self.lines[section["line"]:following-1])
            section_refs = cited_numbers(segment)
            section_years = [years_by_num[num] for num in section_refs if num in years_by_num]
            c5 = sum(lower5 <= year <= current for year in section_years)
            c10 = sum(lower10 <= year <= current for year in section_years)
            recent_counts.append({"title": section["title"], "reference_count": len(section_refs), "citation_n_with_year": len(section_years),
                                  "year_unknown": len(section_refs)-len(section_years), "within_5y": c5,
                                  "within_5y_ratio": round(c5/len(section_years), 4) if section_years else None,
                                  "within_10y": c10, "within_10y_ratio": round(c10/len(section_years), 4) if section_years else None})
        self.metrics["recent_reference_stats"] = recent_counts

        stats_cues = re.compile(r"统计学|statistical|\bn\s*=|mean\s*[±+\-]|±\s*(?:SD|SEM)|标准差|标准误|median|误差棒|误差线|error\s*bars?|显著性|\bP\s*[<=>]|ANOVA|t\s*检验|方差分析", re.I)
        stat_checks = {}
        if stats_cues.search(body_text):
            stat_checks = {
                "sample_size_n": bool(re.search(r"(?<![A-Za-z])n\s*=\s*\d+|样本量|样本数|每组\s*\d+", body_text, re.I)),
                "replicate_type": bool(re.search(r"生物(?:学)?重复|biological replicates?", body_text, re.I)) and bool(re.search(r"技术重复|technical replicates?", body_text, re.I)),
                "summary_expression": bool(re.search(r"mean\s*[±+\-]\s*(?:SD|SEM)|均值\s*[±+\-]\s*(?:标准差|标准误)|median|中位数", body_text, re.I)),
                "error_bar_definition": bool(re.search(r"误差棒|误差线|error\s*bars?", body_text, re.I)),
                "statistical_test": bool(re.search(r"统计检验|t\s*检验|t-test|ANOVA|方差分析|Mann.?Whitney|Kruskal|卡方", body_text, re.I)),
                "p_value_or_significance": bool(re.search(r"\bP\s*[<=>]\s*0?\.?\d+|显著性|\*{1,3}", body_text, re.I)),
                "significance_legend": bool(re.search(r"\*+\s*(?:P\s*[<=>]|表示|代表)|星号|显著性标记", body_text, re.I)),
            }
            missing_stats = [key for key, value in stat_checks.items() if not value]
            if missing_stats:
                self.add("P2", "统计报告要素", "正文/图表", f"启发式未识别到统计要素：{missing_stats}。", "报告n、重复类型、数据表达、误差棒、检验和显著性说明", "检查正文、图注和表注；统计方法是否恰当仍需作者/统计人员判断。", human=True)
        self.metrics["statistics"] = {"detected": bool(stats_cues.search(body_text)), "checks": stat_checks, "manual_review_required": True}

        unit_pattern = r"(?:mol\s*[·.]\s*L(?:\s*\^?\s*[−-]?1)?|mm|cm|μm|um|nm|mL|μL|uL|mg|μg|ug|kg|g|kPa|MPa|Pa|°C|℃|min|h|d|s)"
        missing_unit_space = re.findall(rf"(?<![A-Za-z])\d+(?:\.\d+)?{unit_pattern}(?![A-Za-z])", body_text, re.I)
        rpm_present = bool(re.search(r"\brpm\b", body_text, re.I))
        xg_candidates = re.findall(r"(?<!\d)\d+(?:\.\d+)?\s*[xX]\s*g(?![A-Za-z])", body_text)
        range_markers = sorted(set(re.findall(r"\d+(?:\.\d+)?\s*([-–—~～至到])\s*\d+(?:\.\d+)?", body_text)))
        self.metrics["unit_candidates"] = {"missing_number_space": missing_unit_space, "rpm_present": rpm_present, "x_g_candidates": xg_candidates, "range_markers": range_markers, "significant_digits": "人工核对"}
        if missing_unit_space or rpm_present or xg_candidates or len(range_markers) > 1:
            self.add("P2", "量和单位", "全文", f"单位/范围提示：缺空格{missing_unit_space[:8]}；rpm={rpm_present}；疑似xg={xg_candidates[:8]}；范围符号={range_markers}。", "数字与单位留空格、rpm优先r/min、范围表达统一", "结合科学含义人工确认；不自动改动数值或符号。", human=True)

        acronyms = Counter(re.findall(r"\b[A-Z][A-Z0-9-]{2,10}\b", body_text))
        known_acronyms = {"DNA", "RNA", "ATP", "SEM", "SD", "TEM", "AI", "DOI", "GB", "PDF", "DOCX", "URL", "ISO", "P", "N"}
        abbreviation_candidates = [{"term": term, "occurrences": count} for term, count in acronyms.items() if count >= 3 and term not in known_acronyms]
        self.metrics["abbreviation_candidates"] = abbreviation_candidates
        if abbreviation_candidates:
            self.add("P2", "缩写人工确认", "全文", f"需核对首次定义的缩写候选：{abbreviation_candidates[:12]}。", "首次出现提供中文名、英文全称和缩写", "逐项核对术语和上下文；脚本不自动替换术语。", human=True)

        self.metrics["format_checks"] = {key: "未自动测量" for key in ["page_size", "orientation", "margins", "font", "font_size", "paragraph_spacing", "images", "image_dpi"]}
        self.metrics["statistics"] = {"detected": bool(re.search(r"统计学|\bn\s*=|mean\s*[±+\-]|P\s*[<=>]|ANOVA|误差棒", body_text, re.I)), "manual_review_required": True}
        strong = [{"line": i, "phrase": m.group(0)} for i, line in enumerate(self.lines, start=1) for m in STRONG_RE.finditer(line)]
        self.metrics["strong_claim_candidates"] = strong[:50]
        if strong:
            self.add("P2", "科学表述强度候选", "; ".join(f"行{x['line']}“{x['phrase']}”" for x in strong[:8]), "发现需要与证据强度逐项核对的强结论措辞。", "结论强度需与引用和证据匹配", "由 AI/作者审核证据；脚本不得改写科学结论。", human=True)
        self.metrics["package_provided"] = bool(self.package_dir)
        if self.package_dir:
            required_dirs = ["01_Chapter", "02_Figures", "03_FigureSources", "04_Tables_Data", "05_References", "06_Permissions", "07_Checklist"]
            missing_dirs = [name for name in required_dirs if not (self.package_dir / name).is_dir()]
            self.metrics["package_inventory"] = {
                name: {
                    "exists": (self.package_dir / name).is_dir(),
                    "file_count": sum(1 for item in (self.package_dir / name).rglob("*") if item.is_file()) if (self.package_dir / name).is_dir() else 0,
                    "files": [str(item.relative_to(self.package_dir)) for item in sorted((self.package_dir / name).rglob("*")) if item.is_file()][:50] if (self.package_dir / name).is_dir() else [],
                }
                for name in required_dirs
            }
            self.metrics["package_missing_dirs"] = missing_dirs
            self.metrics["package_missing_dirs"] = missing_dirs
            if missing_dirs:
                self.add("P1", "交稿包", str(self.package_dir), f"缺少标准目录：{', '.join(missing_dirs)}。", "交稿包应检查01–07目录", "补齐或由负责人说明不适用。")
        self.metrics["manual_review_required"] = True
        self.metrics["manual_review_categories"] = ["字体/页面/图片DPI未自动测量", "专著逻辑", "统计方法适当性", "图像视觉/科研诚信", "版权与许可", "科学结论证据强度"]
        self.issues.sort(key=lambda item: (int(item["severity"][1:]), item["category"], item["location"]))
        self.metrics["recent_sections"] = recent_sections
        return self

    def verdict(self) -> str:
        levels = Counter(item["severity"] for item in self.issues)
        if levels["P0"] or levels["P1"] >= 10:
            return "暂不建议交稿"
        if levels["P1"]:
            return "修改后可进入下一轮编辑"
        return "可进入下一轮编辑"

    def as_json(self) -> dict[str, Any]:
        return {"file": str(self.path), "audited_at": datetime.now().isoformat(timespec="seconds"),
                "verdict": self.verdict(), "meta": self.meta, "metrics": self.metrics,
                "issue_counts": dict(Counter(item["severity"] for item in self.issues)), "issues": self.issues}


def main():
    import argparse
    parser = argparse.ArgumentParser(description="Audit Markdown/TXT content without page/layout measurements.")
    parser.add_argument("file", type=Path)
    parser.add_argument("--current-year", type=int, default=datetime.now().year)
    parser.add_argument("--expected-chapter", type=int, default=None)
    parser.add_argument("--package-dir", type=Path)
    parser.add_argument("--json", dest="json_path", type=Path)
    args = parser.parse_args()
    if not args.file.is_file():
        parser.error("input file not found")
    result = TextAuditor(args.file, args.current_year, args.package_dir, args.expected_chapter).run().as_json()
    print(json.dumps(result, ensure_ascii=False, indent=2))
    if args.json_path:
        args.json_path.parent.mkdir(parents=True, exist_ok=True)
        args.json_path.write_text(json.dumps(result, ensure_ascii=False, indent=2), encoding="utf-8")


if __name__ == "__main__":
    main()
