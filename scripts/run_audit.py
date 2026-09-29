#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Internal one-line workflow runner: discover, audit, optionally repair a copy."""

from __future__ import annotations

import argparse
import hashlib
import json
import os
import re
import sys
from dataclasses import asdict, dataclass
from datetime import datetime
from pathlib import Path
from typing import Any

from review_contract import review_completeness

SUPPORTED_SUFFIXES = {".docx": 3, ".md": 2, ".txt": 1}
REQUIRED_PACKAGE_DIRS = ("01_Chapter", "02_Figures", "03_FigureSources", "04_Tables_Data", "05_References", "06_Permissions", "07_Checklist")
EXCLUDED_PARTS = {".git", ".venv", "venv", "node_modules", "vendor", "build", "dist", "__pycache__", ".pytest_cache", "自查结果", "audit-results", "audit_results", "backup", "backups", "old", "archive", "archives", "temp", "tmp", "test", "tests", "fixtures", "testdata", "test_data"}
EXCLUDED_STEM_RE = re.compile(r"(?i)(?:_自查报告|_自动检查结果|_规范修订版|_修订版|(?:^|[_.-])(?:backup|old|archive|temp|tmp)(?:$|[_.-]))")
VERSION_RE = re.compile(r"(?i)(?:^|[_.\-])v(?:er)?[_.\-]?(\d+(?:[._-]\d+)*)")
DATE_RE = re.compile(r"(?<!\d)(20\d{2})[-_.]?(0[1-9]|1[0-2])[-_.]?([0-2]\d|3[01])(?!\d)")


@dataclass
class Invocation:
    chapter: int | None
    title: str | None
    mode: str
    scope: str


def parse_invocation(text: str) -> Invocation:
    chapter_match = re.search(r"第\s*0*(\d+)\s*章", text, re.I)
    if not chapter_match:
        chapter_match = re.search(r"\b(?:Ch(?:apter)?\s*)0*(\d+)\b", text, re.I)
    title_match = re.search(r"[“\"「『](.+?)[”\"」』]\s*(?:章节)?(?:自查|检查|验收)?", text)
    title = title_match.group(1).strip() if title_match else None
    intent_text = text.replace(title, "") if title else text
    repair = bool(re.search(r"自查并修复|修复|修改|规范化|按规范调整", intent_text))
    if re.search(r"参考文献|文献|引文|引证", intent_text):
        scope = "references"
    elif re.search(r"图表|插图|图片|表格", intent_text):
        scope = "figures_tables"
    elif re.search(r"近\s*5\s*[–-]?\s*10\s*年|近年研究|研究进展|前沿", intent_text):
        scope = "recent_progress"
    elif re.search(r"专著写作逻辑|专著逻辑|叙事逻辑", intent_text):
        scope = "monograph_logic"
    elif re.search(r"交稿前验收|交稿验收", intent_text):
        scope = "delivery"
    elif re.search(r"格式|版式|字体|字号|页面", intent_text):
        scope = "format"
    else:
        scope = "full"
    return Invocation(int(chapter_match.group(1)) if chapter_match else None, title, "audit-and-fix" if repair else "audit-only", scope)


def excluded_path(path: Path, root: Path | None = None) -> bool:
    path = Path(path)
    if root is not None:
        try:
            relevant_parts = path.relative_to(root).parts
        except ValueError:
            relevant_parts = path.parts
    else:
        relevant_parts = path.parts
    if any(part.casefold() in EXCLUDED_PARTS for part in relevant_parts):
        return True
    stem = path.stem
    if EXCLUDED_STEM_RE.search(stem):
        return True
    if path.name.startswith("~$") or path.suffix.lower() in {".tmp", ".bak", ".log"}:
        return True
    return False


def chapter_pattern(chapter: int) -> re.Pattern:
    number = str(chapter)
    padded = number.zfill(2)
    return re.compile(
        rf"(?:第\s*0*{chapter}\s*章|(?<![A-Za-z])Ch(?:apter)?\s*0*{chapter}(?!\d)|(?<!\d){re.escape(number)}[_\-.\s]|(?<!\d){re.escape(padded)}[_\-.\s])",
        re.I,
    )


def title_matches(stem: str, title: str) -> bool:
    compact_stem = re.sub(r"[^\w\u4e00-\u9fff]", "", stem.casefold())
    compact_title = re.sub(r"[^\w\u4e00-\u9fff]", "", title.casefold())
    if len(compact_title) >= 3 and compact_title in compact_stem:
        return True
    tokens = [part for part in re.split(r"\s+|[_\-]+", title.casefold()) if len(part) >= 3]
    return bool(tokens) and sum(token in stem.casefold() for token in tokens) >= max(1, (len(tokens) + 1) // 2)


def find_candidates(workspace: Path, chapter: int | None = None, title: str | None = None) -> list[Path]:
    workspace = Path(workspace)
    if not workspace.exists():
        return []
    pattern = chapter_pattern(chapter) if chapter is not None else None
    candidates = []
    for directory, directories, filenames in os.walk(workspace):
        directories[:] = [name for name in directories if name.casefold() not in EXCLUDED_PARTS]
        current = Path(directory)
        for filename in filenames:
            path = current / filename
            if path.suffix.lower() not in SUPPORTED_SUFFIXES or excluded_path(path, workspace):
                continue
            if pattern and not pattern.search(path.stem) and not (title and title_matches(path.stem, title)):
                continue
            if title and not pattern and not title_matches(path.stem, title):
                continue
            candidates.append(path)
    return candidates


def _version_date(stem: str) -> tuple[tuple[int, ...], int]:
    version = VERSION_RE.search(stem)
    version_key: tuple[int, ...] = (0,)
    if version:
        raw = version.group(1).replace("_", ".").replace("-", ".")
        digits = tuple(int(part) for part in raw.split("."))
        version_key = tuple((digits + (0,) * 8)[:8])
    dates = [int("".join(match.groups())) for match in DATE_RE.finditer(stem)]
    file_date = max(dates, default=0)
    return version_key, file_date


def candidate_rank(path: Path) -> tuple[int, tuple[int, ...], int, int]:
    version, date = _version_date(path.stem)
    try:
        modified_ns = path.stat().st_mtime_ns
    except OSError:
        modified_ns = 0
    return SUPPORTED_SUFFIXES[path.suffix.lower()], version, date, modified_ns


def select_candidate(candidates: list[Path]) -> tuple[Path | None, list[Path]]:
    if not candidates:
        return None, []
    ranked = sorted(candidates, key=lambda path: candidate_rank(path), reverse=True)
    if len(ranked) > 1 and candidate_rank(ranked[0]) == candidate_rank(ranked[1]):
        return None, ranked
    return ranked[0], ranked


def infer_package_root(path: Path, explicit: Path | None = None) -> Path | None:
    if explicit is not None:
        return explicit
    for parent in (path.parent, *path.parents):
        if all((parent / name).is_dir() for name in REQUIRED_PACKAGE_DIRS):
            return parent
        if parent.name == "01_Chapter" and (parent.parent / "02_Figures").is_dir():
            return parent.parent
    return None


def _load_auditor(input_path: Path, current_year: int, package_dir: Path | None, expected_chapter: int | None) -> dict[str, Any]:
    if input_path.suffix.lower() == ".docx":
        try:
            from audit_docx import Auditor
        except SystemExit as exc:
            raise RuntimeError("缺少 python-docx。请由调用环境安装 requirements.txt 中的依赖后重试。") from exc
        return Auditor(input_path, current_year, package_dir, expected_chapter).run().as_json()
    from audit_text import TextAuditor
    return TextAuditor(input_path, current_year, package_dir, expected_chapter).run().as_json()


def make_report_base(input_path: Path, chapter: int | None, title: str | None) -> str:
    if chapter is not None:
        return f"第{chapter}章"
    if title:
        return re.sub(r"[\\/:*?\"<>|\s]+", "_", title).strip("_")[:80] or input_path.stem
    match = re.search(r"(?:Ch(?:apter)?|第)\s*0*(\d+)", input_path.stem, re.I)
    if match:
        return f"第{int(match.group(1))}章"
    return input_path.stem


def repaired_output_path(source: Path, chapter: int | None) -> Path:
    stem = f"第{chapter}章" if chapter is not None else source.stem
    candidate = source.with_name(f"{stem}_规范修订版.docx")
    if not candidate.exists():
        return candidate
    index = 2
    while True:
        candidate = source.with_name(f"{stem}_规范修订版_v{index:02d}.docx")
        if not candidate.exists():
            return candidate
        index += 1


def file_sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with Path(path).open("rb") as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


def run_workflow(input_path: Path, chapter: int | None = None, title: str | None = None,
                 mode: str = "audit-only", package_dir: Path | None = None,
                 current_year: int | None = None, scope: str = "full",
                 initial_review: dict[str, Any] | None = None) -> dict[str, Any]:
    current_year = current_year if current_year is not None else datetime.now().year
    input_path = Path(input_path).resolve()
    package_dir = infer_package_root(input_path, package_dir)
    base = make_report_base(input_path, chapter, title)
    output_dir = input_path.parent / "自查结果"
    output_dir.mkdir(parents=True, exist_ok=True)
    repair_info = None
    if mode == "audit-and-fix" and input_path.suffix.lower() == ".docx":
        review_complete, review_gaps = review_completeness(initial_review)
        if not review_complete:
            raise ValueError(f"安全修复前必须完成原稿AI全文审核，审核记录缺项：{review_gaps[:8]}。")
        source_hash = file_sha256(input_path)
        if initial_review.get("source_sha256") != source_hash:
            raise ValueError("修复前AI审核记录的 source_sha256 与当前原稿不一致；为防止使用过期审核结果，已拒绝修复。")
        before = _load_auditor(input_path, current_year, package_dir, chapter)
        output_copy = repaired_output_path(input_path, chapter)
        try:
            from normalize_docx import normalize_document
        except SystemExit as exc:
            raise RuntimeError("缺少 python-docx，无法安全生成规范修订版。") from exc
        repair_info = normalize_document(input_path, output_copy)
        input_path = output_copy
        package_dir = infer_package_root(input_path, package_dir)
        after = _load_auditor(input_path, current_year, package_dir, chapter)
        repair_info["before_issue_counts"] = before.get("issue_counts", {})
        repair_info["after_issue_counts"] = after.get("issue_counts", {})
        repair_info["pre_repair_ai_review_completed"] = True
        repair_info["pre_repair_review_source_sha256_matches"] = True
        audit = after
    else:
        audit = _load_auditor(input_path, current_year, package_dir, chapter)
        if mode == "audit-and-fix":
            repair_info = {"supported": False, "message": "Markdown/TXT仅支持内容检查；格式修复未执行。"}
    audit["metrics"]["invocation_scope"] = scope
    audit["metrics"]["workflow_mode"] = mode
    audit["metrics"]["source_sha256"] = file_sha256(input_path)
    audit["metrics"]["original_manuscript"] = str(input_path if not repair_info else repair_info.get("source", input_path))
    if repair_info:
        audit["metrics"]["normalization"] = repair_info
    audit_path = output_dir / f"{base}_自动检查结果.json"
    audit_path.write_text(json.dumps(audit, ensure_ascii=False, indent=2), encoding="utf-8")
    return {
        "source": str(repair_info.get("source", input_path) if repair_info else input_path),
        "source_name": input_path.name,
        "report_directory": str(output_dir),
        "audit_json": str(audit_path),
        "chapter": chapter or audit.get("meta", {}).get("chapter_no"),
        "mode": mode,
        "scope": scope,
        "repair": repair_info,
        "audit": audit,
    }


def main():
    parser = argparse.ArgumentParser(description="Internal runner used by bioactive-surface-book-auditor; authors do not need to call this manually.")
    parser.add_argument("--workspace", type=Path, default=Path.cwd())
    parser.add_argument("--request", help="Natural-language invocation forwarded internally by the Agent.")
    parser.add_argument("--chapter", type=int)
    parser.add_argument("--title")
    parser.add_argument("--input", type=Path, help="Set by the agent for an already-uploaded manuscript.")
    parser.add_argument("--package-dir", type=Path)
    parser.add_argument("--current-year", type=int, default=datetime.now().year)
    parser.add_argument("--mode", choices=("audit-only", "audit-and-fix"), default="audit-only")
    parser.add_argument("--pre-repair-review", type=Path, help="Internal completed AI review JSON required before creating a repaired DOCX copy.")
    parser.add_argument("--scope", choices=("full", "references", "figures_tables", "format", "recent_progress", "monograph_logic", "delivery"), default="full")
    args = parser.parse_args()
    try:
        if args.request:
            invocation = parse_invocation(args.request)
            args.chapter = args.chapter if args.chapter is not None else invocation.chapter
            args.title = args.title or invocation.title
            args.mode = invocation.mode
            args.scope = invocation.scope
        if args.input:
            selected = args.input
            if not selected.exists():
                raise FileNotFoundError(f"已提供的稿件文件不存在：{selected}")
        else:
            candidates = find_candidates(args.workspace, args.chapter, args.title)
            selected, ranked = select_candidate(candidates)
            if selected is None and ranked:
                names = "\n".join(str(path) for path in ranked[:10])
                print("发现两个或多个同优先级当前版本，需用户选择：\n" + names, file=sys.stderr)
                return 4
            if selected is None:
                target = f"第{args.chapter}章" if args.chapter else (args.title or "指定章节")
                raise FileNotFoundError(f"未在当前工作区找到{target}稿件。")
        if selected.suffix.lower() not in SUPPORTED_SUFFIXES:
            raise ValueError("仅支持DOCX、Markdown和TXT输入。")
        if args.pre_repair_review and not args.pre_repair_review.is_file():
            raise FileNotFoundError(f"修复前审核记录不存在：{args.pre_repair_review}")
        initial_review = json.loads(args.pre_repair_review.read_text(encoding="utf-8")) if args.pre_repair_review else None
        result = run_workflow(selected, args.chapter, args.title, args.mode, args.package_dir, args.current_year, args.scope, initial_review)
        audit = result["audit"]
        counts = audit.get("issue_counts", {})
        print(json.dumps({
            "status": "deterministic_audit_complete_ai_review_required",
            "source": result["source"], "chapter": result["chapter"], "mode": result["mode"], "scope": result["scope"],
            "audit_json": result["audit_json"], "report_directory": result["report_directory"],
            "issue_counts": {key: counts.get(key, 0) for key in ("P0", "P1", "P2", "P3")},
            "repair": result["repair"],
        }, ensure_ascii=False, indent=2))
    except FileNotFoundError as exc:
        print(str(exc), file=sys.stderr)
        return 5
    except Exception as exc:
        print(f"自查工作流未能运行：{exc}", file=sys.stderr)
        return 2
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
