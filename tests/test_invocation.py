from __future__ import annotations

from pathlib import Path
import json
import os
import shutil
import subprocess
import sys

from run_audit import (
    excluded_path,
    find_candidates,
    parse_invocation,
    select_candidate,
)
from generate_report_docx import report_data, render_docx, render_markdown
from review_contract import REQUIRED_REVIEW_SECTIONS


def test_full_invocation_phrases_default_to_audit_only():
    samples = [
        "请调用 bioactive-surface-book-auditor 完成第5章自查",
        "请调用 bioactive-surface-book-auditor 自查第5章",
        "调用 bioactive-surface-book-auditor 检查第5章",
        "请用 bioactive-surface-book-auditor 检查第5章是否符合书稿规范",
        "请调用 bioactive-surface-book-auditor 完成“角膜与眼表修复材料的生物活性界面”章节自查",
    ]
    for text in samples:
        parsed = parse_invocation(text)
        assert parsed.mode == "audit-only"
        assert parsed.scope == "full"
    assert parse_invocation(samples[0]).chapter == 5
    assert parse_invocation(samples[-1]).title == "角膜与眼表修复材料的生物活性界面"


def test_repair_and_special_invocation_phrases():
    repair = parse_invocation("请调用 bioactive-surface-book-auditor 完成第5章自查并修复")
    assert repair.mode == "audit-and-fix"
    assert parse_invocation("请修改第5章格式").mode == "audit-and-fix"
    assert parse_invocation("请调用 bioactive-surface-book-auditor 检查第5章参考文献").scope == "references"
    assert parse_invocation("请调用 bioactive-surface-book-auditor 检查第5章图表").scope == "figures_tables"
    assert parse_invocation("请调用 bioactive-surface-book-auditor 检查第5章格式").scope == "format"
    assert parse_invocation("请调用 bioactive-surface-book-auditor 检查第5章近5–10年研究进展").scope == "recent_progress"
    assert parse_invocation("请调用 bioactive-surface-book-auditor 检查第5章专著写作逻辑").scope == "monograph_logic"
    assert parse_invocation("请调用 bioactive-surface-book-auditor 完成第5章交稿前验收").scope == "delivery"


def test_workspace_selects_latest_explicit_version_and_excludes_outputs(tmp_path):
    for name in [
        "Ch05_v01.docx",
        "Ch05_v03.docx",
        "Ch05_自查报告.docx",
        "Ch05_规范修订版.docx",
        "~$Ch05_v04.docx",
        "Ch05_backup.docx",
        "Ch05_backup_v02.docx",
        "Ch05_old.docx",
        "Ch05_old_notes.docx",
    ]:
        (tmp_path / name).touch()
    output_dir = tmp_path / "自查结果"
    output_dir.mkdir()
    (output_dir / "Ch05_自动检查结果.docx").touch()
    git_objects = tmp_path / ".git" / "objects"
    git_objects.mkdir(parents=True)
    (git_objects / "Ch05_hidden.docx").touch()
    candidates = find_candidates(tmp_path, chapter=5)
    selected, ranked = select_candidate(candidates)
    assert len(ranked) == 2
    assert selected.name == "Ch05_v03.docx"
    assert excluded_path(Path("Ch05_自动检查结果.docx"))
    assert not excluded_path(Path("Chapter5_gold-coating.docx"))


def test_ch05_file_is_selected_without_report_and_missing_file_is_deferred(tmp_path):
    manuscript = tmp_path / "Ch05.docx"
    report = tmp_path / "Ch05_自查报告.docx"
    revised = tmp_path / "Ch05_规范修订版.docx"
    manuscript.touch()
    report.touch()
    revised.touch()
    selected, _ = select_candidate(find_candidates(tmp_path, chapter=5))
    assert selected == manuscript
    empty = tmp_path / "empty"
    empty.mkdir()
    assert find_candidates(empty, chapter=5) == []


def test_two_equal_current_versions_remain_ambiguous(tmp_path):
    left = tmp_path / "Ch05-current.docx"
    right = tmp_path / "Ch05-final.docx"
    left.touch()
    right.touch()
    stamp = 1_700_000_000
    os.utime(left, (stamp, stamp))
    os.utime(right, (stamp, stamp))
    selected, ranked = select_candidate([left, right])
    assert selected is None
    assert len(ranked) == 2


def test_version_precedes_filename_date_then_mtime(tmp_path):
    older_version = tmp_path / "Ch05_v03_20260929.docx"
    newer_version = tmp_path / "Ch05_v04_20260928.docx"
    older_version.touch()
    newer_version.touch()
    os.utime(older_version, (1_800_000_000, 1_800_000_000))
    os.utime(newer_version, (1_700_000_000, 1_700_000_000))
    selected, _ = select_candidate([older_version, newer_version])
    assert selected == newer_version

    same_version_a = tmp_path / "Chapter05_v05_20260928.docx"
    same_version_b = tmp_path / "Chapter05_v05_20260929.docx"
    same_version_a.touch()
    same_version_b.touch()
    selected, _ = select_candidate([same_version_a, same_version_b])
    assert selected == same_version_b


def test_one_line_call_discovers_a_chapter_and_completes_report_pipeline(documents, tmp_path):
    compliant, _ = documents
    workspace = tmp_path / "workspace"
    workspace.mkdir()
    manuscript = workspace / "Ch05_骨与硬组织植入材料.docx"
    shutil.copy2(compliant, manuscript)
    request = "请调用 bioactive-surface-book-auditor 完成第5章自查"
    intent = parse_invocation(request)
    selected, ranked = select_candidate(find_candidates(workspace, intent.chapter, intent.title))
    assert selected == manuscript and len(ranked) == 1

    script = Path(__file__).resolve().parents[1] / "scripts" / "run_audit.py"
    env = os.environ.copy()
    env["PYTHONUTF8"] = "1"
    completed = subprocess.run(
        [sys.executable, str(script), "--workspace", str(workspace), "--request", request, "--current-year", "2026"],
        capture_output=True, text=True, encoding="utf-8", env=env, check=False,
    )
    assert completed.returncode == 0, completed.stderr
    result = json.loads(completed.stdout)
    audit = json.loads(Path(result["audit_json"]).read_text(encoding="utf-8"))
    assert audit["metrics"]["source_sha256"]
    review = {
        "completed": True,
        "source_sha256": audit["metrics"]["source_sha256"],
        "sections": {key: {"status": "△需人工确认", "summary": "合成测试覆盖说明。"} for key in REQUIRED_REVIEW_SECTIONS},
        "recent_progress": {"summary": "测试专用语义审核占位；不构成科学结论。"},
        "findings": [],
    }
    report = report_data(audit, review)
    stale_review = dict(review, source_sha256="stale-review-hash")
    assert not report_data(audit, stale_review)["review_complete"]
    report_dir = Path(result["report_directory"])
    md_path = report_dir / "第5章_自查报告.md"
    docx_path = report_dir / "第5章_自查报告.docx"
    md_path.write_text(render_markdown(report), encoding="utf-8")
    render_docx(report, docx_path)
    assert md_path.is_file() and docx_path.is_file()
