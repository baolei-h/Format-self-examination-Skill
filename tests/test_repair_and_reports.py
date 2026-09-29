from __future__ import annotations

import hashlib
import json
from pathlib import Path

from docx import Document
import pytest

from audit_docx import Auditor
from generate_report_docx import report_data, render_docx, render_markdown
from normalize_docx import body_text_signature, normalize_document
from review_contract import REQUIRED_REVIEW_SECTIONS
from run_audit import run_workflow


def sha256(path: Path) -> str:
    return hashlib.sha256(path.read_bytes()).hexdigest()


def sample_review():
    return {
        "completed": True,
        "sections": {
            key: {"status": "△需人工确认", "summary": "合成测试复核说明。"}
            for key in REQUIRED_REVIEW_SECTIONS
        },
        "recent_progress": {"summary": "仅基于合成测试内容，不作研究结论。", "review_types": "未统计"},
        "findings": [],
    }


def test_safe_normalization_preserves_original_hash_and_body(documents, tmp_path):
    _, broken = documents
    output = tmp_path / "第5章_规范修订版.docx"
    before_hash = sha256(broken)
    before_doc = Document(str(broken))
    before_text = body_text_signature(before_doc)
    before_audit = Auditor(broken, current_year=2026, expected_chapter=5).run()

    repair = normalize_document(broken, output)
    after_audit = Auditor(output, current_year=2026, expected_chapter=5).run()

    assert output.is_file()
    assert sha256(broken) == before_hash
    assert repair["source_sha256_unchanged"] is True
    assert repair["body_text_unchanged"] is True
    assert repair["image_pixels_unchanged"] is True
    assert body_text_signature(Document(str(output))) == before_text
    assert any(item.category == "页面版式" for item in before_audit.issues)
    assert not any(item.category == "页面版式" for item in after_audit.issues)
    assert len([i for i in after_audit.issues if i.category == "字体与段落"]) < len([i for i in before_audit.issues if i.category == "字体与段落"])


def test_workflow_audit_and_fix_names_copy_and_writes_json(documents):
    _, broken = documents
    before_hash = sha256(broken)
    with pytest.raises(ValueError, match="原稿AI全文审核"):
        run_workflow(broken, chapter=5, mode="audit-and-fix", current_year=2026)
    assert not broken.with_name("第5章_规范修订版.docx").exists()
    complete_review = {
        "completed": True,
        "source_sha256": before_hash,
        "sections": {key: {"status": "△需人工确认", "summary": "合成测试完整覆盖。"} for key in REQUIRED_REVIEW_SECTIONS},
        "recent_progress": {"summary": "合成测试审核。"},
        "findings": [],
    }
    stale_review = dict(complete_review, source_sha256="stale-source-hash")
    with pytest.raises(ValueError, match="source_sha256"):
        run_workflow(
            broken, chapter=5, mode="audit-and-fix", current_year=2026,
            initial_review=stale_review,
        )
    result = run_workflow(
        broken, chapter=5, mode="audit-and-fix", current_year=2026,
        initial_review=complete_review,
    )
    output = broken.with_name("第5章_规范修订版.docx")
    assert output.is_file()
    assert sha256(broken) == before_hash
    assert result["repair"]["body_text_unchanged"] is True
    assert result["repair"]["source_sha256_unchanged"] is True
    assert result["repair"]["image_pixels_unchanged"] is True
    assert result["repair"]["pre_repair_ai_review_completed"] is True
    assert result["repair"]["pre_repair_review_source_sha256_matches"] is True
    assert result["audit_json"] and Path(result["audit_json"]).is_file()
    audit_json = json.loads(Path(result["audit_json"]).read_text(encoding="utf-8"))
    assert audit_json["metrics"]["normalization"]["source_sha256_unchanged"] is True


def test_report_generates_readable_markdown_and_openable_word(documents, tmp_path):
    compliant, _ = documents
    audit = Auditor(compliant, current_year=2026, expected_chapter=5).run().as_json()
    report = report_data(audit, sample_review())
    draft = report_data(audit, {"completed": True, "sections": {}, "findings": []})
    assert not draft["review_complete"]
    assert any(item.startswith("sections.monograph_logic") for item in draft["review_gaps"])
    markdown_path = tmp_path / "第5章_自查报告.md"
    word_path = tmp_path / "第5章_自查报告.docx"
    markdown_path.write_text(render_markdown(report), encoding="utf-8")
    render_docx(report, word_path)

    markdown = markdown_path.read_text(encoding="utf-8")
    assert "## 11. 最终自查清单" in markdown
    assert "该结论仅代表本书内部撰写规范自查" in markdown
    assert "P0" in markdown and "P1" in markdown and "P2" in markdown and "P3" in markdown
    assert word_path.is_file() and word_path.stat().st_size > 0
    reopened = Document(str(word_path))
    text = "\n".join(p.text for p in reopened.paragraphs)
    assert "章节自查报告" in text
    assert "最终自查清单" in text
    assert len(reopened.tables) >= 2
