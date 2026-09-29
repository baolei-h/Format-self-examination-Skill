"""Shared completeness contract for full AI manuscript reviews."""

from __future__ import annotations

from typing import Any

REQUIRED_REVIEW_SECTIONS = (
    "monograph_logic",
    "thesis_integration",
    "team_research_continuity",
    "content_repetition",
    "mechanism_evidence",
    "clinical_extrapolation",
    "terminology_consistency",
    "statistics_review",
    "figure_visual_review",
    "copyright_integrity",
)


def review_completeness(review: Any) -> tuple[bool, list[str]]:
    if not isinstance(review, dict):
        return False, ["review object"]
    missing = []
    if review.get("completed") is not True:
        missing.append("completed=true")
    sections = review.get("sections")
    if not isinstance(sections, dict):
        sections = {}
    for name in REQUIRED_REVIEW_SECTIONS:
        value = sections.get(name)
        if not isinstance(value, dict) or not value.get("status") or not value.get("summary"):
            missing.append(f"sections.{name}.status/summary")
    recent = review.get("recent_progress")
    if not isinstance(recent, dict) or not recent.get("summary"):
        missing.append("recent_progress.summary")
    if not isinstance(review.get("findings"), list):
        missing.append("findings list")
    return not missing, missing
