"""Worksheet name resolution helpers."""

from __future__ import annotations

import re
from collections.abc import Callable, Sequence

from excel_tools.exceptions import SheetNotFoundError

_WHITESPACE_RE = re.compile(r"\s+")


def resolve_sheet_name(requested_sheet: str, available_sheets: Sequence[str]) -> str:
    """Resolve a possibly whitespace-normalized sheet name to the workbook name."""
    if requested_sheet in available_sheets:
        return requested_sheet

    available = list(available_sheets)
    collapsed_match = _unique_normalized_match(requested_sheet, available, _collapse_whitespace)
    removed_match = _unique_normalized_match(requested_sheet, available, _remove_whitespace)

    if _has_whitespace(requested_sheet):
        matched = _resolve_unique_match(requested_sheet, available, collapsed_match, "collapsed_whitespace")
        if matched is not None:
            return matched

    matched = _resolve_unique_match(requested_sheet, available, removed_match, "removed_whitespace")
    if matched is not None:
        return matched

    matched = _resolve_unique_match(requested_sheet, available, collapsed_match, "collapsed_whitespace")
    if matched is not None:
        return matched

    raise SheetNotFoundError(
        message=f"表单 '{requested_sheet}' 未找到。",
        suggested_action=f"可用表单: {available}",
        details={"requested_sheet": requested_sheet, "available_sheets": available},
    )


def _unique_normalized_match(
    requested_sheet: str,
    available_sheets: Sequence[str],
    normalizer: Callable[[str], str],
) -> list[str]:
    requested_key = normalizer(requested_sheet)
    return [sheet for sheet in available_sheets if normalizer(sheet) == requested_key]


def _resolve_unique_match(
    requested_sheet: str,
    available_sheets: Sequence[str],
    candidates: list[str],
    strategy: str,
) -> str | None:
    if len(candidates) == 1:
        return candidates[0]
    if len(candidates) > 1:
        raise SheetNotFoundError(
            message=f"表单 '{requested_sheet}' 未找到，空白归一后匹配到多个候选。",
            suggested_action=f"请使用完整工作表名称。候选表单: {candidates}",
            details={
                "requested_sheet": requested_sheet,
                "available_sheets": list(available_sheets),
                "candidate_sheets": candidates,
                "match_strategy": strategy,
            },
        )
    return None


def _has_whitespace(value: str) -> bool:
    return _WHITESPACE_RE.search(value) is not None


def _collapse_whitespace(value: str) -> str:
    return _WHITESPACE_RE.sub(" ", value).strip()


def _remove_whitespace(value: str) -> str:
    return _WHITESPACE_RE.sub("", value)
