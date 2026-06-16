"""Helpers for consistent MCP tool validation and error responses."""

from __future__ import annotations

from typing import TypeVar, cast

from pydantic import ValidationError

from excel_tools.exceptions import ExcelMCPError
from excel_tools.models.schemas import ExcelErrorDetail, MCPResponse

_T = TypeVar("_T")


def error_response_from_exception(
    exc: Exception,
    result_type: type[_T] | None = None,
) -> MCPResponse[_T | ExcelErrorDetail]:
    """Convert known exceptions to a unified MCPResponse error payload."""
    del result_type

    if isinstance(exc, ExcelMCPError):
        detail = ExcelErrorDetail(
            error_code=exc.error_code,
            message=exc.message,
            suggested_action=exc.suggested_action,
            details=exc.details,
        )
        return cast(
            MCPResponse[_T | ExcelErrorDetail],
            MCPResponse[ExcelErrorDetail].error(tool_result=detail, status_message=exc.message),
        )

    if isinstance(exc, ValidationError):
        detail = ExcelErrorDetail(
            error_code="invalid_parameter",
            message="入参校验失败",
            suggested_action="请检查工具入参是否完整且格式正确。",
            details={"validation_errors": exc.errors(include_url=False)},
        )
        return cast(
            MCPResponse[_T | ExcelErrorDetail],
            MCPResponse[ExcelErrorDetail].error(tool_result=detail, status_message=detail.message),
        )

    detail = ExcelErrorDetail(
        error_code="internal_error",
        message=str(exc) or "工具执行失败",
        suggested_action="请稍后重试，或联系开发者排查服务端异常。",
        details={"exception_type": type(exc).__name__},
    )
    return cast(
        MCPResponse[_T | ExcelErrorDetail],
        MCPResponse[ExcelErrorDetail].error(tool_result=detail, status_message=detail.message),
    )


def success_response(tool_result: _T, status_message: str = "成功") -> MCPResponse[_T | ExcelErrorDetail]:
    """Wrap successful tool results with the standard envelope."""
    return cast(
        MCPResponse[_T | ExcelErrorDetail],
        MCPResponse[_T].success(tool_result=tool_result, status_message=status_message),
    )
