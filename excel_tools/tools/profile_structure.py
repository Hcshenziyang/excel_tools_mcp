"""MCP tool: analyze worksheet structure signatures."""

from excel_tools.models.schemas import ExcelErrorDetail, MCPResponse, ProfileStructureRequest, ProfileStructureResult
from excel_tools.operations.profile_structure import profile_structure
from excel_tools.runtime.mcp import excel_mcp
from excel_tools.runtime.tool_runtime import error_response_from_exception, success_response


@excel_mcp.tool()
async def excel_profile_structure(
    request: ProfileStructureRequest,
) -> MCPResponse[ProfileStructureResult | ExcelErrorDetail]:
    """分析工作表的结构签名（数据类型 + 合并状态）。"""
    try:
        result = await profile_structure(
            file_path=request.file_path,
            sheet=request.sheet,
            max_rows_to_scan=request.max_rows_to_scan,
            max_cols_to_scan=request.max_cols_to_scan,
        )
        return success_response(result, status_message="分析成功")
    except Exception as exc:
        return error_response_from_exception(exc, ProfileStructureResult)
