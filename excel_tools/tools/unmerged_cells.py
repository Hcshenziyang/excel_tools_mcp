"""MCP-compatible write helper for unmerging cells."""

from excel_tools.models.schemas import ExcelErrorDetail, MCPResponse, UnmergeCellsRequest, UnmergeCellsResult
from excel_tools.operations.unmerge_cells import unmerge_cells
from excel_tools.runtime.tool_runtime import error_response_from_exception, success_response


async def excel_unmerge_cells(
    request: UnmergeCellsRequest,
) -> MCPResponse[UnmergeCellsResult | ExcelErrorDetail]:
    """解除与输入矩形相交的合并区整段，并可选用锚点值填满原合并格。"""
    try:
        result = await unmerge_cells(
            file_path=request.file_path,
            sheet=request.sheet,
            start_cell=request.start_cell,
            end_cell=request.end_cell,
            fill_with_anchor_value=request.fill_with_anchor_value,
        )
        return success_response(result, status_message="处理成功")
    except Exception as exc:
        return error_response_from_exception(exc, UnmergeCellsResult)
