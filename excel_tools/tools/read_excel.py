"""MCP tools: read worksheet ranges."""

from excel_tools.models.schemas import (
    ExcelErrorDetail,
    MCPResponse,
    ReadRangeRequest,
    ReadRangeResult,
)
from excel_tools.operations.read_range import read_excel_range
from excel_tools.runtime.mcp import excel_mcp
from excel_tools.runtime.tool_runtime import error_response_from_exception, success_response

@excel_mcp.tool()
async def excel_read_range_normalized(
    request: ReadRangeRequest,
) -> MCPResponse[ReadRangeResult | ExcelErrorDetail]:
    """读取指定区域，并在返回数据中虚拟反填合并单元格，不修改源文件。"""
    try:
        result = await read_excel_range(
            file_path=request.file_path,
            sheet=request.sheet,
            start_cell=request.start_cell,
            end_cell=request.end_cell,
            analyze_merged_cells=request.analyze_merged_cells,
            fill_merged_cells=request.analyze_merged_cells,
            )
        return success_response(result, status_message="读取成功")
    except Exception as exc:
        return error_response_from_exception(exc, ReadRangeResult)
