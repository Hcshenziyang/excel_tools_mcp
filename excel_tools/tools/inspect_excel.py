"""MCP tool: inspect workbook metadata."""

from excel_tools.models.schemas import ExcelErrorDetail, InspectExcelRequest, InspectResult, MCPResponse
from excel_tools.operations.inspect_workbook import inspect_excel_file
from excel_tools.runtime.mcp import excel_mcp
from excel_tools.runtime.tool_runtime import error_response_from_exception, success_response


@excel_mcp.tool()
async def excel_inspect(request: InspectExcelRequest) -> MCPResponse[InspectResult | ExcelErrorDetail]:
    """检查Excel文件的基本信息，包括表单数量、每个表单的行列数等。"""
    try:
        result = await inspect_excel_file(request.file_path)
        return success_response(result, status_message="检查成功")
    except Exception as exc:
        return error_response_from_exception(exc, InspectResult)
