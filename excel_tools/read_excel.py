"""按范围读取Excel数据。"""

from excel_tools.core_utils import resolve_excel_file_path
from excel_tools.mcp_instance import excel_mcp
from excel_tools.read_backends import parse_range, read_range


def _format_range_response(
    sheet: str,
    range_label: str,
    result,
    resolved_file_path: str,
) -> dict[str, object]:
    return {
        "sheet": sheet,
        "range": range_label,
        "rows": len(result.values),
        "cols": len(result.values[0]) if result.values else 0,
        "data": result.values,
        "merged_cells_analyzed": result.merged_cells_analyzed,
        "merged_cells_filled": result.merged_cells_filled,
        "merged_cells": result.merged_cells,
        "file_modified": False,
        "backend": result.backend,
        "resolved_file_path": resolved_file_path,
    }


@excel_mcp.tool()
async def excel_read_range(
    sheet: str,
    file_path: str = "",
    start_cell: str = "A1",
    end_cell: str | None = None,
) -> dict[str, object]:
    """读取指定工作表中的矩形区域，不修改源文件。"""
    async with resolve_excel_file_path(file_path=file_path) as path:
        range_label, bounds = parse_range(start_cell, end_cell)
        result = read_range(path, sheet, bounds, analyze_merged_cells=False, fill_merged_cells=False)
        return _format_range_response(sheet, range_label, result, str(path))


@excel_mcp.tool()
async def excel_read_range_normalized(
    sheet: str,
    file_path: str = "",
    start_cell: str = "A1",
    end_cell: str | None = None,
    analyze_merged_cells: bool = True,
) -> dict[str, object]:
    """读取指定区域，并在返回数据中虚拟反填合并单元格，不修改源文件。"""
    async with resolve_excel_file_path(file_path=file_path) as path:
        range_label, bounds = parse_range(start_cell, end_cell)
        result = read_range(
            path,
            sheet,
            bounds,
            analyze_merged_cells=analyze_merged_cells,
            fill_merged_cells=analyze_merged_cells,
        )
        return _format_range_response(sheet, range_label, result, str(path))
