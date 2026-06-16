# -*- coding:utf-8 -*-
"""
Excel 文件检查工具
"""

# 工具介绍：只读取excel文件元信息，不触碰cell内容。
from excel_tools.mcp_instance import excel_mcp
from excel_tools.core_utils import resolve_excel_file_path
from excel_tools.read_backends import inspect_workbook
from pydantic import BaseModel, Field


class SheetInfo(BaseModel):
    name: str = Field(..., description="Sheet名称，在workbook中显示的名称。")
    rows: int | None = Field(...,description="包含内容的行数。如果workbook没有声明维度，并且检测失败，则可能为None。",)
    cols: int | None = Field(...,description="包含内容的列数。如果workbook没有声明维度，并且检测失败，则可能为None。",)
    dimension_source: str = Field(...,description="行/列的确定方式：'declared' (从xlsx维度标签确定)，'calculated' (从calculate_dimension确定)，或'unknown'。")

class InspectResult(BaseModel):
    file_path: str  # 文件路径
    file_size_bytes: int  # 文件大小
    sheet_count: int  # 表单数量
    sheets: list[SheetInfo]  # 每个表单的详细信息
    backend: str  # 读取后端

@excel_mcp.tool()
async def excel_inspect(file_path: str = "") -> dict[str, object]:
    async with resolve_excel_file_path(file_path=file_path) as path:
        file_size = path.stat().st_size  # 获取文件大小
        backend, summaries = inspect_workbook(path)
        sheet_infos = [
            SheetInfo(
                name=summary.name,
                rows=summary.rows,
                cols=summary.cols,
                dimension_source=summary.dimension_source,
            )
            for summary in summaries
        ]
        result = InspectResult(
            file_path=str(path),
            file_size_bytes=file_size,
            sheet_count=len(sheet_infos),
            sheets=sheet_infos,
            backend=backend,
        )
        return result.model_dump()
