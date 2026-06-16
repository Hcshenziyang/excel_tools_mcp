"""Shared Pydantic models for MCP tool inputs and outputs."""

from __future__ import annotations

from typing import Any, Generic, TypeVar, overload

from pydantic import BaseModel, ConfigDict, Field, field_validator

_T = TypeVar("_T")


class ToolModel(BaseModel):
    """Base model with stricter defaults for tool schemas."""

    model_config = ConfigDict(extra="forbid", str_strip_whitespace=True)


class MCPResponse(ToolModel, Generic[_T]):
    status: bool = Field(default=True, description="调用状态，成功为True，失败为False")
    status_message: str = Field(default="调用成功", description="调用状态描述信息，失败时为错误信息")
    tool_result: _T | None = Field(default=None, description="业务数据，失败时可为None")

    @overload
    @staticmethod
    def success(tool_result: _T, status_message: str = "成功") -> "MCPResponse[_T]": ...

    @overload
    @staticmethod
    def success(tool_result: None = None, status_message: str = "成功") -> "MCPResponse[None]": ...

    @staticmethod
    def success(tool_result=None, status_message="成功"):
        return MCPResponse(status=True, status_message=status_message, tool_result=tool_result)

    @overload
    @staticmethod
    def error(tool_result: _T, status_message: str = "失败") -> "MCPResponse[_T]": ...

    @overload
    @staticmethod
    def error(tool_result: None = None, status_message: str = "失败") -> "MCPResponse[None]": ...

    @staticmethod
    def error(tool_result=None, status_message="失败"):
        return MCPResponse(status=False, status_message=status_message, tool_result=tool_result)


class ExcelErrorDetail(ToolModel):
    error_code: str = Field(..., description="错误码")
    message: str = Field(..., description="错误信息")
    suggested_action: str = Field(default="", description="建议操作")
    details: dict[str, Any] = Field(default_factory=dict, description="结构化错误上下文")


class Bounds(ToolModel):
    model_config = ConfigDict(extra="forbid", frozen=True)

    min_col: int = Field(..., ge=1)
    min_row: int = Field(..., ge=1)
    max_col: int = Field(..., ge=1)
    max_row: int = Field(..., ge=1)


class MergedCellRange(ToolModel):
    model_config = ConfigDict(extra="forbid", frozen=True)

    bounds: Bounds
    anchor_value: Any = Field(default=None)


class MergedCellInfo(ToolModel):
    range: str = Field(..., description="合并单元格范围")
    anchor: str = Field(..., description="锚点单元格坐标")
    value: Any = Field(default=None, description="锚点值")
    applied: bool = Field(default=False, description="是否已将锚点值虚拟填充到返回数据中")


class SheetSummary(ToolModel):
    model_config = ConfigDict(extra="forbid", frozen=True)

    name: str = Field(..., description="Sheet名称")
    rows: int | None = Field(default=None, description="工作表行数")
    cols: int | None = Field(default=None, description="工作表列数")
    dimension_source: str = Field(..., description="行列维度来源")


class RangeReadResult(ToolModel):
    backend: str = Field(..., description="读取后端")
    values: list[list[Any]] = Field(default_factory=list, description="二维表格数据")
    merged_cells: list[MergedCellInfo] = Field(default_factory=list, description="合并单元格信息")
    merged_cells_analyzed: bool = Field(default=False, description="是否分析了合并单元格")
    merged_cells_filled: bool = Field(default=False, description="是否已虚拟回填合并单元格值")


class FilePathRequest(ToolModel):
    file_path: str = Field(..., min_length=1, description="本地 Excel 文件路径")

    @field_validator("file_path")
    @classmethod
    def validate_file_path(cls, value: str) -> str:
        if not value:
            raise ValueError("file_path 不能为空")
        return value


class SheetRequest(FilePathRequest):
    sheet: str = Field(..., min_length=1, description="工作表名称")

    @field_validator("sheet")
    @classmethod
    def validate_sheet(cls, value: str) -> str:
        if not value:
            raise ValueError("sheet 不能为空")
        return value


class InspectExcelRequest(FilePathRequest):
    """Request schema for workbook inspection."""


class SheetInfo(ToolModel):
    name: str = Field(..., description="Sheet名称，在workbook中显示的名称。")
    rows: int | None = Field(
        default=None,
        description="包含内容的行数。如果workbook没有声明维度，并且检测失败，则可能为None。",
    )
    cols: int | None = Field(
        default=None,
        description="包含内容的列数。如果workbook没有声明维度，并且检测失败，则可能为None。",
    )
    dimension_source: str = Field(
        ...,
        description="行/列的确定方式：'declared'、'calamine' 或 'unknown'。",
    )


class InspectResult(ToolModel):
    file_path: str = Field(..., description="解析后的文件绝对路径")
    file_size_bytes: int = Field(..., ge=0, description="文件大小")
    sheet_count: int = Field(..., ge=0, description="工作表数量")
    sheets: list[SheetInfo] = Field(default_factory=list, description="工作表摘要")
    backend: str = Field(..., description="读取后端")


class ReadRangeRequest(SheetRequest):
    start_cell: str = Field(default="A1", min_length=1, description="起始单元格")
    end_cell: str | None = Field(default=None, description="结束单元格")
    analyze_merged_cells: bool = Field(default=True, description="是否分析并虚拟填充合并单元格")

    @field_validator("start_cell")
    @classmethod
    def validate_start_cell(cls, value: str) -> str:
        if not value:
            raise ValueError("start_cell 不能为空")
        return value

    @field_validator("end_cell", mode="before")
    @classmethod
    def normalize_end_cell(cls, value: Any) -> Any:
        if value is None:
            return None
        normalized = str(value).strip()
        return normalized or None


class ReadRangeResult(ToolModel):
    sheet: str = Field(..., description="工作表名称")
    range: str = Field(..., description="归一化后的范围标识")
    rows: int = Field(..., ge=0, description="返回数据行数")
    cols: int = Field(..., ge=0, description="返回数据列数")
    data: list[list[Any]] = Field(default_factory=list, description="二维单元格数据")
    merged_cells_analyzed: bool = Field(default=False, description="是否分析了合并单元格")
    merged_cells_filled: bool = Field(default=False, description="是否虚拟回填了合并单元格")
    merged_cells: list[MergedCellInfo] = Field(default_factory=list, description="合并单元格详情")
    file_modified: bool = Field(default=False, description="只读工具固定为False")
    backend: str = Field(..., description="读取后端")
    resolved_file_path: str = Field(..., description="解析后的文件绝对路径")


class PatternInfo(ToolModel):
    types: str = Field(..., description="行内数据类型签名")
    merged: str = Field(..., description="行内合并状态签名")


class ProfileStructureRequest(SheetRequest):
    max_rows_to_scan: int = Field(default=5000, ge=1, description="最多扫描多少行")
    max_cols_to_scan: int = Field(default=100, ge=1, description="最多扫描多少列")


class ProfileStructureResult(ToolModel):
    sheet: str = Field(..., description="工作表名称")
    scanned_rows: tuple[int, int] = Field(..., description="实际扫描行范围")
    scanned_cols: tuple[int, int] = Field(..., description="实际扫描列范围")
    total_rows_in_sheet: int = Field(..., ge=0, description="工作表总行数")
    total_cols_in_sheet: int = Field(..., ge=0, description="工作表总列数")
    sampled: bool = Field(default=False, description="是否只采样了部分区域")
    patterns: dict[str, PatternInfo] = Field(default_factory=dict, description="模式定义")
    pattern_row_spans: dict[str, list[tuple[int, int]]] = Field(default_factory=dict, description="模式连续行段")
    backend: str = Field(..., description="分析后端")
    resolved_file_path: str = Field(..., description="解析后的文件绝对路径")


class UnmergeCellsRequest(SheetRequest):
    start_cell: str = Field(..., min_length=1, description="起始单元格")
    end_cell: str = Field(..., min_length=1, description="结束单元格")
    fill_with_anchor_value: bool = Field(default=True, description="是否用锚点值填充解除合并后的单元格")


class UnmergeCellsResult(ToolModel):
    sheet: str = Field(..., description="工作表名称")
    requested_range: str = Field(..., description="请求处理的矩形区域")
    unmerged_count: int = Field(..., ge=0, description="解除合并的区域数量")
    unmerged_ranges: list[str] = Field(default_factory=list, description="已解除的区域列表")
    fill_with_anchor_value: bool = Field(..., description="是否回填了锚点值")
    resolved_file_path: str = Field(..., description="解析后的文件绝对路径")
