"""python-calamine workbook reader adapter."""

from pathlib import Path
from typing import Any

from excel_tools.domain.merged_cells import build_range_read_result
from excel_tools.domain.ranges import rects_intersect
from excel_tools.domain.sheets import resolve_sheet_name
from excel_tools.exceptions import FileCorruptedError, UnsupportedFileError
from excel_tools.models.schemas import Bounds, MergedCellRange, RangeReadResult, SheetSummary


class CalamineReader:
    backend_name = "python-calamine"

    def inspect_workbook(self, path: Path) -> list[SheetSummary]:
        workbook = self._open_workbook(path)
        try:
            summaries: list[SheetSummary] = []
            for sheet_name in workbook.sheet_names:
                worksheet = workbook.get_sheet_by_name(sheet_name)
                if worksheet.end is None:
                    rows, cols = 0, 0
                else:
                    rows, cols = worksheet.end[0] + 1, worksheet.end[1] + 1
                summaries.append(
                    SheetSummary(
                        name=sheet_name,
                        rows=rows,
                        cols=cols,
                        dimension_source="calamine",
                    )
                )
            return summaries
        finally:
            workbook.close()

    def read_range(
        self,
        path: Path,
        sheet: str,
        bounds: Bounds,
        analyze_merged_cells: bool = False,
        fill_merged_cells: bool = False,
    ) -> RangeReadResult:
        workbook = self._open_workbook(path)
        try:
            sheet_names = list(workbook.sheet_names)
            resolved_sheet = resolve_sheet_name(sheet, sheet_names)
            worksheet = workbook.get_sheet_by_name(resolved_sheet)
            rows = worksheet.to_python(skip_empty_area=False, nrows=bounds.max_row)
            values = self._slice_rows(rows, bounds)
            merged_ranges: list[MergedCellRange] = []

            if analyze_merged_cells:
                merged_ranges = self._merged_ranges(worksheet, rows, bounds)

            return build_range_read_result(
                backend=self.backend_name,
                sheet=resolved_sheet,
                values=values,
                requested=bounds,
                merged_ranges=merged_ranges,
                analyze_merged_cells=analyze_merged_cells,
                fill_merged_cells=fill_merged_cells,
            )
        finally:
            workbook.close()

    def _open_workbook(self, path: Path):
        try:
            from python_calamine import CalamineWorkbook
        except ImportError as exc:
            raise UnsupportedFileError(
                message="读取此格式需要安装 python-calamine。",
                suggested_action="安装 python-calamine，或将文件转换为.xlsx后再处理。",
                details={"path": str(path), "extension": path.suffix},
            ) from exc

        try:
            return CalamineWorkbook.from_path(path)
        except Exception as exc:
            msg = str(exc).lower()
            if "unsupported" in msg or "format" in msg:
                raise UnsupportedFileError(
                    message=f"不支持的Excel文件格式: {exc}",
                    suggested_action="请确认文件扩展名和真实格式一致，或转换为.xlsx。",
                    details={"path": str(path), "extension": path.suffix},
                ) from exc
            raise FileCorruptedError(
                message=f"无法打开文件: {exc}",
                suggested_action="验证文件是否损坏或加密。",
                details={"path": str(path), "underlying_error": type(exc).__name__},
            ) from exc

    def _slice_rows(self, rows: list[list[Any]], bounds: Bounds) -> list[list[Any]]:
        values: list[list[Any]] = []
        width = bounds.max_col - bounds.min_col + 1
        for row_idx in range(bounds.min_row - 1, bounds.max_row):
            if row_idx >= len(rows):
                values.append([None] * width)
                continue
            source_row = rows[row_idx]
            row_values: list[Any] = []
            for col_idx in range(bounds.min_col - 1, bounds.max_col):
                if col_idx >= len(source_row):
                    row_values.append(None)
                else:
                    value = source_row[col_idx]
                    row_values.append(None if value == "" else value)
            values.append(row_values)
        return values

    def _merged_ranges(self, worksheet, rows: list[list[Any]], requested: Bounds) -> list[MergedCellRange]:
        merged_ranges: list[MergedCellRange] = []
        for start, end in worksheet.merged_cell_ranges:
            merged_bounds = Bounds(
                min_row=start[0] + 1,
                min_col=start[1] + 1,
                max_row=end[0] + 1,
                max_col=end[1] + 1,
            )
            if not rects_intersect(requested, merged_bounds):
                continue
            anchor_value = self._value_from_rows(rows, merged_bounds.min_row, merged_bounds.min_col)
            merged_ranges.append(MergedCellRange(bounds=merged_bounds, anchor_value=anchor_value))
        return merged_ranges

    def _value_from_rows(self, rows: list[list[Any]], row: int, col: int) -> Any:
        row_idx = row - 1
        col_idx = col - 1
        if row_idx >= len(rows) or col_idx >= len(rows[row_idx]):
            return None
        value = rows[row_idx][col_idx]
        return None if value == "" else value
