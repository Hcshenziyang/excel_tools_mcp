"""openpyxl workbook reader adapter."""

from pathlib import Path

from excel_tools.domain.merged_cells import build_range_read_result
from excel_tools.domain.ranges import rects_intersect
from excel_tools.exceptions import SheetNotFoundError
from excel_tools.models.schemas import Bounds, MergedCellRange, RangeReadResult, SheetSummary
from excel_tools.readers.files import open_workbook


class OpenpyxlReader:
    backend_name = "openpyxl"

    def inspect_workbook(self, path: Path) -> list[SheetSummary]:
        workbook = open_workbook(path, read_only=False, data_only=False)
        try:
            return [
                SheetSummary(
                    name=sheet_name,
                    rows=workbook[sheet_name].max_row,
                    cols=workbook[sheet_name].max_column,
                    dimension_source="declared",
                )
                for sheet_name in workbook.sheetnames
            ]
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
        workbook = open_workbook(path, read_only=False, data_only=False)
        try:
            if sheet not in workbook.sheetnames:
                raise SheetNotFoundError(
                    message=f"表单 '{sheet}' 未找到。",
                    suggested_action=f"可用表单: {workbook.sheetnames}",
                    details={"requested_sheet": sheet, "available_sheets": workbook.sheetnames},
                )

            worksheet = workbook[sheet]
            rows = worksheet.iter_rows(
                min_row=bounds.min_row,
                max_row=bounds.max_row,
                min_col=bounds.min_col,
                max_col=bounds.max_col,
            )
            values = [[cell.value for cell in row_cells] for row_cells in rows]
            merged_ranges: list[MergedCellRange] = []

            if analyze_merged_cells:
                for merged_range in worksheet.merged_cells.ranges:
                    merged_bounds = Bounds(
                        min_col=merged_range.min_col,
                        min_row=merged_range.min_row,
                        max_col=merged_range.max_col,
                        max_row=merged_range.max_row,
                    )
                    if not rects_intersect(bounds, merged_bounds):
                        continue
                    anchor_value = worksheet.cell(merged_bounds.min_row, merged_bounds.min_col).value
                    merged_ranges.append(MergedCellRange(bounds=merged_bounds, anchor_value=anchor_value))

            return build_range_read_result(
                backend=self.backend_name,
                values=values,
                requested=bounds,
                merged_ranges=merged_ranges,
                analyze_merged_cells=analyze_merged_cells,
                fill_merged_cells=fill_merged_cells,
            )
        finally:
            workbook.close()
