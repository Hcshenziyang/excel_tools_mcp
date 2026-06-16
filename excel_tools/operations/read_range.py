"""Read a worksheet range."""

from excel_tools.domain.ranges import parse_range
from excel_tools.models.schemas import ReadRangeResult
from excel_tools.readers.files import resolve_excel_file_path
from excel_tools.readers.selector import get_reader


async def read_excel_range(
    *,
    file_path: str,
    sheet: str,
    start_cell: str,
    end_cell: str | None,
    analyze_merged_cells: bool,
    fill_merged_cells: bool,
) -> ReadRangeResult:
    """Read a rectangular range from a worksheet."""
    async with resolve_excel_file_path(file_path=file_path) as path:
        range_label, bounds = parse_range(start_cell, end_cell)
        reader = get_reader(path)
        result = reader.read_range(
            path,
            sheet,
            bounds,
            analyze_merged_cells=analyze_merged_cells,
            fill_merged_cells=fill_merged_cells,
        )
        return ReadRangeResult(
            sheet=sheet,
            range=range_label,
            rows=len(result.values),
            cols=len(result.values[0]) if result.values else 0,
            data=result.values,
            merged_cells_analyzed=result.merged_cells_analyzed,
            merged_cells_filled=result.merged_cells_filled,
            merged_cells=result.merged_cells,
            file_modified=False,
            backend=result.backend,
            resolved_file_path=str(path),
        )
