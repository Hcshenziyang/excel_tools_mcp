"""Inspect workbook metadata."""

from excel_tools.models.schemas import InspectResult, SheetInfo
from excel_tools.readers.files import resolve_excel_file_path
from excel_tools.readers.selector import get_reader


async def inspect_excel_file(file_path: str) -> InspectResult:
    """Inspect workbook metadata without reading full sheet contents."""
    async with resolve_excel_file_path(file_path=file_path) as path:
        file_size = path.stat().st_size
        reader = get_reader(path)
        summaries = reader.inspect_workbook(path)
        return InspectResult(
            file_path=str(path),
            file_size_bytes=file_size,
            sheet_count=len(summaries),
            sheets=[
                SheetInfo(
                    name=summary.name,
                    rows=summary.rows,
                    cols=summary.cols,
                    dimension_source=summary.dimension_source,
                )
                for summary in summaries
            ],
            backend=reader.backend_name,
        )
