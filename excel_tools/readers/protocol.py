"""Reader interface shared by workbook backend adapters."""

from pathlib import Path
from typing import Protocol

from excel_tools.models.schemas import Bounds, RangeReadResult, SheetSummary


class WorkbookReader(Protocol):
    backend_name: str

    def inspect_workbook(self, path: Path) -> list[SheetSummary]: ...

    def read_range(
        self,
        path: Path,
        sheet: str,
        bounds: Bounds,
        analyze_merged_cells: bool = False,
        fill_merged_cells: bool = False,
    ) -> RangeReadResult: ...
