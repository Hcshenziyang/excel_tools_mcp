"""Reader selection by workbook format."""

from pathlib import Path

from excel_tools.readers.calamine_reader import CalamineReader
from excel_tools.readers.files import OPENPYXL_EXTENSIONS
from excel_tools.readers.openpyxl_reader import OpenpyxlReader
from excel_tools.readers.protocol import WorkbookReader

_OPENPYXL_READER = OpenpyxlReader()
_CALAMINE_READER = CalamineReader()


def get_reader(path: Path) -> WorkbookReader:
    """Return the reader adapter for a workbook path."""
    if path.suffix.lower() in OPENPYXL_EXTENSIONS:
        return _OPENPYXL_READER
    return _CALAMINE_READER
