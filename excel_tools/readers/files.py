"""File validation and workbook opening helpers."""

from collections.abc import AsyncIterator
from contextlib import asynccontextmanager
import os
from pathlib import Path

from dotenv import load_dotenv
from openpyxl import load_workbook
from openpyxl.utils.exceptions import InvalidFileException
from openpyxl.workbook import Workbook

from excel_tools.exceptions import (
    FileCorruptedError,
    FileEncryptedError,
    FileNotFoundError_,
    InvalidParameterError,
    UnsupportedFileError,
)

_ = load_dotenv()

SUPPORTED_EXTENSIONS = {
    ext.strip().lower()
    for ext in os.getenv("SUPPORTED_EXTENSIONS", ".xlsx,.xlsm,.xls,.xlsb,.ods").split(",")
    if ext.strip()
}
OPENPYXL_EXTENSIONS = {".xlsx", ".xlsm", ".xltx", ".xltm"}


def validate_file_path(file_path: str) -> Path:
    """Validate that the path exists, is a file, and has a supported extension."""
    path = Path(file_path)

    if not path.exists():
        raise FileNotFoundError_(
            message=f"文件未找到: {file_path}",
            suggested_action="验证文件路径是否正确，并且文件存在。",
            details={"path": str(file_path)},
        )

    if not path.is_file():
        raise FileNotFoundError_(
            message=f"路径不是文件: {file_path}",
            suggested_action="提供一个文件路径，而不是目录。",
            details={"path": str(file_path)},
        )

    if path.suffix.lower() not in SUPPORTED_EXTENSIONS:
        raise UnsupportedFileError(
            message=f"不支持的文件扩展名: {path.suffix} 支持的扩展名: {sorted(SUPPORTED_EXTENSIONS)}",
            suggested_action=".xls/.xlsb/.ods 仅支持只读工具；需要写回文件时请转换为.xlsx。",
            details={"path": str(file_path), "extension": path.suffix},
        )

    return path.resolve()


@asynccontextmanager
async def resolve_excel_file_path(file_path: str = "") -> AsyncIterator[Path]:
    """Validate and resolve a local Excel file path."""
    normalized_file_path = (file_path or "").strip()

    if not normalized_file_path:
        raise InvalidParameterError(
            message="file_path 不能为空",
            suggested_action="请提供本地 Excel 文件路径。",
            details={"file_path": file_path},
        )

    yield validate_file_path(normalized_file_path)


def open_workbook(
    path: Path,
    read_only: bool = False,
    data_only: bool = False,
) -> Workbook:
    """Open a workbook with openpyxl for supported editable formats."""
    if path.suffix.lower() not in OPENPYXL_EXTENSIONS:
        raise UnsupportedFileError(
            message=f"当前写入/编辑工具不支持此文件格式: {path.suffix}",
            suggested_action="请使用只读读取工具，或先将文件转换为.xlsx。",
            details={"path": str(path), "extension": path.suffix},
        )

    try:
        return load_workbook(str(path), read_only=read_only, data_only=data_only, keep_links=False)
    except InvalidFileException as exc:
        raise UnsupportedFileError(
            message=f"文件不是有效的xlsx文件: {exc}",
            suggested_action="确保文件是有效的.xlsx (Office Open XML)文件，而不是重命名自.xls或其他格式。",
            details={"path": str(path)},
        ) from exc
    except Exception as exc:
        msg = str(exc).lower()
        if "encrypted" in msg or "password" in msg:
            raise FileEncryptedError(
                message=f"文件加密或密码保护: {exc}",
                suggested_action="在处理文件之前删除加密。",
                details={"path": str(path)},
            ) from exc
        raise FileCorruptedError(
            message=f"无法打开文件: {exc}",
            suggested_action="验证文件是否损坏。",
            details={"path": str(path), "underlying_error": type(exc).__name__},
        ) from exc
