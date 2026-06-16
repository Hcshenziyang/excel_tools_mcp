"""结构化错误，映射到MCP工具错误响应。"""

class ExcelMCPError(Exception):
    """基础类。每个工具错误都继承自这个类，所以MCP服务器层可以捕获一种类型并统一序列化。"""

    error_code: str = "unknown_error"  # 错误代码

    def __init__(
        self,
        message: str,
        suggested_action: str = "",
        details: dict[str, object] | None = None,
    ):
        super().__init__(message)
        self.message = message
        self.suggested_action = suggested_action
        self.details = details or {}

    def to_dict(self) -> dict[str, object]:
        return {
            "error_code": self.error_code,
            "message": self.message,
            "suggested_action": self.suggested_action,
            "details": self.details,
        }


class FileNotFoundError_(ExcelMCPError):
    """文件未找到"""
    error_code = "file_not_found"  # 文件未找到


class UnsupportedFileError(ExcelMCPError):
    """不支持的文件"""
    error_code = "unsupported_file"  # 不支持的文件


class FileCorruptedError(ExcelMCPError):
    """文件损坏"""
    error_code = "file_corrupted"  # 文件损坏


class FileEncryptedError(ExcelMCPError):
    """文件加密"""
    error_code = "file_encrypted"


class SheetNotFoundError(ExcelMCPError):
    """表单未找到"""
    error_code = "sheet_not_found"


class CapExceededError(ExcelMCPError):
    """超出容量限制"""
    error_code = "cap_exceeded"


class InvalidParameterError(ExcelMCPError):
    """无效的参数"""
    error_code = "invalid_parameter"