from enum import IntEnum
import logging

class ErrorCode(IntEnum):
    """错误代码枚举（全局唯一，按模块划分）"""
    SUCCESS = 0  # 成功（无错误）
    # Office 文件操作错误（1000-1099）
    FILE_NOT_FOUND = 1001
    FILE_READ_ERROR = 1002
    FILE_WRITE_ERROR = 1003
    FILE_INVALID_ERROR = 1004
    # 数据处理错误（1100-1199）
    DATA_VALIDATION_FAILED = 1101
    DATA_TYPE_MISMATCH = 1102
    # UI 交互错误（1200-1299）
    UI_COMPONENT_MISSING = 1201

class OfficeBaseException(Exception):
    """Office 工具基础异常类（所有业务异常的父类）"""
    def __init__(
        self,
        error_code: ErrorCode,
        message: str,
        level: int = logging.ERROR  # 默认日志级别为 ERROR
    ):
        self.error_code = error_code
        self.message = f"[错误代码 {error_code}] {message}"
        self.level = level
        super().__init__(self.message)

    def log(self, logger: logging.Logger):
        """直接通过异常对象记录日志"""
        logger.log(self.level, self.message)

# 子类示例：文件未找到异常
class FileNotFoundError(OfficeBaseException):
    def __init__(self, file_path: str):
        error_code = ErrorCode.FILE_NOT_FOUND
        message = f"文件不存在，路径：{file_path}"
        super().__init__(error_code, message, level=logging.WARNING)  # 警告级别

# 子类示例：数据验证失败异常
class DataValidationError(OfficeBaseException):
    def __init__(self, field: str, expected_type: str, actual_value: any):
        error_code = ErrorCode.DATA_VALIDATION_FAILED
        message = (
            f"字段 '{field}' 数据验证失败："
            f"期望类型 {expected_type}，实际值 {actual_value}（类型：{type(actual_value).__name__}）"
        )
        super().__init__(error_code, message, level=logging.CRITICAL)  # 严重级别