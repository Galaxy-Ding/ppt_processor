import logging
from logging.handlers import TimedRotatingFileHandler
from datetime import datetime
import os
import sys

LOG_LEVELS = {"DEBUG": 10, "INFO": 20, "WARNING": 30, "ERROR": 40}
LEVEL_MAP = {
        "DEBUG": logging.DEBUG,
        "INFO": logging.INFO,
        "WARNING": logging.WARNING,
        "ERROR": logging.ERROR
    }
class LoggerFactory:
    """增强版日志工厂，支持GUI显示和文件记录"""

    _loggers = {}
    _global_config = {}

    @classmethod
    def set_global_config(cls, log_level="INFO", log_dir="logs", fmt="%Y_%m_%d", retention_days=30):
        cls._global_config = {
            "log_level": log_level,
            "log_dir": log_dir,
            "fmt": fmt,
            "retention_days": retention_days
        }

    @classmethod
    def create_logger(
        cls,
        name: str = "default",
        log_level: str = None,
        log_dir: str = None,
        fmt: str = None,
        retention_days: int = None
    ) -> logging.Logger:
        # 优先用全局配置
        cfg = cls._global_config
        log_level = log_level or cfg.get("log_level", "INFO")
        log_dir = log_dir or cfg.get("path", "logs")
        fmt = fmt or cfg.get("fmt", "%Y_%m_%d")
        retention_days = retention_days or cfg.get("retention_days", 30)
        # 日志目录改为 main 程序根目录下 logs 文件夹
        if hasattr(sys, 'frozen'):
            # 打包后 exe 运行
            base_dir = os.path.dirname(sys.executable)
        else:
            # 普通脚本运行
            base_dir = os.path.dirname(os.path.abspath(sys.argv[0]))
        logs_dir = os.path.join(base_dir, log_dir)
        os.makedirs(logs_dir, exist_ok=True)
        now = datetime.now()
        date_prefix = now.strftime(fmt)
        # 查找当天已存在的日志文件
        existing_logs = [f for f in os.listdir(logs_dir) if f.startswith(date_prefix) and f.endswith(".log")]
        log_file_path = None
        max_size = 5 * 1024 * 1024  # 5MB

        if existing_logs:
            # 按修改时间排序，取最新的
            existing_logs.sort(key=lambda f: os.path.getmtime(os.path.join(logs_dir, f)), reverse=True)
            latest_log = os.path.join(logs_dir, existing_logs[0])
            # 判断是否超过5MB
            if os.path.getsize(latest_log) < max_size:
                log_file_path = latest_log
            else:
                # 超过5MB，新建一个
                file_name = now.strftime("%Y_%m_%d_%H_%M_%S.log")
                log_file_path = os.path.join(logs_dir, file_name)
        else:
            # 当天没有日志，新建
            file_name = now.strftime("%Y_%m_%d_%H_%M_%S.log")
            log_file_path = os.path.join(logs_dir, file_name)

        log_file = log_file_path

        if name in cls._loggers:
            logger = cls._loggers[name]
        else:
            logger = logging.getLogger(name)
            logger.setLevel(getattr(logging, log_level.upper(), logging.INFO))
            formatter = logging.Formatter(
                "%(asctime)s %(levelname)s %(name)s %(message)s"
            )
            file_handler = TimedRotatingFileHandler(
                log_file, when="midnight", backupCount=retention_days, encoding="utf-8"
            )
            file_handler.setFormatter(formatter)
            logger.addHandler(file_handler)
            # 可选：控制台输出
            stream_handler = logging.StreamHandler()
            stream_handler.setFormatter(formatter)
            logger.addHandler(stream_handler)
            cls._loggers[name] = logger
        return logger

    @classmethod
    def get_global_log_level(cls):
        """返回当前全局日志等级（字符串）"""
        return cls._global_config.get("log_level", "INFO")

    @classmethod
    def get_logger_level(cls, name: str = "default"):
        logger = cls._loggers.get(name)
        if logger:
            return logging.getLevelName(logger.level)
        return None