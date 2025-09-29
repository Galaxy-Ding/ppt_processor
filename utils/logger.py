import logging
from logging.handlers import TimedRotatingFileHandler
from datetime import datetime
import os

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

    @classmethod
    def create_logger(
        cls,
        name: str = "default",
        log_level: str = "INFO",
        log_file: str = None,
        retention_days: int = 30,
        gui_display: object = None
    ) -> logging.Logger:
        # 日志文件路径自动按日期命名
        logs_dir = os.path.join(os.path.dirname(__file__), "..", "logs")
        os.makedirs(logs_dir, exist_ok=True)
        now = datetime.now()
        date_prefix = now.strftime("%Y_%m_%d")
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

        log_file = log_file or log_file_path

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