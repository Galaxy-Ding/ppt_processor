import logging
from logging.handlers import TimedRotatingFileHandler
import os
from typing import Optional, Dict, Any
from datetime import datetime

class LoggerFactory:
    """增强版日志工厂，支持GUI显示和文件记录"""
    
    _loggers = {}
    
    @classmethod
    def create_logger(
        cls, 
        name: str,
        log_level: str = "INFO",
        log_file: Optional[str] = None,
        retention_days: int = 30,
        gui_display: Optional[object] = None
    ) -> logging.Logger:
        """创建或获取日志记录器
        
        Args:
            name: 日志记录器名称
            log_level: 日志级别 (DEBUG/INFO/WARNING/ERROR)
            log_file: 日志文件路径
            retention_days: 日志保留天数
            gui_display: GUI显示对象，需有log_message方法
        """
        if name in cls._loggers:
            return cls._loggers[name]
            
        logger = logging.getLogger(name)
        logger.setLevel(log_level)
        
        # 清除现有处理器
        for handler in logger.handlers[:]:
            logger.removeHandler(handler)
        
        # 控制台输出
        console_handler = logging.StreamHandler()
        console_handler.setLevel(log_level)
        console_formatter = logging.Formatter(
            '%(asctime)s - %(name)s - %(levelname)s - %(message)s')
        console_handler.setFormatter(console_formatter)
        logger.addHandler(console_handler)
        
        # 文件输出
        if log_file:
            os.makedirs(os.path.dirname(log_file), exist_ok=True)
            file_handler = TimedRotatingFileHandler(
                log_file,
                when='midnight',
                backupCount=retention_days
            )
            file_handler.setLevel(log_level)
            file_formatter = logging.Formatter(
                '%(asctime)s - %(name)s - %(levelname)s - %(message)s')
            file_handler.setFormatter(file_formatter)
            logger.addHandler(file_handler)
        
        # 添加GUI显示处理器
        if gui_display:
            class GuiLogHandler(logging.Handler):
                def emit(self, record):
                    msg = self.format(record)
                    gui_display.log_message(msg, record.levelname)
                    
            gui_handler = GuiLogHandler()
            gui_handler.setLevel(log_level)
            logger.addHandler(gui_handler)
        
        cls._loggers[name] = logger
        return logger

    @classmethod
    def update_log_level(cls, name: str, level: str):
        """更新日志记录器级别"""
        if name in cls._loggers:
            logger = cls._loggers[name]
            logger.setLevel(level)
            for handler in logger.handlers:
                handler.setLevel(level)

class ProgressLogger:
    """带进度记录的日志记录器"""
    
    def __init__(self, logger: logging.Logger):
        self.logger = logger
        self.progress = 0
        self.max_progress = 100
        
    def set_progress(self, value: int, max_value: int = None):
        """设置当前进度"""
        self.progress = value
        if max_value is not None:
            self.max_progress = max_value
        self.logger.info(f"进度: {value}/{self.max_progress}")
        
    def increment(self, step: int = 1):
        """增加进度"""
        self.progress += step
        self.logger.info(f"进度: {self.progress}/{self.max_progress}")