import os
import yaml
from typing import Dict, Any
from utils.logger import LoggerFactory, LOG_LEVELS
from pathlib import Path

class ConfigLoader:
    """增强版配置加载器，支持动态配置更新"""
    
    def __init__(self, config_dir: str = "config"):
        self.config_dir = config_dir
        # 1. 加载所有配置文件
        self.config = self._load_all_configs()
        # 2. 日志配置优先从 app_settings.yaml 读取
        log_config = self.config.get("logs", {})
        log_level = log_config.get("level", "INFO")
        log_dir = log_config.get("log_dir", "logs")
        fmt = log_config.get("format", "%Y_%m_%d")
        retention_days = log_config.get("retention_days", 30)
         # 全局设置一次
        LoggerFactory.set_global_config(
            log_level=log_level,
            log_dir=log_dir,
            fmt=fmt,
            retention_days=retention_days
        )
        self.logger = LoggerFactory.create_logger("通用loader")

    def _load_all_configs(self) -> Dict[str, Any]:
        """加载所有配置文件"""
        configs = {}
        # 加载主配置文件
        main_config = os.path.join(self.config_dir, "app_settings.yaml")
        if os.path.exists(main_config):
            with open(main_config, 'r', encoding='utf-8') as f:
                configs.update(yaml.safe_load(f))
        # 加载字段配置
        fields_config = os.path.join(self.config_dir, "filelds_config.yaml")
        if os.path.exists(fields_config):
            with open(fields_config, 'r', encoding='utf-8') as f:
                configs.update(yaml.safe_load(f))
        # 加载title配置
        title_config = os.path.join(self.config_dir, "title_positions.yaml")
        if os.path.exists(title_config):
            with open(title_config, 'r', encoding='utf-8') as f:
                configs.update(yaml.safe_load(f))
        return configs

    def get_template_path(self) -> str:
        """获取模板文件路径"""
        templates = self.config.get('templates', {})
        docx_config = templates.get('docx', {})
        docx_path = docx_config.get('path', 'templates')
        filename = docx_config.get('filename', '')
        return os.path.join(docx_path, filename)

    def get_project_excel_path(self) -> str:
        """获取总表位置"""
        templates = self.config.get('templates', {})
        excel_config = templates.get('excel', {})
        excel_path = excel_config.get('path', 'templates')
        filename = excel_config.get('filename', '')
        return os.path.join(excel_path, filename)

    def get_all_projects_info(self) -> dict:
        """获取方案总表的字典信息"""
        templates = self.config.get('templates', {})
        return templates.get('excel', {})

    def get_log_config(self) -> Dict[str, Any]:
        """获取日志配置"""
        return self.config.get('logs', {})

    def save_current_config(self, config_path: str = None):
        """保存当前配置到文件"""
        if not config_path:
            config_path = os.path.join(self.config_dir, "current_settings.yaml")
        with open(config_path, 'w', encoding='utf-8') as f:
            yaml.safe_dump(self.config, f)

    def get_title_position(self, version: str = "v1") -> Dict[str, Any]:
        """获取标题配置"""
        return self.config.get("versions", {}).get(version, {})

    def update_config(self, key: str, value: Any):
        """更新配置项"""
        keys = key.split('.')
        current = self.config
        for k in keys[:-1]:
            if k not in current:
                current[k] = {}
            current = current[k]
        current[keys[-1]] = value
