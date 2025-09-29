import os
import yaml
from typing import Dict, Any
from utils.logger import LoggerFactory, LOG_LEVELS

class ConfigLoader:
    """增强版配置加载器，支持动态配置更新"""
    
    def __init__(self, config_dir: str = "config"):
        self.config_dir = config_dir
        self.config = self._load_all_configs()
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
        docx_path = templates.get('docx', {}).get('path', 'templates')
        filename = templates.get('docx', {}).get('filename', '')
        return os.path.join(docx_path, filename)
    
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
        """获取日志配置"""
        return self.config.get("versions").get(version, {})
            
    def update_config(self, key: str, value: Any):
        """更新配置项"""
        keys = key.split('.')
        current = self.config
        for k in keys[:-1]:
            if k not in current:
                current[k] = {}
            current = current[k]
        current[keys[-1]] = value