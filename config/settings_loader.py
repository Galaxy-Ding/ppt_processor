import os
import yaml
from typing import Dict, Any
from utils.logger import LoggerFactory

class SettingsLoader:
    def __init__(self, config_dir: str = "config"):
        self.logger = LoggerFactory.create_logger("SettingsLoader")
        self.config_dir = config_dir
        self.settings = None
        self._load_settings()
    
    def _load_settings(self) -> None:
        """加载YAML配置文件"""
        try:
            config_path = os.path.join(self.config_dir, "app_settings.yaml")
            self.logger.debug(f"加载配置文件: {config_path}")
            
            with open(config_path, 'r', encoding='utf-8') as f:
                self.settings = yaml.safe_load(f)
                
            self.logger.info("配置加载成功")
        except Exception as e:
            self.logger.error(f"加载配置文件失败: {e}", exc_info=True)
            raise
    
    def get_folder_level(self) -> int:
        """获取目录层级"""
        return self.settings["folder_structure"]["level"]
    
    def get_path_config(self, level: int) -> Dict[str, Any]:
        """获取指定层级的路径配置"""
        return self.settings["folder_structure"]["paths"][str(level)]
    
    def get_template_config(self) -> Dict[str, str]:
        """获取模板配置"""
        return self.settings["templates"]["docx"]
    
    def get_log_config(self) -> Dict[str, Any]:
        """获取日志配置"""
        return self.settings["logs"]