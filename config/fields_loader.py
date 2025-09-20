import yaml
from pathlib import Path
from typing import Dict, Any

class FieldsConfigLoader:
    """专门用于加载字段配置的加载器"""
    
    def __init__(self, config_dir: str = "config"):
        self.config_dir = Path(config_dir)
        self.fields_config: Dict[str, Any] = {}
        self._load_config()

    def _load_config(self) -> None:
        """加载filelds_config.yaml配置文件"""
        config_path = self.config_dir / "filelds_config.yaml"
        if not config_path.exists():
            raise FileNotFoundError(f"字段配置文件不存在：{config_path}")
            
        with open(config_path, "r", encoding="utf-8") as f:
            self.fields_config = yaml.safe_load(f)

    def get_config(self, config_name: str) -> Dict[str, Any]:
        """获取指定名称的字段配置"""
        return self.fields_config.get("FIELDS_CONFIG", {}).get(config_name, {})

# 全局字段配置加载器实例
fields_loader = FieldsConfigLoader()
FIELDS_CONFIG = fields_loader.fields_config.get("FIELDS_CONFIG", {})