# src/office_ops/config/loader.py
import yaml
from pathlib import Path
from typing import Dict, Any, Optional, List


class ConfigLoader:
    """配置文件加载器（支持YAML）"""

    def __init__(self, config_dir: str = "config"):
        self.config_dir = Path(config_dir)
        self.title_positions: Dict[str, Dict[str, List[float]]] = {}  # 版本→标题位置映射
        self._load_configs()

    def _load_configs(self) -> None:
        """加载所有配置文件（当前仅需标题位置）"""
        title_config_path = self.config_dir / "title_positions.yaml"
        if not title_config_path.exists():
            raise FileNotFoundError(f"标题位置配置文件不存在：{title_config_path}")

        with open(title_config_path, "r", encoding="utf-8") as f:
            config_data = yaml.safe_load(f)
            self.title_positions = config_data.get("versions", {})

    def get_title_position(self, version: Optional[str] = None) -> Dict[str, List[float]]:
        """
        获取指定版本的标题位置配置
        :param version: 版本号（如"v1"），未指定时使用默认版本
        :return: 标题位置字典（{"title": [...], "second_title": [...]}）
        """
        # 确定目标版本
        target_version = version or self.title_positions.get("default", "v1")
        if target_version not in self.title_positions:
            raise ValueError(f"未支持的版本号：{target_version}，可用版本：{list(self.title_positions.keys())}")

        return self.title_positions[target_version]


if __name__ == "__main__":
    # 测试配置加载
    config_loader = ConfigLoader(config_dir="config")
    print("配置加载结果:")
    print(config_loader.get_config("发包规范V1"))