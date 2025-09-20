# 抽象提取器
# src/office_ops/ppt_processor/extractors/base_extractor.py
from abc import ABC, abstractmethod
from typing import List, Dict
import sys
import os
print("当前模块路径:", __file__)          # 输出：D:\pythonf\office_ops\ppt_processor\extractors\base_extractor.py
print("Python 搜索路径:", sys.path)       # 检查是否包含 project_root（D:\pythonf）
print("当前模块所属包:", __package__)     # 正确应为：office_ops.ppt_processor.extractors
from content_models import Slide  # 绝对导入

class BaseExtractor(ABC):
    """抽象提取器基类"""
    def __init__(self, slides: List[Slide]):
        self.slides = slides  # 输入的结构化幻灯片数据

    @abstractmethod
    def extract(self) -> Dict:
        """提取逻辑（由子类实现）"""
        pass

