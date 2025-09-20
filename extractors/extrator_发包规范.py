# src/office_ops/ppt_processor/extractors/extractor_a.py
# 实现需求A：提取指定页面的标题等字段并导出为 发包规范
import os
import yaml
from typing import List, Dict
from extractors.base_extrator import BaseExtractor
from content_models import Slide, calculate_iou
# from config.filelds_config import FIELDS_CONFIG
from config.fields_loader import FIELDS_CONFIG
from utils.text_utils import split_after_colon  # 确保已导入
from utils.logger import LoggerFactory
import re
import math

class ExtractorA(BaseExtractor):
    """需求A提取器（支持多类型 shape 提取）"""
    def __init__(self, slides: List[Slide]):
        self.logger = LoggerFactory.create_logger("Extractor_发包规范")
        self.logger.info("初始化提取器")
        super().__init__(slides)
        self.config = FIELDS_CONFIG.get("发包规范V1", {})
        if not self.config:
            self.logger.warning("未找到发包规范V1的配置，请检查config/filelds_config.yaml文件")


    def _calc_utilization_rate(self, failure_rate: str) -> str:
            """
            根据故障率字符串自动计算利用率（百分比取反），如 '≤1.23%' -> '≥98.77%'
            """
            import re
            # 去掉所有空格
            failure_rate = failure_rate.replace(" ", "")
            match = re.search(r"([≤<]?)(\d+(\.\d+)?)%", failure_rate)
            if match:
                num = float(match.group(2))
                utilization = 100 - num
                # 保留两位小数
                return f"≥ {utilization:.2f}%"
            return ""

    def extract(self) -> Dict:
        self.logger.info("开始提取内容")
        try:
            flat_result = {}

            # 处理 master 字段（用 master_shapes 匹配）
            master_cfg = self.config.get("master", {})
            if master_cfg:
                for field, pos_dict in master_cfg.items():
                    for key, position in pos_dict.items():
                        # 遍历所有 slide，查找 master_shapes
                        for slide in self.slides:
                            for shape in slide.get("master_shapes", []):
                                iou = calculate_iou(shape["box"], position)
                                if iou > 0.3:
                                    # 只取第一个匹配
                                    flat_result[key] = shape.get("text", "")
                                    break

            # 按 page 定位
            for page_num, page_cfg in self.config.get("page", {}).items():
                slide = self._get_slide_by_page(int(page_num))
                if not slide:
                    continue
                for field, iou_cfg in page_cfg.get("iou", {}).items():
                    box = iou_cfg.get("box")
                    need_split = iou_cfg.get("need_split")
                    storage_vars = iou_cfg.get("storage_var")
                    # IOU 匹配 shape
                    for shape in slide["shapes"]:
                        iou = calculate_iou(shape["box"], box)
                        if iou > 0.3:
                            text = shape.get("text", "")
                            if need_split and isinstance(storage_vars, list):
                                parts = text.split(need_split)
                                for idx, var in enumerate(storage_vars):
                                    if var and idx < len(parts):
                                        flat_result[var] = parts[idx]
                            else:
                                flat_result[field] = text
                            break

            # 按 title 定位
            for title_cfg in self.config.get("title", []):
                slide = self._get_slide_by_title(title_cfg["first"], title_cfg.get("second", ""))
                if not slide:
                    continue
                # re 匹配文本字段
                for field, field_cfg in title_cfg.get("re", {}).items():
                    storage_var = field
                    flat_result[storage_var] = self._extract_text_by_re(slide, field, field_cfg)
                # table 匹配表格字段
                for field, field_cfg in title_cfg.get("table", {}).items():
                    storage_var = field
                    flat_result[storage_var] = self._extract_table_last_row(slide, field_cfg["match_key_string"])

            # 自动补充 dev_utilization_rate 字段
            failure_rate = flat_result.get("dev_failure_rate", "")
            if failure_rate and isinstance(failure_rate, str):
                flat_result["dev_utilization_rate"] = self._calc_utilization_rate(failure_rate)

            self.logger.info("内容提取完成")
            return flat_result
        except Exception as e:
            self.logger.error(f"提取内容时出错: {str(e)}", exc_info=True)

    def _get_slide_by_page(self, page_number: int):
        for slide in self.slides:
            if slide["page_number"] == page_number:
                return slide
        return None

    def _get_slide_by_title(self, first_title: str, second_title: str):
        for slide in self.slides:
            if slide["title"] == first_title and (not second_title or slide["second_title"] == second_title):
                return slide
        return None

    def _extract_shape_by_position(self, slide, position):
        """
        通过 IOU 判断形状位置是否匹配，提取内容
        position: (left, top, width, height)
        """
        iou_threshold = 0.3  # 可根据实际情况调整
        for shape in slide["shapes"]:
            iou = calculate_iou(shape["box"], position)
            if iou > iou_threshold:
                if shape["type"] == "TextBox" or shape["type"] == "文本框" or shape["type"] == "矩形":
                    return shape.get("text", "")
        return None

    def _extract_text_by_re(self, slide, field_name, field_cfg):
        """
        根据正则表达式提取文本内容或相关形状
        
        Args:
            slide: 幻灯片数据
            field_name: 字段名称
            field_cfg: 字段配置
            
        Returns:
            根据match_rule返回不同类型的数据:
            - match_rule > 0: 返回带页码的shape对象
            - match_rule = 0: 返回匹配的文本
            - match_rule = -1: 返回冒号后的提取值
        """
        match_rule = field_cfg.get("match_rule", 0)
        re_rule = field_cfg.get("re_rule", "")
        shapes = slide["shapes"]
        direction_map = {1: "down", 2: "left", 3: "up", 4: "right"}

        # 遍历所有shape
        for shape in shapes:
            # 处理群组
            if shape["type"] == "Group":
                result = self._process_group_shape(
                    shape, re_rule, match_rule, 
                    slide["page_number"], direction_map
                )
                if result:
                    return result
                    
            # 处理独立文本框
            elif shape["type"] in ["文本框", "矩形"]:
                result = self._process_single_shape(
                    shape, re_rule, match_rule,
                    shapes, slide["page_number"], direction_map
                )
                if result:
                    return result

        return None

    def _process_group_shape(self, group, re_rule, match_rule, page_number, direction_map):
        """处理群组内的shape"""
        for sub_shape in group["shapes"]:
            if sub_shape["type"] in ["文本框", "矩形"]:
                text = sub_shape.get("text", "")
                if re_rule and re.search(re_rule, text):
                    if match_rule > 0:
                        if match_rule == 5:
                            # 返回当前sub_shape本身和页码
                            sub_shape["page_number"] = page_number
                            return sub_shape
                        else:
                            # 原有的群组内搜索逻辑
                            for target in group["shapes"]:
                                if target["type"] in ["Image", "CustomShape", "Group", "矩形"]:
                                    target["page_number"] = page_number
                                    return target
                    elif match_rule == 0:
                        # 清理文本前后的空白字符
                        return text.strip() if text else ""
                    elif match_rule == -1:
                        _text = self._extract_value_after_colon(text, re_rule)
                        return _text.strip() if _text else ""
        return None

    def _process_single_shape(self, shape, re_rule, match_rule, all_shapes, page_number, direction_map):
        """处理单独的文本框"""
        text = shape.get("text", "")
        if re_rule and re.search(re_rule, text):
            if match_rule > 0:
                if match_rule == 5:
                    # 返回当前shape本身和页码
                    shape["page_number"] = page_number
                    return shape
                else:
                    # 原有的近邻搜索逻辑
                    direction = direction_map.get(match_rule, "down")
                    target = self._find_nearest_shape(shape, all_shapes, direction)
                    if target:
                        target["page_number"] = page_number
                        return target
            elif match_rule == 0:
                # 清理文本前后的空白字符
                return text.strip() if text else ""
            elif match_rule == -1:
                _text = self._extract_value_after_colon(text, re_rule)
                return _text.strip() if _text else ""
        return None

    def _extract_value_after_colon(self, text: str, re_rule: str) -> str:
        """提取冒号后的值"""
        regex = re.compile(re_rule)
        match = regex.search(text)
        if match and len(match.regs) > 0:
            temp = text[match.regs[0][0]:match.regs[0][1]]
            return split_after_colon(temp) or ""
        return ""

    def _find_nearest_shape(self, ref_shape, shapes, direction="down"):
        """
        查找与 ref_shape 最近的目标 shape（如图片或 custom shape），支持下方/左侧/上方/右侧
        """
        ref_box = ref_shape["box"]
        ref_center = (ref_box[0] + ref_box[2] / 2, ref_box[1] + ref_box[3] / 2)
        min_dist = float("inf")
        nearest = None
        for shape in shapes:
            if shape is ref_shape:
                continue
            # 目标类型可根据需求调整
            if shape["type"] in ["Image", "图片", "矩形", "Group"]:
                box = shape["box"]
                center = (box[0] + box[2] / 2, box[1] + box[3] / 2)
                # 按方向筛选
                if direction == "down" and center[1] > ref_center[1]:
                    dist = math.hypot(center[0] - ref_center[0], center[1] - ref_center[1])
                elif direction == "up" and center[1] < ref_center[1]:
                    dist = math.hypot(center[0] - ref_center[0], ref_center[1] - center[1])
                elif direction == "left" and center[0] < ref_center[0]:
                    dist = math.hypot(ref_center[0] - center[0], center[1] - ref_center[1])
                elif direction == "right" and center[0] > ref_center[0]:
                    dist = math.hypot(center[0] - ref_center[0], center[1] - ref_center[1])
                else:
                    continue
                if dist < min_dist:
                    min_dist = dist
                    nearest = shape
        return nearest

    def _extract_table_last_row(self, slide, field_name):
        for shape in slide["shapes"]:
            if shape["type"] == "Table":
                return shape["last_row"].get(field_name, "")
        return None

if __name__ == "__main__":
    from config.loader import ConfigLoader
    from content_models import Slide
    from pptx import Presentation
    
    # 测试提取器
    pptx_path = r"..\examples/templates/26xdemo2.pptx"
    config_loader = ConfigLoader(config_dir=r"..\config")
    
    # 读取PPT并结构化
    prs = Presentation(pptx_path)
    slides = []
    for page_number, slide in enumerate(prs.slides, start=1):
        slide_obj = Slide(slide, prs.slide_master, page_number, config_loader, "v1")
        slides.append(slide_obj)
    
    slides_dicts = [slide.to_dict() for slide in slides]
    
    # 测试提取
    extractor = ExtractorA(slides_dicts)
    result = extractor.extract()
    print("字段提取完成")
    print(f"提取结果: {result}")