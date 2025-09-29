import os
import re
import glob
from pptx import Presentation
from content_models import Slide
from typing import List, Dict, Optional
from config.loader import ConfigLoader
from utils.logger import LoggerFactory, LOG_LEVELS
from utils.text_utils import traditional_to_simplified
from extractors.extrator_发包规范 import ExtractorA
from exporters.exporter_发包规范 import ExporterA
import traceback



class PackingFileProcessor:
    def set_log_callback(self, callback):
        self._external_log_callback = callback
    """发包规范 处理核心类"""
    
    def __init__(self, config: ConfigLoader, log_level: str = "INFO"):
        self.config = config
        self.logger = LoggerFactory.create_logger("PackingFile（发包规范）Processor")
        self.log_level = log_level
        
    def _log(self, message: str, level: str = "INFO"):
        # 记录到logger
        if level == "DEBUG":
            self.logger.debug(message)
        elif level == "INFO":
            self.logger.info(message)
        elif level == "WARNING":
            self.logger.warning(message)
        elif level == "ERROR":
            self.logger.error(message)
        # 只在等级大于等于当前设置时输出到UI
        if hasattr(self, '_external_log_callback') and self._external_log_callback:
            if LOG_LEVELS[level] >= LOG_LEVELS.get(self.log_level, 20):
                self._external_log_callback(f"[{level}] {message}")

    def change_log_level(self, log_level: str):
        self.log_level = log_level
        

    def get_v3_pptx_directories(self, root_dir: str) -> list[str]:
        """
        处理指定根目录下的所有目录，查找包含pptx文件的目录，并返回包含最大版本v3 PPT文件的目录列表。

        Args:
        root_dir (str): 根目录路径。

        Returns:
        list[str]: 包含最大版本v3 PPT文件的目录列表。

        """
        results = []
        exclude_dirs = {"result", "temp", "config", "__pycache__"}
        # 1. 用glob递归查找所有pptx文件
        pptx_files = glob.glob(os.path.join(root_dir, "**", "*.pptx"), recursive=True)
        # 2. 收集所有包含pptx的目录
        all_dirs = set()
        for pptx_file in pptx_files:
            dir_path = os.path.dirname(pptx_file)
            # 检查是否包含排除目录
            parts = set(os.path.normpath(dir_path).split(os.sep))
            if not (parts & exclude_dirs):
                all_dirs.add(dir_path)
        # 3. 对每个目录，查找最大版本的v3 PPT
        for dir_path in all_dirs:
            v3_ppt_files = [f for f in os.listdir(dir_path) if f.lower().endswith('.pptx') and 'v3' in f.lower()]
            if not v3_ppt_files:
                continue
            max_version = float('-inf')
            max_file = None
            for file in v3_ppt_files:
                version = self._get_ppt_version(file)
                if version is not None and version > max_version:
                    max_version = version
                    max_file = file
            if max_file:
                ppt_file = os.path.join(dir_path, max_file)
                results.append(ppt_file)
        return results
    

    def process_generate_reports(self, selected_dir: str, data_list=None, manual_proj_name_value=None,
                                 manual_proj_type_value=None):
        """
        UI - 调用接口处理多个PPT文件，生成对应的发包规范文档。

        Args:
        selected_dir (str): 选定的目录，包含多个PPT文件。
        data_list (list[dict], optional): 从Excel读取的结构化数据，默认为None。
        manual_proj_name_value (str, optional): 手动输入的工程名字，默认为None。
        manual_proj_type_value (str, optional): 手动输入的工程类型，默认为None。

        Returns:
        dict: 包含每个PPT文件的处理结果，键为PPT文件路径，值为生成的发包规范文档路径。

        """
        result_dirs = {}
        pptx_paths = self.get_v3_pptx_directories(selected_dir)

        for pptx_path in pptx_paths:
            try:
                dir_path = os.path.dirname(pptx_path)
                result_dir = self._create_result_dir(dir_path)
                output_path = self._process_single_ppt(
                    pptx_path, result_dir,
                    data_list=data_list,
                    manual_proj_name_value=manual_proj_name_value,
                    manual_proj_type_value=manual_proj_type_value
                )
            except Exception as e:
                self._log(f"处理文件 {pptx_path} 失败: {str(e)}\n{traceback.format_exc()}", level="ERROR")
            result_dirs[pptx_path] = output_path
        return result_dirs

    def process_ppt_to_docx(self, pptx_path: str, docx_path: str, output_path: str, version: str = "v1",
                           data_list=None, manual_proj_name_value=None, manual_proj_type_value=None):
        """
        应用层统一入口：处理PPT，提取信息并生成文档
        :param pptx_path: PPTX文件路径
        :param docx_path: DOCX模板路径
        :param output_path: 输出DOCX路径
        :param version: 版本号
        :param data_list: Excel结构化数据
        :param manual_proj_name_value: 手动工程名字
        :param manual_proj_type_value: 手动工程类型
        """
        # 读取PPT并结构化
        slides = self._read_pptx(pptx_path, version)
        slides_dicts = [slide.to_dict() for slide in slides]
        # 提取需要的信息
        extractor = ExtractorA(slides_dicts, self.config)
        result = extractor.extract()
        self.logger.debug("\n发包规范字段提取结果:")
        self.logger.debug(str(result))

        # Excel数据匹配逻辑
        project_code = result.get("ProjectCode")
        matched_scheme_name = None
        if data_list and project_code:
            for item in data_list:
                if str(item.get("ProjectCode", "")).strip() == str(project_code).strip():
                    matched_scheme_name = item.get("name")
                    matched_scheme_type = item.get("Action")
                    break
        # 匹配成功
        if matched_scheme_name:
            self.logger.debug(f"{project_code} 匹配到 方案總表的方案代碼\n")
            result["name"] = matched_scheme_name
            result["Action"] = matched_scheme_type
            self.last_matched_name = matched_scheme_name
            self.last_matched_type = matched_scheme_type
        else:
            # 未匹配到
            self.logger.debug(f"{project_code} 采用 手動的名稱和Action\n")
            result["name"], result["Action"] = "", ""
            self.last_matched_name = "未匹配到"
            self.last_matched_type = "未匹配到"
            # 手动输入覆盖
            if manual_proj_name_value:
                result["name"] = manual_proj_name_value
            if manual_proj_type_value:
                result["Action"] = manual_proj_type_value


        if traditional_to_simplified(result.get("Action")) in ["新制","研发","大改造", "大改"]:
            result["InstalledDate"] = "21"
        elif traditional_to_simplified(result.get("Action")) in ["再制", "中改造", "中改"]:
            result["InstalledDate"] = "14"
        elif traditional_to_simplified(result.get("Action")) in ["小改造","小改"]:
            result["InstalledDate"] = "7"

        if traditional_to_simplified(result.get("Action")) in ["新制","研发","大改造", "大改"]:
            result["LQStartDate"] = "21"
        elif traditional_to_simplified(result.get("Action")) in ["再制", "中改造", "中改"]:
            result["LQStartDate"] = "14"
        elif traditional_to_simplified(result.get("Action")) in ["小改造","小改"]:
            result["LQStartDate"] = "7"
        # OSSDate 固定为0
        result["OSSDate"] = "0"

        exporter = ExporterA(pptx_path, docx_path, output_path)
        success = exporter.process(result)
        if success:
            self._log(f"\n文档已成功生成: {output_path}", level="INFO")
        else:
            self._log("\n文档生成失败", level="INFO")

    def _create_result_dir(self, base_dir: str) -> str:
        """创建结果目录"""
        result_base = "result"
        counter = 0
        
        while True:
            if counter == 0:
                result_dir = os.path.join(base_dir, result_base)
            else:
                result_dir = os.path.join(base_dir, f"{result_base}_{counter}")
                
            if not os.path.exists(result_dir):
                os.makedirs(result_dir)
                self._log(f"创建结果目录: {result_dir}", level="INFO")
                return result_dir
                
            counter += 1
    
    def _get_ppt_version(self, filename: str) -> Optional[float]:
        """从PPT文件名中提取并验证版本号"""
        v3_pattern = re.compile(r'[vV](\d+\.?\d?)')
        match = v3_pattern.search(filename)
        return float(match.group(1)) if match else None

    def _find_v3_ppt_files(self, dir_path: str) -> str:
        """查找V3版本的PPT文件并返回最大版本号的文件路径"""
        max_version = float('-inf')
        max_version_file = None
        
        # 第一层：遍历目录
        for root, _, files in os.walk(dir_path):
            # 第二层：处理文件
            for file in files:
                if not file.lower().endswith('.pptx'):
                    continue
                    
                version = self._get_ppt_version(file)
                if version and version > max_version:
                    max_version = version
                    max_version_file = os.path.join(root, file)
                    self._log(f"找到更高版本PPT文件: {max_version_file}, 版本: {version}", level="DEBUG")
                    
        return max_version_file

    def _read_pptx(self, pptx_path: str, version: Optional[str] = None) -> List[Slide]:
        """
        读取PPT文件，返回结构化Slide对象列表
        """
        self._log(f"开始读取PPT文件: {pptx_path}", level="INFO")
        try:
            prs = Presentation(pptx_path)
            slides = []
            for page_number, slide in enumerate(prs.slides, start=1):
                self._log(f"处理第 {page_number} 页", level="DEBUG")
                slide_obj = Slide(slide, prs.slide_master, page_number, self.config, version)
                slides.append(slide_obj)
            self._log(f"成功读取 {len(slides)} 页", level="DEBUG")
            return slides
        except Exception as e:
            self._log(f"读取PPT文件时出错: {str(e)}\n{traceback.format_exc()}", level="ERROR")
            raise
    
    def _process_single_ppt(self, ppt_path: str, output_dir: str, data_list=None, manual_proj_name_value=None, manual_proj_type_value=None) -> str:
        """处理单个PPT文件"""
        self._log(f"开始处理文件: {ppt_path}", level="INFO")
        
        # 提取文件名（不含扩展名）
        filename = os.path.basename(ppt_path)
        base_name = os.path.splitext(filename)[0]
        
        # 提取 V3 前面的字符串
        v3_pattern = re.compile(r'^(.*?)[vV]3')
        match = v3_pattern.search(base_name)
        prefix = match.group(1).strip('_') if match else base_name
        
        # 生成输出文件名
        output_filename = f"{prefix}_v1_发包规范.docx"
        output_path = os.path.join(output_dir, output_filename)
        self.process_ppt_to_docx(
            ppt_path, self.get_template_path(), output_path, version="v1",
            data_list=data_list,
            manual_proj_name_value=manual_proj_name_value,
            manual_proj_type_value=manual_proj_type_value
        )
       
        self._log(f"处理完成，输出到: {output_path}", level="INFO")
        return output_path

    def get_template_path(self) -> str:
        """获取模板文件路径"""
        return self.config.get_template_path()