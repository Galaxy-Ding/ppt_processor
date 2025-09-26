import os
import re
import glob
from pptx import Presentation
from content_models import Slide
from typing import List, Dict, Optional
from config.loader import ConfigLoader
from utils.logger import LoggerFactory
from extractors.extrator_发包规范 import ExtractorA
from exporters.exporter_发包规范 import ExporterA
import traceback

LOG_LEVELS = {"DEBUG": 10, "INFO": 20, "WARNING": 30, "ERROR": 40}

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
    

    def process_generate_reports(self, selected_dir: str):
        """
        处理多个PPT文件，生成对应的发包规范文档
        :param pptx_paths: 包含最大版本v3 PPT文件的路径列表
        """
        result_dirs = {}
        pptx_paths = self.get_v3_pptx_directories(selected_dir)

        for pptx_path in pptx_paths:
            try:
                dir_path = os.path.dirname(pptx_path)
                result_dir = self._create_result_dir(dir_path)
                output_path = self._process_single_ppt(pptx_path, result_dir)
            except Exception as e:
                self._log(f"处理文件 {pptx_path} 失败: {str(e)}\n{traceback.format_exc()}", level="ERROR")
            result_dirs[pptx_path] = output_path
        return result_dirs
    
    def process_ppt_to_docx(self, pptx_path: str, docx_path: str, output_path: str, version: str = "v1"):
        """
        应用层统一入口：处理PPT，提取信息并生成文档
        :param pptx_path: PPTX文件路径
        :param docx_path: DOCX模板路径
        :param output_path: 输出DOCX路径
        :param version: 版本号
        """
        # 读取PPT并结构化
        slides = self._read_pptx(pptx_path, version)
        slides_dicts = [slide.to_dict() for slide in slides]
        # 提取需要的信息
        extractor = ExtractorA(slides_dicts, self.config)
        result = extractor.extract()
        self.logger.debug("\n发包规范字段提取结果:")
        self.logger.debug(str(result))
        # 导出图片并生成文档
        exporter = ExporterA(pptx_path, docx_path, output_path)
        # 需要从excel读取的数据，暂时无特定的表格
        result["name"] = "avbaaa"
        result["Action"] = "大改造"
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
    
    def _process_single_ppt(self, ppt_path: str, output_dir: str) -> str:
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
        self.process_ppt_to_docx(ppt_path, self.get_template_path(), output_path, version="v1")
       
        self._log(f"处理完成，输出到: {output_path}", level="INFO")
        return output_path

    def get_template_path(self) -> str:
        """获取模板文件路径"""
        return self.config.get_template_path()