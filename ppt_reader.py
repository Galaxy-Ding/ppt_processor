# PPT 读取模块，负责读取每页内容，返回结构化的数据
from pptx import Presentation
from content_models import Slide
from config.loader import ConfigLoader
from typing import Optional, List
from extractors.extrator_发包规范 import ExtractorA    
# 4. 导出图片并生成文档
from exporters.exporter_发包规范 import ExporterA

def read_pptx(pptx_path: str, config_loader: ConfigLoader, version: Optional[str] = None) -> List[Slide]:
    """读取 PPT 文件并返回结构化的幻灯片列表（支持动态标题位置）"""
    prs = Presentation(pptx_path)
    slides = []
    for page_number, slide in enumerate(prs.slides, start=1):  # 页码从1开始
        # 传递 config_loader 和 version 到 Slide 实例
        slide_obj = Slide(slide, prs.slide_master, page_number, config_loader, version)
        slides.append(slide_obj)
    return slides

if __name__ == "__main__":
    # 测试读取逻辑
    pptx_path = r"D:\pythonf\26xdemo2.pptx"
    docx_path = r"D:\pythonf\26xdemo1.docx"
    output_path = r"D:\pythonf\output.docx"
    version = "v1"

    # 1. 初始化配置加载器
    config_loader = ConfigLoader(config_dir="config")

    # 2. 读取 PPT 并结构化
    slides = read_pptx(pptx_path, config_loader, version)

    slides_dicts = []
    for slide in slides:
        slides_dicts.append(slide.to_dict())

    # 3. 提取需要的信息
    extractor = ExtractorA(slides_dicts)
    result = extractor.extract()
    print("\n发包规范字段提取结果:")
    print(result)

    # 4. 导出图片并生成文档
    # 创建导出器实例
    exporter = ExporterA(pptx_path, docx_path, output_path)
    # 需要从excel 读取的数据，暂时无特定的表格
    result["name"] = "Demo测试机器253改247阿德撒旦_adasadjqh"
    result["Action"] = "大改造"
    success = exporter.process(result)
    
    if success:
        print(f"\n文档已成功生成: {output_path}")
    else:
        print("\n文档生成失败")





