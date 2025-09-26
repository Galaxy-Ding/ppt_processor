import os
import sys
import shutil
from pathlib import Path

RELATIVE_CFG = os.path.join('config', 'fields_config.yaml')

def resolve_config_path() -> str:
    """解析配置文件路径，支持多种环境"""
    # 1. 尝试在打包后的exe同级目录查找
    if getattr(sys, 'frozen', False):
        exe_dir = Path(sys.executable).parent
        exe_cfg = exe_dir / RELATIVE_CFG
        if exe_cfg.exists():
            return str(exe_cfg)
    
    # 2. 尝试在项目根目录查找
    project_root = Path(__file__).parent.parent
    project_cfg = project_root / RELATIVE_CFG
    if project_cfg.exists():
        return str(project_cfg)
        
    # 3. 尝试从PyInstaller打包资源中复制
    try:
        if hasattr(sys, '_MEIPASS'):
            internal_cfg = Path(sys._MEIPASS) / RELATIVE_CFG
            if internal_cfg.exists():
                os.makedirs(project_root / 'config', exist_ok=True)
                shutil.copy2(str(internal_cfg), str(project_cfg))
                return str(project_cfg)
    except Exception:
        pass
        
    # 返回默认期望路径
    return str(project_cfg)