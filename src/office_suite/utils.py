
import os
import subprocess
import shutil
from typing import Optional, Dict, Any

def convert_with_libreoffice(input_path: str, target_format: str, output_path: str, **kwargs) -> Dict[str, Any]:
    """
    使用LibreOffice进行格式转换，支持大部分Office格式互转
    Args:
        input_path: 输入文件路径
        target_format: 目标格式扩展名，比如pdf、docx、pptx等
        output_path: 输出文件路径
    """
    # 检查LibreOffice是否安装
    libreoffice_path = shutil.which('libreoffice') or shutil.which('soffice')
    if not libreoffice_path:
        raise RuntimeError("LibreOffice not found, please install it first. On Ubuntu: sudo apt install libreoffice")
    
    output_dir = os.path.dirname(output_path)
    output_name = os.path.basename(output_path)
    
    # 构建命令
    cmd = [
        libreoffice_path,
        "--headless",
        "--convert-to", target_format,
        "--outdir", output_dir,
        input_path
    ]
    
    try:
        result = subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            check=True,
            timeout=kwargs.get("timeout", 300)
        )
        
        # 检查是否自动生成的文件名可能和预期不一致，重命名为指定的output_path
        expected_generated_path = os.path.join(
            output_dir,
            f"{os.path.splitext(os.path.basename(input_path))[0]}.{target_format}"
        )
        
        if os.path.exists(expected_generated_path) and expected_generated_path != output_path:
            os.rename(expected_generated_path, output_path)
        
        return {"success": True, "output_path": output_path}
        
    except subprocess.CalledProcessError as e:
        raise RuntimeError(f"LibreOffice转换失败: {e.stderr}")

def get_file_type(file_path: str) -> str:
    """
    获取文件类型，返回扩展名（小写）
    """
    if not os.path.exists(file_path):
        return ""
    return os.path.splitext(file_path)[1].lower().lstrip('.')

def is_office_file(file_path: str) -> bool:
    """
    判断是否是支持的Office文件类型
    """
    supported_extensions = {"docx", "doc", "xlsx", "xls", "pptx", "ppt", "pdf"}
    ext = get_file_type(file_path)
    return ext in supported_extensions

def get_file_size(file_path: str, unit: str = "bytes") -> float:
    """
    获取文件大小，支持单位：bytes/KB/MB/GB
    """
    if not os.path.exists(file_path):
        return 0
        
    size_bytes = os.path.getsize(file_path)
    
    units = {
        "bytes": 1,
        "kb": 1024,
        "mb": 1024*1024,
        "gb": 1024*1024*1024
    }
    
    unit = unit.lower()
    return round(round(size_bytes / units.get(unit, 1), 2)

def ensure_dir(path: str) -> None:
    """
    确保目录存在，不存在则创建
    """
    os.makedirs(path, exist_ok=True)

def clean_temp_files(temp_dir: str, older_than_hours: int = 24) -> int:
    """
    清理临时文件，删除指定目录下超过指定时间的文件
    """
    import time
    now = time.time()
    deleted_count = 0
    
    if not os.path.exists(temp_dir):
        return 0
        
    for filename in os.listdir(temp_dir):
        file_path = os.path.join(temp_dir, filename)
        if os.path.isfile(file_path):
            if (now - os.path.getmtime(file_path)) > older_than_hours * 3600:
                os.remove(file_path)
                deleted_count += 1
                
    return deleted_count
