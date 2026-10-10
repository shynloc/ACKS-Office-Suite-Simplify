
import glob
import os
import re
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path
from typing import Optional, Dict, Any, List


def data_dir() -> Path:
    """本工具的用户数据目录（下载的字体等放在这里）；可用环境变量 ACKS_OFFICE_HOME 覆盖。"""
    override = os.environ.get("ACKS_OFFICE_HOME")
    if override:
        return Path(override).expanduser()
    if sys.platform == "darwin":
        return Path.home() / "Library" / "Application Support" / "acks-office"
    if os.name == "nt":
        return Path(os.environ.get("LOCALAPPDATA") or Path.home() / "AppData" / "Local") / "acks-office"
    return Path(os.environ.get("XDG_DATA_HOME") or Path.home() / ".local" / "share") / "acks-office"


def soffice_candidates() -> List[str]:
    """各系统上 LibreOffice 的标准安装位置（不含 PATH）。"""
    if sys.platform == "darwin":
        return ["/Applications/LibreOffice.app/Contents/MacOS/soffice",
                os.path.expanduser("~/Applications/LibreOffice.app/Contents/MacOS/soffice")]
    if os.name == "nt":
        bases = [os.environ.get("ProgramFiles", r"C:\Program Files"),
                 os.environ.get("ProgramFiles(x86)", r"C:\Program Files (x86)")]
        return [os.path.join(b, "LibreOffice", "program", "soffice.exe") for b in bases if b]
    return (["/usr/bin/soffice", "/usr/lib/libreoffice/program/soffice",
             "/opt/libreoffice/program/soffice", "/snap/bin/libreoffice"]
            + sorted(glob.glob("/opt/libreoffice*/program/soffice"), reverse=True))


def find_soffice() -> Optional[str]:
    """找到可执行的 LibreOffice：环境变量 ACKS_OFFICE_SOFFICE → PATH → 各系统标准安装位置。"""
    override = os.environ.get("ACKS_OFFICE_SOFFICE")
    if override:
        return override if os.path.isfile(override) else None
    for name in ("soffice", "libreoffice"):
        found = shutil.which(name)
        if found:
            return found
    for path in soffice_candidates():
        if os.path.isfile(path):
            return path
    return None


def _seed_profile(profile: Path) -> int:
    """macOS 上的 LibreOffice 只认本地化后的字体名（中文系统里是「苹方-简」，不认「PingFang SC」）：
    把英文名 → 本地化名写进这次转换用的临时配置的字体替换表。返回写入的条数。"""
    try:
        from .fonts import localized_family_names
        pairs = localized_family_names()
    except Exception:
        pairs = {}
    if not pairs:
        return 0
    from xml.sax.saxutils import escape
    invalid = re.compile(r"[\x00-\x08\x0b\x0c\x0e-\x1f\ufffe\uffff\ud800-\udfff]")
    pairs = {a: b for a, b in pairs.items() if not invalid.search(a + b)}
    path = "/org.openoffice.Office.Common/Font/Substitution"
    items = [f'<item oor:path="{path}"><prop oor:name="Replacement" oor:op="fuse"><value>true</value></prop></item>']
    for i, (name, local) in enumerate(sorted(pairs.items())):
        items.append(
            f'<item oor:path="{path}/FontPairs"><node oor:name="_{i}" oor:op="replace">'
            '<prop oor:name="Always" oor:op="fuse"><value>true</value></prop>'
            '<prop oor:name="OnScreenOnly" oor:op="fuse"><value>false</value></prop>'
            f'<prop oor:name="ReplaceFont" oor:op="fuse"><value>{escape(name)}</value></prop>'
            f'<prop oor:name="SubstituteFont" oor:op="fuse"><value>{escape(local)}</value></prop>'
            '</node></item>')
    user = profile / "user"
    user.mkdir(parents=True, exist_ok=True)
    (user / "registrymodifications.xcu").write_text(
        '<?xml version="1.0" encoding="UTF-8"?>\n'
        '<oor:items xmlns:oor="http://openoffice.org/2001/registry" xmlns:xs="http://www.w3.org/2001/XMLSchema" '
        'xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">\n' + "\n".join(items) + "\n</oor:items>\n",
        encoding="utf-8")
    return len(pairs)


def convert_with_libreoffice(input_path: str, target_format: str, output_path: str, **kwargs) -> Dict[str, Any]:
    """
    使用LibreOffice进行格式转换，支持大部分Office格式互转
    Args:
        input_path: 输入文件路径
        target_format: 目标格式扩展名，比如pdf、docx、pptx等
        output_path: 输出文件路径
    """
    soffice = find_soffice()
    if not soffice:
        raise RuntimeError(
            "未找到 LibreOffice。安装方式：macOS 用 brew install --cask libreoffice；"
            "Windows 用 winget install TheDocumentFoundation.LibreOffice；"
            "Ubuntu/Debian 用 sudo apt install libreoffice。也可以用环境变量 ACKS_OFFICE_SOFFICE 指定路径。")

    with tempfile.TemporaryDirectory(prefix="acks-office-") as tmp:
        out_dir = Path(tmp) / "out"
        # 独立的用户配置目录：用户正开着 LibreOffice 时，共用配置会让无界面转换静默失败
        _seed_profile(Path(tmp) / "profile")
        profile = (Path(tmp) / "profile").as_uri()
        cmd = [soffice, f"-env:UserInstallation={profile}", "--headless", "--norestore",
               "--convert-to", target_format, "--outdir", str(out_dir), os.path.abspath(input_path)]
        try:
            result = subprocess.run(cmd, capture_output=True, text=True,
                                    timeout=kwargs.get("timeout", 300))
        except subprocess.TimeoutExpired:
            raise RuntimeError("LibreOffice转换超时") from None

        produced = list(out_dir.glob("*")) if out_dir.exists() else []
        if result.returncode != 0 or len(produced) != 1:
            detail = (result.stderr or result.stdout or "").strip() or "没有生成输出文件"
            raise RuntimeError(f"LibreOffice转换失败: {detail}")

        os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)
        shutil.move(str(produced[0]), output_path)

    return {"success": True, "output_path": output_path, "engine": "libreoffice"}

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
    return round(size_bytes / units.get(unit, 1), 2)

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
