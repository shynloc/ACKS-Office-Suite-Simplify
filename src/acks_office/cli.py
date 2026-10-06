"""acks-office 命令行。

所有命令都接受 --json，输出统一格式：{"ok", "data", "artifacts", "warnings", "error"}；
error 含 code、message、hint。命令默认不覆盖已有文件（加 --overwrite 才覆盖）。
退出码：0 成功，1 执行失败，2 参数错误。
"""

import argparse
import csv
import json
import os
import re
import sys
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

from . import __version__


class CliError(Exception):
    def __init__(self, code: str, message: str, hint: Optional[str] = None):
        super().__init__(message)
        self.code, self.message, self.hint = code, message, hint


def _envelope(ok: bool, data: Any = None, artifacts: Optional[List[Dict]] = None,
              warnings: Optional[List[Dict]] = None, error: Optional[Dict] = None) -> Dict:
    return {"ok": ok, "data": data, "artifacts": artifacts or [], "warnings": warnings or [], "error": error}


def _artifact(path: str) -> Dict:
    p = Path(path)
    return {"path": str(p.resolve()), "size": p.stat().st_size} if p.exists() else {"path": str(p)}


def _check_input(path: str) -> None:
    if not Path(path).is_file():
        raise CliError("FILE_NOT_FOUND", f"文件不存在：{path}")


def _check_output(path: str, overwrite: bool) -> None:
    if Path(path).exists() and not overwrite:
        raise CliError("OUTPUT_EXISTS", f"输出文件已存在：{path}", "换一个输出路径，或加 --overwrite")


def _read_text(value: Optional[str], file: Optional[str]) -> str:
    if file == "-":
        return sys.stdin.read()
    if file:
        _check_input(file)
        raw = Path(file).read_bytes()
        try:
            return raw.decode("utf-8-sig")  # Excel 导出的 CSV 常带 BOM
        except UnicodeDecodeError:
            return raw.decode("gb18030")  # Windows 中文环境下另存的文本
    return value or ""


_INT = re.compile(r"^-?(0|[1-9]\d{0,14})$")
_FLOAT = re.compile(r"^-?(0|[1-9]\d{0,14})\.\d{1,10}$")


def _csv_value(text: str) -> Any:
    """CSV 里的数字转成数值；以 0 开头的编号、超长数字（如身份证号）保持文本。"""
    if _INT.match(text):
        return int(text)
    if _FLOAT.match(text):
        return float(text)
    return text


def _load_json(file: str) -> Any:
    try:
        return json.loads(_read_text(None, file))
    except json.JSONDecodeError as exc:
        raise CliError("INVALID_INPUT", f"{file} 不是有效的 JSON：{exc}") from None


def _load_table(file: str) -> Tuple[Optional[List[Any]], Optional[List[Dict[str, Any]]], Dict[str, Any]]:
    """Excel 数据 → (data, sheets, meta)。

    JSON 二维数组（第一行为表头）、对象数组（每行一个对象，键为表头，和 extract 的输出一样）、
    {"meta": {...}, "sheets": [...]}（多张工作表），或 CSV。
    """
    if file != "-" and file.lower().endswith(".csv"):
        rows = list(csv.reader(_read_text(None, file).splitlines()))
        return rows[:1] + [[_csv_value(v) for v in row] for row in rows[1:]], None, {}
    data = _load_json(file)
    if isinstance(data, dict):
        meta, sheets = data.get("meta") or {}, data.get("sheets")
        if not (isinstance(sheets, list) and sheets and isinstance(meta, dict)
                and all(isinstance(s, dict) and isinstance(s.get("rows", []), list) for s in sheets)):
            raise CliError("INVALID_INPUT", "多张工作表需要写成 {\"sheets\": [{\"name\": \"Q3\", "
                                            "\"header\": [\"部门\", \"7月\"], \"rows\": [[\"一部\", 150]]}]}")
        return None, sheets, meta
    if isinstance(data, list) and (all(isinstance(r, list) for r in data) or
                                   (data and all(isinstance(r, dict) for r in data))):
        return data, None, {}
    raise CliError("INVALID_INPUT", "Excel 数据需要是二维数组，例如 [[\"部门\", \"1月\"], [\"一部\", 150]]；"
                                    "也可以是对象数组，或 {\"sheets\": [...]}（多张工作表）")


def _load_slides(file: str) -> Tuple[List[Dict[str, Any]], Dict[str, Any]]:
    """幻灯片 JSON：数组，或 {"meta": {...}, "slides": [...]}（meta 为整套的元数据）。"""
    data = _load_json(file)
    meta: Dict[str, Any] = {}
    if isinstance(data, dict):
        meta = data.get("meta") or {}
        data = data.get("slides")
    if not (isinstance(data, list) and data and all(isinstance(s, dict) for s in data)
            and isinstance(meta, dict)):
        raise CliError("INVALID_INPUT", "幻灯片需要是非空 JSON 数组，或 {\"meta\": {...}, \"slides\": [...]}，例如 "
                                        "[{\"title\": \"封面\", \"layout\": \"title\"}, {\"title\": \"要点\", \"content\": \"…\"}]")
    return data, meta


def _suite(theme: str):
    try:
        from .core import OfficeSuite
    except ImportError as exc:
        raise CliError("DEPENDENCY_MISSING", f"缺少依赖：{exc.name or exc}",
                       "运行 doctor 命令查看缺少哪些依赖及安装步骤") from None
    return OfficeSuite(theme=theme)


def _from_result(result: Dict, code: str, output: Optional[str] = None) -> Dict:
    if not result.get("success"):
        message = result.get("error", "未知错误")
        if "LibreOffice" in message and "未找到" in message:
            raise CliError("ENGINE_UNAVAILABLE", message, "运行 acks-office doctor 查看安装方式")
        raise CliError(code, message)
    data = {k: v for k, v in result.items() if k not in ("success", "warnings")}
    return _envelope(True, data, [_artifact(output)] if output else [], result.get("warnings"))


# ---------------------------------------------------------------- 命令

def cmd_version(args) -> Dict:
    return _envelope(True, {"version": __version__})


def cmd_doctor(args) -> Dict:
    from . import doctor
    return _envelope(True, doctor.run(skill_dir=args.skill_dir, network=not args.no_network,
                                      probe=not args.no_probe))


_CREATE_SUFFIX = {"word": ".docx", "pdf": ".pdf", "excel": ".xlsx", "pptx": ".pptx"}
_EXTRACT_KEY = {".xlsx": "rows", ".xls": "rows", ".pptx": "slides", ".docx": "text", ".pdf": "text"}


def cmd_create(args) -> Dict:
    kind = {"docx": "word", "xlsx": "excel", "ppt": "pptx"}.get(args.type, args.type)
    suffix = _CREATE_SUFFIX[kind]
    if Path(args.output).suffix.lower() != suffix:
        raise CliError("INVALID_INPUT", f"输出文件应以 {suffix} 结尾：{args.output}",
                       f"例如 -o {Path(args.output).stem or '输出'}{suffix}")
    _check_output(args.output, args.overwrite)
    if args.theme not in ("acks", "default"):
        from .themes import load_theme
        load_theme(args.theme)  # 主题不存在时直接报 THEME_NOT_FOUND
    kwargs: Dict[str, Any] = {"output_path": args.output, "font_policy": args.font_policy}
    for key in ("brand_name", "footer_label", "subtitle"):
        if getattr(args, key) is not None:
            kwargs[key] = getattr(args, key)
    for item in args.meta or []:
        key, sep, value = item.partition("=")
        if not sep or not key.strip():
            raise CliError("INVALID_INPUT", f"--meta 应写成 键=值：{item}", "例如 --meta author=经营分析部")
        kwargs[key.strip()] = value.replace("\\n", "\n")

    title = args.title
    if kind in ("word", "pdf"):
        kwargs["content"] = _read_text(args.content, args.content_file)
        if title is None:
            from .render.common import split_front_matter
            title = split_front_matter(kwargs["content"])[0].get("title")
        if args.content_file and args.content_file != "-":
            kwargs["base_dir"] = str(Path(args.content_file).resolve().parent)
        if kind == "pdf" and args.font:
            kwargs["font"] = args.font
    elif kind == "excel":
        if not args.data_file:
            raise CliError("INVALID_INPUT", "生成 Excel 需要 --data-file（JSON 二维数组、对象数组、{\"sheets\": [...]} 或 CSV）")
        data, sheets, table_meta = _load_table(args.data_file)
        if sheets is not None:
            if args.theme in ("acks", "default"):
                raise CliError("INVALID_INPUT", "多张工作表需要主题，2.x 的内置样式只支持二维数组",
                               "加上 --theme neutral（或 slate、folio）")
            kwargs["sheets"] = sheets
        else:
            kwargs["data"] = data
        kwargs["create_chart"] = args.chart
        for key, value in table_meta.items():
            if key == "title" and title is None:
                title = value
            elif key != "title":
                kwargs.setdefault(key, value)
    elif kind == "pptx":
        if not args.slides_file:
            raise CliError("INVALID_INPUT", "生成 PPT 需要 --slides-file（JSON 数组，每项含 title、content、layout）")
        kwargs["slides"], deck_meta = _load_slides(args.slides_file)
        for key, value in deck_meta.items():
            if key == "title" and title is None:
                title = value
            elif key != "title":
                kwargs.setdefault(key, value)

    if kind == "excel" and args.theme not in ("acks", "default"):
        # 表格的标题写在第 1 行：没给标题就不加标题行，工作表名用文件名
        kwargs["title"] = title or kwargs.get("title")
        kwargs.setdefault("sheet_name", Path(args.output).stem)
    else:
        kwargs["title"] = title or kwargs.get("title") or Path(args.output).stem
    os.makedirs(os.path.dirname(os.path.abspath(args.output)), exist_ok=True)
    return _from_result(_suite(args.theme).create(kind, **kwargs), "CREATE_FAILED", args.output)


def cmd_extract(args) -> Dict:
    _check_input(args.file)
    suffix = Path(args.file).suffix.lower()
    if suffix not in _EXTRACT_KEY:
        hint = "旧版 Office 格式可先用 convert --to docx / pptx 转换" if suffix in (".doc", ".ppt") else None
        raise CliError("UNSUPPORTED_FORMAT", f"不支持读取 {suffix or '无扩展名'} 文件，"
                                             f"支持：{', '.join(_EXTRACT_KEY)}", hint)
    kwargs = {}
    if args.sheet is not None:
        kwargs["sheet_name"] = int(args.sheet) if args.sheet.isdigit() else args.sheet
    result = _suite("acks").extract_data(args.file, **kwargs)
    if not result.get("success"):
        raise CliError("EXTRACT_FAILED", result.get("error", "未知错误"))
    return _envelope(True, {"format": suffix.lstrip("."), _EXTRACT_KEY[suffix]: result["data"]})


def cmd_convert(args) -> Dict:
    _check_input(args.file)
    output = args.output or str(Path(args.file).with_suffix("." + args.to.split(":")[0]))
    _check_output(output, args.overwrite)
    return _from_result(_suite("acks").convert(args.file, to=args.to, output_path=output),
                        "CONVERSION_FAILED", output)


def cmd_watermark(args) -> Dict:
    _check_input(args.file)
    source = Path(args.file)
    output = args.output or str(source.with_name(f"{source.stem}_watermarked{source.suffix}"))
    _check_output(output, args.overwrite)
    return _from_result(_suite("acks").add_watermark(args.file, args.text, output_path=output),
                        "WATERMARK_FAILED", output)


def cmd_merge(args) -> Dict:
    _check_output(args.output, args.overwrite)
    if not any(Path(f).is_file() for f in args.files):
        raise CliError("FILE_NOT_FOUND", f"要合并的文件都不存在：{', '.join(args.files)}")
    suffixes = {Path(f).suffix.lower() for f in args.files}
    if suffixes == {".pdf"}:
        from .pdf import merge_pdfs as merge
    elif suffixes == {".docx"}:
        from .docx import merge_documents as merge
    else:
        raise CliError("UNSUPPORTED_FORMAT", "只能合并同一种格式：全部 .pdf 或全部 .docx")
    result = merge(args.files, args.output)
    warnings = list(result.pop("warnings", []))
    if result.get("skipped"):
        warnings.append({"code": "INPUT_SKIPPED", "message": f"以下文件不存在，已跳过：{', '.join(result['skipped'])}"})
    return _envelope(True, result, [_artifact(args.output)], warnings)


def cmd_fonts(args) -> Dict:
    from . import fonts
    if args.action == "install":
        if args.theme:
            from .themes import load_theme
            load_theme(args.theme)  # 主题不存在时报 THEME_NOT_FOUND
            keys = fonts.theme_font_status(args.theme)["install"]
            if args.name:
                keys = [k for k in keys if k == args.name] or [args.name]
        elif args.name:
            keys = [args.name]
        else:
            raise CliError("INVALID_INPUT", f"请指定字体（{', '.join(fonts.CATALOG)}）或 --theme 主题",
                           "例如 fonts install noto-sans-sc，或 fonts install --theme folio --system")
        installed, artifacts = [], []
        for key in keys:
            try:
                data = fonts.install_font(key, progress=lambda msg: print(msg, file=sys.stderr), system=args.system)
            except (ValueError, RuntimeError, OSError) as exc:
                raise CliError("FONT_INSTALL_FAILED", f"{key}：{exc}") from None
            installed.append(data)
            artifacts += [_artifact(f) for f in data["files"]]
        if len(installed) == 1 and not args.theme:
            return _envelope(True, installed[0], artifacts)
        return _envelope(True, {"theme": args.theme, "installed": installed}, artifacts)
    return _envelope(True, fonts.fonts_report())


def cmd_theme(args) -> Dict:
    from . import themes
    if args.action == "list":
        return _envelope(True, {"default": themes.DEFAULT_THEME, "user_dir": str(themes.user_theme_dir()),
                                "themes": themes.list_themes()})
    if not args.name:
        raise CliError("USAGE_ERROR", "请指定主题名称或主题目录", f"例如 theme {args.action} slate")
    if args.action == "show":
        return _envelope(True, themes.load_theme(args.name).describe())
    report = themes.validate_theme(args.name, check_installed=not args.no_fonts)
    if report["ok"]:
        return _envelope(True, report, warnings=report["warnings"])
    first = report["errors"][0]
    return _envelope(False, report, warnings=report["warnings"],
                     error={"code": "THEME_INVALID", "message": f"主题 {report['theme']} 有 {len(report['errors'])} 处错误",
                            "hint": f"{first['path']}：{first['message']}"})


# ---------------------------------------------------------------- 入口

class _Parser(argparse.ArgumentParser):
    """参数错误也走统一的错误格式，--json 时 Agent 能直接解析。"""

    def error(self, message):
        raise CliError("USAGE_ERROR", message, f"运行 {self.prog} --help 查看用法")


def _parser() -> argparse.ArgumentParser:
    # --json 放在子命令前后都可以；子命令里用 SUPPRESS，避免覆盖前面已设的值
    common = argparse.ArgumentParser(add_help=False)
    common.add_argument("--json", action="store_true", default=argparse.SUPPRESS,
                        help="以 JSON 输出（Agent 调用时使用）")

    parser = _Parser(prog="acks-office", description="按文档设计规范生成、读取、转换 Office 文档")
    parser.add_argument("--json", action="store_true", help="以 JSON 输出（Agent 调用时使用）")
    parser.add_argument("--version", action="version", version=f"acks-office {__version__}")
    sub = parser.add_subparsers(dest="command", required=True)

    sub.add_parser("version", parents=[common], help="显示版本").set_defaults(func=cmd_version)

    p = sub.add_parser("doctor", parents=[common], help="只读检查环境，给出能力等级与补齐计划")
    p.add_argument("--skill-dir", help="技能所在目录（用于判断宿主 Agent）")
    p.add_argument("--no-network", action="store_true", help="不检查网络")
    p.add_argument("--no-probe", action="store_true", help="不运行 LibreOffice --version")
    p.set_defaults(func=cmd_doctor)

    p = sub.add_parser("create", parents=[common], help="生成文档")
    p.add_argument("type", choices=["word", "docx", "excel", "xlsx", "pdf", "pptx", "ppt"])
    p.add_argument("-o", "--output", required=True, help="输出文件路径")
    p.add_argument("--title", help="标题（默认用输出文件名）")
    p.add_argument("--content", help="正文 Markdown（Word / PDF）")
    p.add_argument("--content-file", help="正文 Markdown 文件，- 表示从标准输入读取")
    p.add_argument("--data-file", help="Excel 数据：JSON 二维数组或 CSV")
    p.add_argument("--slides-file", help="PPT 幻灯片：JSON 数组")
    p.add_argument("--theme", default="acks",
                   help="主题名称（neutral、slate、folio 或已安装的主题）或主题目录；acks、default 是 2.x 的内置样式")
    p.add_argument("--meta", action="append", metavar="KEY=VALUE",
                   help="文档元数据，可重复：kicker、author、date、version、issue、lede 等（也可写在正文 front matter）")
    p.add_argument("--font-policy", default="local", choices=["local", "theme"],
                   help="local：缺主题字体时改用本机字体；theme：总是写主题字体名")
    p.add_argument("--brand-name", help="品牌名，传空字符串去掉品牌")
    p.add_argument("--footer-label", help="页脚文字")
    p.add_argument("--subtitle", help="封面副标题（Word）")
    p.add_argument("--font", help="PDF 中文字体文件路径（TrueType）")
    p.add_argument("--chart", action="store_true", help="Excel 附柱状图（default 主题）")
    p.add_argument("--overwrite", action="store_true", help="允许覆盖已有文件")
    p.set_defaults(func=cmd_create)

    p = sub.add_parser("extract", parents=[common], help="提取文本或表格数据")
    p.add_argument("file")
    p.add_argument("--sheet", help="Excel 工作表名或从 0 开始的索引")
    p.set_defaults(func=cmd_extract)

    p = sub.add_parser("convert", parents=[common], help="格式转换（需要 LibreOffice）")
    p.add_argument("file")
    p.add_argument("--to", required=True, help="目标格式，如 pdf、docx")
    p.add_argument("-o", "--output", help="输出路径（默认与原文件同名、换扩展名）")
    p.add_argument("--overwrite", action="store_true")
    p.set_defaults(func=cmd_convert)

    p = sub.add_parser("watermark", parents=[common], help="添加水印（PDF / Word），默认写到新文件")
    p.add_argument("file")
    p.add_argument("--text", required=True)
    p.add_argument("-o", "--output", help="输出路径（默认 原文件名_watermarked）")
    p.add_argument("--overwrite", action="store_true")
    p.set_defaults(func=cmd_watermark)

    p = sub.add_parser("merge", parents=[common], help="合并多个 PDF 或 Word")
    p.add_argument("files", nargs="+")
    p.add_argument("-o", "--output", required=True)
    p.add_argument("--overwrite", action="store_true")
    p.set_defaults(func=cmd_merge)

    p = sub.add_parser("theme", parents=[common], help="查看、校验主题")
    p.add_argument("action", choices=["list", "show", "validate"])
    p.add_argument("name", nargs="?", help="主题名称（如 slate）或主题目录")
    p.add_argument("--no-fonts", action="store_true", help="校验时不检查本机字体")
    p.set_defaults(func=cmd_theme)

    p = sub.add_parser("fonts", parents=[common], help="查看字体情况或安装开源字体")
    p.add_argument("action", choices=["list", "install"])
    p.add_argument("name", nargs="?", help="要安装的字体，如 noto-sans-sc")
    p.add_argument("--theme", help="安装某个主题缺少的全部字体，如 --theme folio")
    p.add_argument("--system", action="store_true",
                   help="同时装到当前用户的字体目录，Word、PowerPoint 等软件也能用（不需要管理员权限）")
    p.set_defaults(func=cmd_fonts)
    return parser


def _print_theme_report(report: Dict) -> None:
    status = "通过" if report["ok"] else f"未通过：{len(report['errors'])} 处错误"
    print(f"主题 {report['theme']}（{report.get('title', '')}）校验{status}，{len(report['warnings'])} 条提醒")
    for item in report["errors"]:
        print(f"  错误 {item['path']}：{item['message']}")
    for item in report["warnings"]:
        print(f"  提醒 {item['path']}：{item['message']}")


def _print_human(command: str, env: Dict) -> None:
    if command == "theme" and isinstance(env["data"], dict) and "contrast" in env["data"]:
        _print_theme_report(env["data"])
        return
    if not env["ok"]:
        err = env["error"]
        print(f"错误：{err['message']}")
        if err.get("hint"):
            print(f"提示：{err['hint']}")
        for warning in env["warnings"]:
            print(f"注意：{warning.get('message', warning)}")
        return
    if command == "theme" and isinstance(env["data"], dict) and "themes" in env["data"]:
        for t in env["data"]["themes"]:
            mark = "（默认）" if t["name"] == env["data"]["default"] else ""
            print(f"{t['name']:12} {t.get('title', '')}{mark} · {t['source']} · {t.get('description', '')}")
        return
    if command == "doctor":
        r = env["data"]
        print(f"acks-office {r['version']} · 能力等级 {r['level']} · 宿主 {r['host']['guess'] or '未识别'}")
        missing = [d["name"] for d in r["dependencies"] if not d["ok"]]
        print(f"依赖：{'齐全' if not missing else '缺少 ' + ', '.join(missing)}")
        pdf = r["fonts"]["pdf_cjk"]
        print(f"PDF 中文字体：{pdf.get('family', '无')}（{'嵌入' if pdf.get('embedded') else '不嵌入'}）")
        print(f"主题字体缺少：{', '.join(r['fonts']['theme']['missing']) or '无'}")
        lo = r["office_apps"][0]
        print(f"LibreOffice：{lo['version'] or lo['path'] or '未安装'}")
        for item in r["plan"]:
            print(f"- 建议：{item['why']}")
        return
    for artifact in env["artifacts"]:
        print(f"已生成：{artifact['path']}")
    for warning in env["warnings"]:
        print(f"注意：{warning.get('message', warning)}")
    if not env["artifacts"]:
        print(json.dumps(env["data"], ensure_ascii=False, indent=2))


def _error(exc: CliError) -> Dict:
    return _envelope(False, error={"code": exc.code, "message": exc.message, "hint": exc.hint})


def main(argv: Optional[List[str]] = None) -> int:
    argv = sys.argv[1:] if argv is None else argv
    try:
        args = _parser().parse_args(argv)
    except CliError as exc:
        args = argparse.Namespace(json="--json" in argv, command=None)
        env, status = _error(exc), 2
    else:
        try:
            env = args.func(args)
        except CliError as exc:
            env = _error(exc)
        except Exception as exc:  # 主题错误带自己的错误码；其他未预料的错误也按统一格式返回
            from .themes import ThemeError
            if isinstance(exc, ThemeError):
                env = _error(CliError(exc.code, exc.message, exc.hint))
            else:
                env = _error(CliError("INTERNAL_ERROR", f"{type(exc).__name__}: {exc}",
                                      "可先运行 acks-office doctor 检查环境"))
        status = 0 if env["ok"] else 1
    if args.json:
        # Windows 管道里控制台编码不一定是 UTF-8，此时输出转义后的 JSON，保证可解析
        ascii_only = os.name == "nt" and not sys.stdout.isatty()
        print(json.dumps(env, ensure_ascii=ascii_only, indent=2, default=str))
    else:
        _print_human(args.command, env)
    return status
