
import os
import warnings
from typing import Optional, List, Dict, Any
from . import api, email

class OfficeSuite:
    """
    Office Suite 主类（已弃用，4.0 移除）：请改用 acks_office.create / extract / convert /
    add_watermark / merge，出错时抛出异常，而不是返回 {"success": False} 字典。
    """
    def __init__(self, config: Optional[Dict] = None, theme: Any = None):
        warnings.warn("OfficeSuite 已弃用，将在 4.0 移除；请改用 acks_office.create、extract、convert、"
                      "add_watermark、merge 等函数", DeprecationWarning, stacklevel=2)
        self.config = config or {}
        self.email_config = None
        self.theme = theme  # 主题名称或目录；None 为默认主题 neutral

    def create(self, doc_type: str, **kwargs) -> Dict[str, Any]:
        """
        创建Office文档
        Args:
            doc_type: 文档类型：word/excel/pdf/pptx
            **kwargs: 各类型文档的参数
        Returns:
            处理结果字典
        """
        if doc_type.lower() not in api.KINDS:
            return {"success": False, "error": f"不支持的文档类型: {doc_type}"}
        try:
            kwargs.setdefault("theme", self.theme)
            result = api.create(doc_type, **kwargs)
            return {"success": True, **result}
        except Exception as e:
            return {"success": False, "error": f"创建文档失败: {str(e)}"}
            
    def convert(self, input_path: str, to: str, output_path: Optional[str] = None, **kwargs) -> Dict[str, Any]:
        """格式转换（需要 LibreOffice）；默认输出到原文件旁边。"""
        if not os.path.exists(input_path):
            return {"success": False, "error": f"输入文件不存在: {input_path}"}
        try:
            return {"success": True, **api.convert(input_path, to, output_path, **kwargs)}
        except Exception as e:
            return {"success": False, "error": f"格式转换失败: {str(e)}"}

    def batch_convert(self, source_dir: str, target_dir: str, source_format: str, target_format: str, **kwargs) -> Dict[str, Dict]:
        """
        批量转换文档格式
        Args:
            source_dir: 源目录
            target_dir: 目标目录
            source_format: 源格式，比如docx
            target_format: 目标格式，比如pdf
        Returns:
            每个文件的处理结果字典
        """
        if not os.path.exists(source_dir):
            return {"error": f"源目录不存在: {source_dir}"}
        
        os.makedirs(target_dir, exist_ok=True)
        results = {}
        
        # 遍历源目录
        for filename in os.listdir(source_dir):
            if filename.lower().endswith(f".{source_format.lower()}"):
                input_path = os.path.join(source_dir, filename)
                output_name = f"{os.path.splitext(filename)[0]}.{target_format.lower()}"
                output_path = os.path.join(target_dir, output_name)
                
                try:
                    result = self.convert(input_path, target_format, output_path, **kwargs)
                    results[filename] = result
                except Exception as e:
                    results[filename] = {"success": False, "error": str(e)}
        
        return results
        
    def add_watermark(self, input_path: str, watermark_text: str, output_path: Optional[str] = None, **kwargs) -> Dict[str, Any]:
        """给 PDF 或 Word 加水印；默认写到新文件 <原名>_watermarked.<扩展名>，不覆盖原文件。"""
        if not os.path.exists(input_path):
            return {"success": False, "error": f"输入文件不存在: {input_path}"}
        try:
            return {"success": True, **api.add_watermark(input_path, watermark_text, output_path, **kwargs)}
        except Exception as e:
            return {"success": False, "error": f"添加水印失败: {str(e)}"}

    def batch_add_watermark(self, source_dir: str, target_dir: str, watermark_text: str, **kwargs) -> Dict[str, Dict]:
        """
        批量添加水印
        """
        if not os.path.exists(source_dir):
            return {"error": f"源目录不存在: {source_dir}"}
            
        os.makedirs(target_dir, exist_ok=True)
        results = {}
        
        for filename in os.listdir(source_dir):
            ext = os.path.splitext(filename)[1].lower()
            if ext in [".pdf", ".docx"]:
                input_path = os.path.join(source_dir, filename)
                output_path = os.path.join(target_dir, filename)
                
                try:
                    result = self.add_watermark(input_path, watermark_text, output_path, **kwargs)
                    results[filename] = result
                except Exception as e:
                    results[filename] = {"success": False, "error": str(e)}
        
        return results
        
    def config_email(self, **kwargs) -> None:
        """
        配置邮件参数
        """
        self.email_config = kwargs
        
    def send_email(self, to: List[str], subject: str, body: str, attachments: Optional[List[str]] = None, **kwargs) -> Dict[str, Any]:
        """
        发送邮件
        Args:
            to: 收件人列表
            subject: 邮件主题
            body: 邮件内容
            attachments: 附件路径列表
        Returns:
            发送结果
        """
        if not self.email_config:
            return {"success": False, "error": "请先调用config_email配置邮箱参数"}

        config = dict(self.email_config)
        if not config.get("password"):
            config["password"] = (os.environ.get("OFFICE_EMAIL_PASSWORD")
                                  or os.environ.get("ACKS_OFFICE_EMAIL_PASSWORD"))
        if not config["password"]:
            return {"success": False,
                    "error": "缺少邮箱密码/授权码：请在 config_email 传入 password，或设置环境变量 OFFICE_EMAIL_PASSWORD"}

        try:
            result = email.send_email(
                to=to,
                subject=subject,
                body=body,
                attachments=attachments,
                **config,
                **kwargs
            )
            return {"success": True, **result}
        except Exception as e:
            return {"success": False, "error": f"发送邮件失败: {str(e)}"}
            
    def execute_workflow(self, workflow_config: Dict) -> Dict[str, Any]:
        """
        执行自动化工作流（YAML 配置，见 examples/workflow_example.yaml）。

        支持步骤类型：data_extract / document_generate / convert /
        add_watermark / send_email / cleanup / notification / retry。
        支持 `{{ date }}`、`{{ yesterday }}` 及 data_extract 累积字段的占位符替换。
        """
        steps = workflow_config.get("steps", [])
        results = []
        context = self._build_workflow_context(workflow_config)

        for step in steps:
            name = step.get("name", "")
            stype = step.get("type", "")
            config = step.get("config", {}) or {}
            try:
                resolved = self._resolve_placeholders(config, context)
                if stype == "data_extract":
                    r = self._wf_data_extract(resolved)
                    if r.get("success") and isinstance(r.get("data"), dict):
                        context.update(r["data"])
                elif stype in ("document_generate", "generate"):
                    r = self._wf_document_generate(resolved)
                elif stype == "convert":
                    r = self._wf_convert(resolved)
                elif stype == "add_watermark":
                    r = self._wf_add_watermark(resolved)
                elif stype == "send_email":
                    r = self._wf_send_email(resolved)
                elif stype == "cleanup":
                    r = self._wf_cleanup(resolved)
                elif stype == "notification":
                    r = self._wf_notification(resolved)
                elif stype == "retry":
                    r = {"success": True, "skipped": True, "note": "retry 通常配置在 on_failure，不作为顺序步骤执行"}
                else:
                    r = {"success": False, "error": f"不支持的步骤类型: {stype}"}
                results.append({"step": name, "type": stype, **r})
            except Exception as e:
                results.append({"step": name, "type": stype, "success": False, "error": str(e)})

        all_ok = all(r.get("success") for r in results)
        return {"success": all_ok, "steps": results}

    def _build_workflow_context(self, workflow_config: Dict) -> Dict[str, Any]:
        import datetime
        today = datetime.date.today()
        yesterday = today - datetime.timedelta(days=1)
        return {
            "date": today.strftime("%Y-%m-%d"),
            "today": today.strftime("%Y-%m-%d"),
            "yesterday": yesterday.strftime("%Y-%m-%d"),
            "name": workflow_config.get("name", ""),
            "description": workflow_config.get("description", ""),
        }

    def _resolve_placeholders(self, value, context: Dict) -> Any:
        import re
        if isinstance(value, str):
            def _repl(m):
                key = m.group(1).strip()
                return str(context.get(key, m.group(0)))
            return re.sub(r"\{\{\s*([\w_]+)\s*\}\}", _repl, value)
        if isinstance(value, dict):
            return {k: self._resolve_placeholders(v, context) for k, v in value.items()}
        if isinstance(value, list):
            return [self._resolve_placeholders(v, context) for v in value]
        return value

    def _wf_data_extract(self, config: Dict) -> Dict[str, Any]:
        import json
        src = config.get("source", "file")
        if src == "file" or "path" in config:
            path = config.get("path") or config.get("file")
            with open(path, "r", encoding="utf-8") as f:
                data = json.load(f)
            return {"success": True, "data": data}
        if src == "api":
            endpoint = config.get("endpoint")
            params = config.get("params", {})
            try:
                import requests
                resp = requests.get(endpoint, params=params, timeout=30)
                resp.raise_for_status()
                data = resp.json()
            except ImportError:
                return {"success": False, "error": "data_extract 的 api 源需要安装 requests"}
            output = config.get("output")
            if output and data is not None:
                os.makedirs(os.path.dirname(output) or ".", exist_ok=True)
                with open(output, "w", encoding="utf-8") as f:
                    json.dump(data, f, ensure_ascii=False, indent=2)
            return {"success": True, "data": data}
        return {"success": False, "error": f"不支持的 data_extract source: {src}"}

    def _wf_document_generate(self, config: Dict) -> Dict[str, Any]:
        import json
        fmt = config.get("format", "docx")
        doc_type = {"docx": "word", "xlsx": "excel", "pptx": "pptx", "pdf": "pdf"}.get(fmt, fmt)
        data = dict(config.get("data") or {})
        if config.get("data_source"):
            with open(config["data_source"], "r", encoding="utf-8") as f:
                loaded = json.load(f)
            if isinstance(loaded, dict):
                data = {**loaded, **data}
        title = config.get("title") or data.get("title", "报告")
        content = config.get("content") or data.get("content", "")
        output = config.get("output")
        return self.create(doc_type, title=title, content=content, output_path=output)

    def _wf_convert(self, config: Dict) -> Dict[str, Any]:
        target = config.get("target_format") or config.get("to", "pdf")
        return self.convert(config["input"], to=target, output_path=config.get("output"))

    def _wf_add_watermark(self, config: Dict) -> Dict[str, Any]:
        return self.add_watermark(
            config["input"], config.get("watermark_text", "机密"),
            output_path=config.get("output"),
        )

    def _wf_send_email(self, config: Dict) -> Dict[str, Any]:
        to = config.get("to", [])
        if isinstance(to, str):
            to = [to]
        return self.send_email(
            to=to,
            subject=config.get("subject", ""),
            body=config.get("body", ""),
            attachments=config.get("attachments"),
        )

    def _wf_cleanup(self, config: Dict) -> Dict[str, Any]:
        import glob
        removed = 0
        for pattern in config.get("paths", []):
            for path in glob.glob(pattern):
                try:
                    os.remove(path)
                    removed += 1
                except OSError:
                    pass
        return {"success": True, "removed": removed}

    def _wf_notification(self, config: Dict) -> Dict[str, Any]:
        # 简化版：打印通知；后续可接入 wecom / slack 等
        print(f"[workflow-notification] {config.get('message', '')} → {config.get('channels', [])}")
        return {"success": True, "channels": config.get("channels", [])}
    
    def extract_data(self, input_path: str, **kwargs) -> Dict[str, Any]:
        """
        提取文档中的数据：Excel 返回行字典列表；Word、PDF 返回文本（表格按行输出）；
        PPT 返回 {幻灯片序号: 文本}
        """
        if not os.path.exists(input_path):
            return {"success": False, "error": f"输入文件不存在: {input_path}"}
        try:
            return {"success": True, "data": api.extract(input_path, **kwargs)}
        except ValueError as e:
            return {"success": False, "error": f"不支持该格式的数据提取: {e}"}
        except Exception as e:
            return {"success": False, "error": f"提取数据失败: {str(e)}"}
