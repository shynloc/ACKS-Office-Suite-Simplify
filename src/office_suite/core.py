
import os
from typing import Optional, List, Dict, Any
from . import docx, xlsx, pdf, pptx, email, utils

class OfficeSuite:
    """
    Office Suite 主类，统一接口处理所有Office文档操作
    """
    def __init__(self, config: Optional[Dict] = None):
        self.config = config or {}
        self.email_config = None
        
    def create(self, doc_type: str, **kwargs) -> Dict[str, Any]:
        """
        创建Office文档
        Args:
            doc_type: 文档类型：word/excel/pdf/pptx
            **kwargs: 各类型文档的参数
        Returns:
            处理结果字典
        """
        doc_type = doc_type.lower()
        handlers = {
            "word": docx.create_word,
            "docx": docx.create_word,
            "excel": xlsx.create_excel,
            "xlsx": xlsx.create_excel,
            "pdf": pdf.create_pdf,
            "pptx": pptx.create_pptx,
            "powerpoint": pptx.create_pptx
        }
        
        if doc_type not in handlers:
            return {"success": False, "error": f"不支持的文档类型: {doc_type}"}
        
        try:
            result = handlers[doc_type](**kwargs)
            return {"success": True, **result}
        except Exception as e:
            return {"success": False, "error": f"创建文档失败: {str(e)}"}
            
    def convert(self, input_path: str, to: str, output_path: Optional[str] = None, **kwargs) -> Dict[str, Any]:
        """
        格式转换
        Args:
            input_path: 输入文件路径
            to: 目标格式：docx/pdf/pptx/xlsx/html等
            output_path: 输出文件路径，可选，默认和输入同目录
        Returns:
            处理结果字典
        """
        if not os.path.exists(input_path):
            return {"success": False, "error": f"输入文件不存在: {input_path}"}
        
        try:
            # 自动生成输出路径
            if not output_path:
                input_dir = os.path.dirname(input_path)
                input_name = os.path.splitext(os.path.basename(input_path))[0]
                output_path = os.path.join(input_dir, f"{input_name}.{to.lower()}")
            
            # 调用对应转换方法
            input_ext = os.path.splitext(input_path)[1].lower().lstrip('.')
            target_ext = to.lower()
            
            # 基于libreoffice的通用转换（需要安装libreoffice）
            result = utils.convert_with_libreoffice(input_path, target_ext, output_path, **kwargs)
            return {"success": True, "output_path": output_path, **result}
            
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
        """
        给文档添加水印
        Args:
            input_path: 输入文件路径
            watermark_text: 水印文本
            output_path: 输出文件路径
        Returns:
            处理结果字典
        """
        if not os.path.exists(input_path):
            return {"success": False, "error": f"输入文件不存在: {input_path}"}
            
        try:
            ext = os.path.splitext(input_path)[1].lower()
            if ext == ".pdf":
                result = pdf.add_watermark(input_path, watermark_text, output_path, **kwargs)
            elif ext in [".docx", ".doc"]:
                result = docx.add_watermark(input_path, watermark_text, output_path, **kwargs)
            else:
                return {"success": False, "error": f"不支持给该格式添加水印: {ext}"}
                
            return {"success": True, **result}
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
            
        try:
            result = email.send_email(
                to=to,
                subject=subject,
                body=body,
                attachments=attachments,
                **self.email_config,
                **kwargs
            )
            return {"success": True, **result}
        except Exception as e:
            return {"success": False, "error": f"发送邮件失败: {str(e)}"}
            
    def execute_workflow(self, workflow_config: Dict) -> Dict[str, Any]:
        """
        执行自动化工作流
        """
        from . import workflow
        try:
            wf = workflow.WorkflowEngine(workflow_config)
            result = wf.execute()
            return {"success": True, **result}
        except Exception as e:
            return {"success": False, "error": f"执行工作流失败: {str(e)}"}
            
    def extract_data(self, input_path: str, **kwargs) -> Dict[str, Any]:
        """
        提取文档中的数据
        """
        if not os.path.exists(input_path):
            return {"success": False, "error": f"输入文件不存在: {input_path}"}
            
        try:
            ext = os.path.splitext(input_path)[1].lower()
            if ext in [".xlsx", ".xls"]:
                data = xlsx.extract_data(input_path, **kwargs)
            elif ext in [".docx", ".doc"]:
                data = docx.extract_text(input_path, **kwargs)
            elif ext == ".pdf":
                data = pdf.extract_text(input_path, **kwargs)
            else:
                return {"success": False, "error": f"不支持该格式的数据提取: {ext}"}
                
            return {"success": True, "data": data}
        except Exception as e:
            return {"success": False, "error": f"提取数据失败: {str(e)}"}
