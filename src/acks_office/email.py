
import os
import smtplib
import ssl
from typing import Optional, List, Dict, Any
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from email.header import Header
from email.utils import formatdate, make_msgid

def send_email(to: List[str], 
              subject: str, 
              body: str,
              smtp_server: str,
              smtp_port: int,
              username: str,
              password: str,
              attachments: Optional[List[str]] = None,
              cc: Optional[List[str]] = None,
              bcc: Optional[List[str]] = None,
              content_type: str = "plain",
              display_name: Optional[str] = None,
              allow_insecure: bool = False,
              **kwargs) -> Dict[str, Any]:
    """
    发送邮件
    Args:
        to: 收件人列表
        subject: 邮件主题
        body: 邮件内容
        smtp_server: SMTP服务器地址
        smtp_port: SMTP端口
        username: 发件人邮箱
        password: 邮箱密码/授权码
        attachments: 附件路径列表
        cc: 抄送人列表
        bcc: 密送人列表
        content_type: 内容类型：plain/html
        display_name: 发件人显示名称
        allow_insecure: 默认校验服务器证书、并要求加密连接；仅在内网自签名证书或服务器不支持加密时设为 True
    """
    # 创建邮件对象
    msg = MIMEMultipart()
    
    # 设置发件人
    if display_name:
        msg['From'] = f'{Header(display_name).encode()} <{username}>'
    else:
        msg['From'] = username
    
    # 设置收件人、抄送；密送只出现在投递名单里，不写进邮件头，否则所有收件人都能看到
    msg['To'] = ', '.join(to)
    if cc:
        msg['Cc'] = ', '.join(cc)

    # 设置主题
    msg['Subject'] = Header(subject, 'utf-8')
    msg['Date'] = formatdate(localtime=True)
    msg['Message-ID'] = make_msgid()
    
    # 添加正文
    msg.attach(MIMEText(body, content_type, 'utf-8'))
    
    # 添加附件
    attachment_count = 0
    missing_attachments = []
    if attachments:
        for file_path in attachments:
            if not os.path.exists(file_path):
                missing_attachments.append(file_path)
                continue
                
            filename = os.path.basename(file_path)
            with open(file_path, 'rb') as f:
                part = MIMEApplication(f.read())
            
            part.add_header(
                'Content-Disposition',
                'attachment',
                filename=Header(filename, 'utf-8').encode()
            )
            msg.attach(part)
            attachment_count += 1
    
    # 发送邮件
    all_recipients = to.copy()
    if cc:
        all_recipients.extend(cc)
    if bcc:
        all_recipients.extend(bcc)
    
    # smtplib 默认不校验证书，这里显式使用校验证书与主机名的上下文
    context = ssl._create_unverified_context() if allow_insecure else ssl.create_default_context()
    try:
        if smtp_port in (465, 994):
            # SSL连接
            with smtplib.SMTP_SSL(smtp_server, smtp_port, timeout=30, context=context) as server:
                server.login(username, password)
                server.sendmail(username, all_recipients, msg.as_string())
        else:
            # 普通连接，升级为 STARTTLS 后再登录
            with smtplib.SMTP(smtp_server, smtp_port, timeout=30) as server:
                server.ehlo()
                if server.has_extn('starttls'):
                    server.starttls(context=context)
                    server.ehlo()
                elif not allow_insecure:
                    raise RuntimeError("邮件服务器不支持加密连接（STARTTLS），为避免明文发送密码已停止；"
                                       "请改用 SSL 端口 465，或确认风险后传 allow_insecure=True")
                server.login(username, password)
                server.sendmail(username, all_recipients, msg.as_string())
        
        result = {
            "success": True,
            "recipients": len(all_recipients),
            "attachments": attachment_count,
            "message_id": msg.get("Message-ID")
        }
        if missing_attachments:
            result["missing_attachments"] = missing_attachments
        return result
        
    except Exception as e:
        raise RuntimeError(f"发送邮件失败: {str(e)}") from e
