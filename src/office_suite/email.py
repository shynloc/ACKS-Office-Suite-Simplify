
import os
import smtplib
from typing import Optional, List, Dict, Any
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from email.header import Header

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
    """
    # 创建邮件对象
    msg = MIMEMultipart()
    
    # 设置发件人
    if display_name:
        msg['From'] = f'{Header(display_name).encode()} <{username}>'
    else:
        msg['From'] = username
    
    # 设置收件人、抄送、密送
    msg['To'] = ', '.join(to)
    if cc:
        msg['Cc'] = ', '.join(cc)
    if bcc:
        msg['Bcc'] = ', '.join(bcc)
    
    # 设置主题
    msg['Subject'] = Header(subject, 'utf-8')
    
    # 添加正文
    msg.attach(MIMEText(body, content_type, 'utf-8'))
    
    # 添加附件
    attachment_count = 0
    if attachments:
        for file_path in attachments:
            if not os.path.exists(file_path):
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
    
    try:
        if smtp_port in (465, 994):
            # SSL连接
            with smtplib.SMTP_SSL(smtp_server, smtp_port, timeout=30) as server:
                server.login(username, password)
                server.sendmail(username, all_recipients, msg.as_string())
        else:
            # 普通连接，支持STARTTLS
            with smtplib.SMTP(smtp_server, smtp_port, timeout=30) as server:
                server.ehlo()
                # 尝试启用TLS
                if server.has_extn('starttls'):
                    server.starttls()
                    server.ehlo()
                server.login(username, password)
                server.sendmail(username, all_recipients, msg.as_string())
        
        return {
            "success": True,
            "recipients": len(all_recipients),
            "attachments": attachment_count,
            "message_id": msg.get("Message-ID")
        }
        
    except Exception as e:
        raise RuntimeError(f"发送邮件失败: {str(e)}")
