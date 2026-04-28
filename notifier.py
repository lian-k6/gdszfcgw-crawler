"""
通知模块：支持邮件发送（SMTP）和微信通知（Server酱/企业微信机器人）
"""

import logging
import smtplib
import os
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from datetime import datetime

try:
    import requests
except ImportError:
    requests = None

try:
    from config import EMAIL_CONFIG, WECHAT_CONFIG
except ImportError:
    EMAIL_CONFIG = {}
    WECHAT_CONFIG = {}

# 支持从环境变量读取配置（用于GitHub Actions等CI环境）
def _env_override():
    """用环境变量覆盖配置文件中的敏感信息"""
    global EMAIL_CONFIG, WECHAT_CONFIG

    if not EMAIL_CONFIG:
        EMAIL_CONFIG = {}
    if not WECHAT_CONFIG:
        WECHAT_CONFIG = {}

    # 邮件配置
    if os.environ.get("SMTP_SERVER"):
        EMAIL_CONFIG["smtp_server"] = os.environ["SMTP_SERVER"]
    if os.environ.get("SMTP_PORT"):
        EMAIL_CONFIG["smtp_port"] = int(os.environ["SMTP_PORT"])
    if os.environ.get("SENDER_EMAIL"):
        EMAIL_CONFIG["sender_email"] = os.environ["SENDER_EMAIL"]
    if os.environ.get("SENDER_PASSWORD"):
        EMAIL_CONFIG["sender_password"] = os.environ["SENDER_PASSWORD"]
    if os.environ.get("RECEIVER_EMAILS"):
        EMAIL_CONFIG["receiver_emails"] = [
            e.strip() for e in os.environ["RECEIVER_EMAILS"].split(",") if e.strip()
        ]

    # 微信配置
    if os.environ.get("SERVER_CHAN_KEY"):
        WECHAT_CONFIG["server_chan_key"] = os.environ["SERVER_CHAN_KEY"]
    if os.environ.get("QYWX_WEBHOOK"):
        WECHAT_CONFIG["qywx_webhook"] = os.environ["QYWX_WEBHOOK"]

_env_override()


def send_email(subject, body, attachment_path=None):
    """
    发送邮件通知
    :param subject: 邮件主题
    :param body: 邮件正文
    :param attachment_path: 附件路径（CSV文件）
    :return: 是否发送成功
    """
    cfg = EMAIL_CONFIG
    smtp_server = cfg.get("smtp_server", "")
    smtp_port = cfg.get("smtp_port", 465)
    sender = cfg.get("sender_email", "")
    password = cfg.get("sender_password", "")
    receivers = cfg.get("receiver_emails", [])

    if not all([smtp_server, sender, password, receivers]):
        logging.info("邮件配置不完整，跳过邮件发送")
        return False

    try:
        msg = MIMEMultipart()
        msg["From"] = sender
        msg["To"] = ", ".join(receivers)
        msg["Subject"] = subject

        # 添加正文
        msg.attach(MIMEText(body, "plain", "utf-8"))

        # 添加附件
        if attachment_path and os.path.exists(attachment_path):
            with open(attachment_path, "rb") as f:
                attachment = MIMEApplication(f.read())
                attachment.add_header(
                    "Content-Disposition",
                    "attachment",
                    filename=os.path.basename(attachment_path)
                )
                msg.attach(attachment)

        # 发送邮件
        server = smtplib.SMTP_SSL(smtp_server, smtp_port, timeout=30)
        server.login(sender, password)
        server.sendmail(sender, receivers, msg.as_string())
        server.quit()

        logging.info(f"邮件发送成功: {subject}")
        return True

    except Exception as e:
        logging.error(f"邮件发送失败: {e}")
        return False


def send_wechat(title, content):
    """
    发送微信通知
    优先使用企业微信机器人，其次使用Server酱
    :param title: 消息标题
    :param content: 消息内容
    :return: 是否发送成功
    """
    if requests is None:
        logging.warning("未安装requests库，无法发送微信通知")
        return False

    cfg = WECHAT_CONFIG

    # 优先尝试企业微信机器人
    qywx_webhook = cfg.get("qywx_webhook", "")
    if qywx_webhook:
        try:
            data = {
                "msgtype": "text",
                "text": {
                    "content": f"{title}\n{content}"
                }
            }
            resp = requests.post(qywx_webhook, json=data, timeout=10)
            if resp.json().get("errcode") == 0:
                logging.info("企业微信通知发送成功")
                return True
            else:
                logging.warning(f"企业微信通知失败: {resp.text}")
        except Exception as e:
            logging.error(f"企业微信通知异常: {e}")

    # 尝试Server酱
    server_key = cfg.get("server_chan_key", "")
    if server_key:
        try:
            url = f"https://sctapi.ftqq.com/{server_key}.send"
            data = {
                "title": title,
                "desp": content
            }
            resp = requests.post(url, data=data, timeout=10)
            if resp.json().get("code") == 0:
                logging.info("Server酱通知发送成功")
                return True
            else:
                logging.warning(f"Server酱通知失败: {resp.text}")
        except Exception as e:
            logging.error(f"Server酱通知异常: {e}")

    logging.info("微信通知配置不完整，跳过微信发送")
    return False


def notify_all(title, content, attachment_path=None):
    """
    同时发送邮件和微信通知
    :param title: 通知标题
    :param content: 通知内容
    :param attachment_path: 邮件附件路径
    """
    send_wechat(title, content)

    email_subject = f"{title} - {datetime.now().strftime('%Y-%m-%d')}"
    send_email(email_subject, content, attachment_path)
