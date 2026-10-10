"""邮件：密送不进邮件头、校验服务器证书、不明文发送密码，密码可从环境变量读取。不连接真实服务器。"""

import email
import smtplib
import ssl

import pytest

from acks_office import OfficeSuite
from acks_office.email import send_email


class FakeSMTP:
    sent = []
    starttls_supported = True

    def __init__(self, host, port, timeout=None, context=None):
        self.host, self.port, self.context = host, port, context

    def __enter__(self):
        return self

    def __exit__(self, *exc):
        return False

    def ehlo(self):
        pass

    def has_extn(self, name):
        return name == "starttls" and FakeSMTP.starttls_supported

    def starttls(self, context=None):
        self.context = context

    def login(self, user, password):
        self.password = password

    def sendmail(self, sender, recipients, message):
        FakeSMTP.sent.append({"sender": sender, "recipients": recipients, "message": message,
                              "password": self.password, "context": self.context})


@pytest.fixture(autouse=True)
def fake_smtp(monkeypatch):
    FakeSMTP.sent = []
    FakeSMTP.starttls_supported = True
    monkeypatch.setattr(smtplib, "SMTP", FakeSMTP)
    monkeypatch.setattr(smtplib, "SMTP_SSL", FakeSMTP)
    return FakeSMTP.sent


def _send(**kwargs):
    args = dict(to=["a@example.com"], subject="周报", body="正文", smtp_server="smtp.example.com",
                smtp_port=465, username="me@example.com", password="secret")
    args.update(kwargs)
    return send_email(**args)


def test_bcc_is_delivered_but_not_in_headers(fake_smtp):
    result = _send(cc=["c@example.com"], bcc=["hidden@example.com"])

    (sent,) = fake_smtp
    assert sent["recipients"] == ["a@example.com", "c@example.com", "hidden@example.com"]
    message = email.message_from_string(sent["message"])
    assert message["Bcc"] is None and "hidden@example.com" not in sent["message"]
    assert message["Cc"] == "c@example.com"
    assert message["Date"] and message["Message-ID"] == result["message_id"]


def test_missing_attachments_are_reported(tmp_path, fake_smtp):
    present = tmp_path / "附件.txt"
    present.write_text("内容", encoding="utf-8")
    result = _send(smtp_port=587, attachments=[str(present), str(tmp_path / "缺失.txt")])
    assert result["attachments"] == 1
    assert result["missing_attachments"] == [str(tmp_path / "缺失.txt")]


@pytest.mark.parametrize("variable", ["ACKS_OFFICE_EMAIL_PASSWORD", "OFFICE_EMAIL_PASSWORD"])
def test_password_can_come_from_environment(monkeypatch, fake_smtp, variable):
    monkeypatch.delenv("ACKS_OFFICE_EMAIL_PASSWORD", raising=False)
    monkeypatch.delenv("OFFICE_EMAIL_PASSWORD", raising=False)
    monkeypatch.setenv(variable, "from-env")
    with pytest.warns(DeprecationWarning):
        suite = OfficeSuite()
    suite.config_email(smtp_server="smtp.example.com", smtp_port=465, username="me@example.com")

    assert suite.send_email(["a@example.com"], "周报", "正文")["success"]
    assert fake_smtp[0]["password"] == "from-env"
    assert "password" not in suite.email_config  # 不把环境变量里的密码写回配置


def test_missing_password_is_an_error_not_a_crash(monkeypatch):
    monkeypatch.delenv("ACKS_OFFICE_EMAIL_PASSWORD", raising=False)
    monkeypatch.delenv("OFFICE_EMAIL_PASSWORD", raising=False)
    with pytest.warns(DeprecationWarning):
        suite = OfficeSuite()
    suite.config_email(smtp_server="smtp.example.com", smtp_port=465, username="me@example.com")
    result = suite.send_email(["a@example.com"], "周报", "正文")
    assert not result["success"] and "密码" in result["error"]


@pytest.mark.parametrize("port", [465, 587])
def test_server_certificate_is_verified(fake_smtp, port):
    _send(smtp_port=port)
    context = fake_smtp[0]["context"]
    assert context.verify_mode == ssl.CERT_REQUIRED and context.check_hostname


def test_refuses_to_send_password_in_plain_text(fake_smtp):
    FakeSMTP.starttls_supported = False
    with pytest.raises(RuntimeError, match="不支持加密连接"):
        _send(smtp_port=25)
    assert not fake_smtp

    _send(smtp_port=25, allow_insecure=True)
    assert fake_smtp and fake_smtp[0]["context"] is None
