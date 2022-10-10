# 发送邮件
# 不支持邮件正文和内嵌资源排版
# 更新时间： 2022-10-10 16:37:38
import filetype
from pathlib import Path
from smtplib import SMTP
from email.message import EmailMessage


class Mail:
    def __init__(self, host: str, port: int, user: str, password: str):
        self.smtp = SMTP(host=host, port=port)
        self.smtp.login(user=user, password=password)

    def send(
        self,
        _from: str,
        _to=[],
        _cc=[],
        _bcc=[],
        **kwargs,
    ):
        msg = EmailMessage()
        assert _to or _cc or _bcc, '收件人、抄送人、密抄送人不能同时为空'
        for k, v in kwargs.items():
            # 邮件主题
            if k == 'subject':
                msg.add_header('Subject', v)
            # 邮件正文
            elif k == 'content':
                msg.add_alternative(v)
            # 邮件内嵌资源
            elif k == 'resources':
                for resource in v:
                    resource_path = Path(resource)
                    if resource_path.exists():
                        mime_type = filetype.guess_mime(resource_path.absolute())
                        if not mime_type:
                            mime_type = 'application/octet-stream'
                        maintype, subtype = mime_type.split('/', 1)
                        with open(resource_path, 'rb') as f:
                            if maintype in ['image', 'video', 'audio']:
                                msg.add_related(
                                    f.read(),
                                    maintype=maintype,
                                    subtype=subtype,
                                    filename=resource_path.name,
                                )
                            else:
                                msg.add_attachment(
                                    f.read(),
                                    maintype=maintype,
                                    subtype=subtype,
                                    filename=resource_path.name,
                                )
            # 邮件附件
            elif k == 'attachments':
                for attachment in v:
                    attachment_path = Path(attachment)
                    if attachment_path.exists():
                        mime_type = filetype.guess_mime(attachment_path.absolute())
                        if not mime_type:
                            mime_type = 'application/octet-stream'
                        maintype, subtype = mime_type.split('/', 1)
                        with open(attachment_path, 'rb') as f:
                            msg.add_attachment(
                                f.read(),
                                maintype=maintype,
                                subtype=subtype,
                                filename=attachment_path.name,
                            )
            else:
                assert False, f'无效参数[{k}]'
        msg.add_header('From', _from)
        if _to:
            msg.add_header('To', ','.join(_to))
        if _cc:
            msg.add_header('Cc', ','.join(_cc))
        if _bcc:
            msg.add_header('Bcc', ','.join(_bcc))
        self.smtp.sendmail(
            from_addr=_from, to_addrs=_to + _cc + _bcc, msg=msg.as_string()
        )
