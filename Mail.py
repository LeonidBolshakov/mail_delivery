import os
from smtplib import SMTP, SMTPAuthenticationError
import mimetypes
from email.mime.multipart import MIMEMultipart
from email import encoders
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.mime.image import MIMEImage
from email.mime.audio import MIMEAudio
from loguru import logger
import constant as c


class CreateMails:
    @staticmethod
    def create_message_body(self, msg, content_: str, list_attachments: [str]):
        msg.attach(MIMEText(content_, 'plain'))

        for path in list_attachments:
            self.attach_file(msg, path)

    def create(self, from_, list_to, subject, content_, list_attachments):
        msgs = []
        for to in list_to:
            msg = MIMEMultipart()
            msg['Subject'] = subject
            msg['To'] = to
            msg['From'] = from_

            self.create_message_body(self, msg, content_, list_attachments)
            msgs.append(msg)

        return msgs

    @staticmethod
    def attach_file(msg, filepath):
        filename = os.path.basename(filepath)
        ctype, encoding = mimetypes.guess_type(filepath)      # Определяем тип файла на основе его расширения
        if ctype is None or encoding is not None:
            ctype = 'application/octet-stream'                # Будем использовать общий тип
        maintype, subtype = ctype.split('/', 1)  # Получаем тип и подтип
        if maintype == 'text':
            with open(filepath) as fp:
                file = MIMEText(fp.read(), _subtype=subtype)
                fp.close()
        elif maintype == 'image':
            with open(filepath, 'rb') as fp:
                file = MIMEImage(fp.read(), _subtype=subtype)
                fp.close()
        elif maintype == 'audio':
            with open(filepath, 'rb') as fp:
                file = MIMEAudio(fp.read(), _subtype=subtype)
                fp.close()
        else:
            with open(filepath, 'rb') as fp:
                file = MIMEBase(maintype, subtype)
                file.set_payload(fp.read())
                fp.close()
                encoders.encode_base64(file)
        file.add_header('Content-Disposition', 'attachment', filename=filename)  # Добавляем заголовки
        msg.attach(file)  # Присоединяем файл к сообщению

class PutMessages:
    def __init__(self):
        self.count_sends = 0
        self.count_errors = 0

    def put(self, msgs_):
        smtp_server = c.SMTP_SERVER
        port = c.PORT
        some = c.G
        with SMTP(smtp_server, port) as s:
            login = c.FROM
            password = os.getenv(some)

            try:
                s.starttls()
                s.login(login, password)
            except SMTPAuthenticationError:
                logger.error('Не удалось соединиться с сервером\n'
                             'Проверьте правильность логина и пароля.')
                exit(1)

            for msg in msgs_:
                try:
                    s.send_message(msg)
                    self.count_sends += 1
                except Exception:
                    self.count_errors += 1
            logger.info(f'Отправлено {self.count_sends} писем')
            if self.count_errors != 0:
                logger.error(f'Не удалось отправить {self.count_errors} писем')

