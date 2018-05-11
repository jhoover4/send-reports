import mimetypes
import os
import smtplib
from email import encoders
from email.mime.base import MIMEBase
from email.mime.image import MIMEImage
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

from secret import secret_dict


class SendEmail:
    def __init__(self, emailto, subject, body, attach=None, emailfrom="jhoover@tpwf.org", emailtype="plain", **kwargs):
        self.emailfrom = emailfrom
        self.emailto = emailto
        self.subject = subject
        self.body = body
        self.attach = attach
        self.type = emailtype

        for key, value in kwargs.items():
            setattr(self, key, value)

    @staticmethod
    def login(un=secret_dict["global_un"], pw=secret_dict["email"]):
        """Logs into email using the TPWF email exchange client.
        Microsoft exchange technically uses its own protocol called MAPI.
        """

        login = [un, pw]

        # SMTP login. AKA sending emails.
        def auth_login():
            auth_login.smtpObj = smtplib.SMTP(secret_dict["email_server"], 587)  # connect to server
            type(auth_login.smtpObj)
            auth_login.smtpObj.ehlo()
            auth_login.smtpObj.starttls()  # establishes secure connection
            auth_login.smtpObj.login(un, pw)  # logs into email

            return auth_login.smtpObj

        server = auth_login()
        login.append(server)

        print("Connection to smtp is successful.\n")

        return login

    @staticmethod
    def parse_emailto(emailto):
        """Translate string of emailto to list for sending if string, else keep as list."""

        if "," in emailto and type(emailto) == 'string':
            emailto.replace(" ", "")
            emailto = emailto.split(",")
        else:
            emailto = [emailto]

        return emailto

    def send(self):
        """Basic functionality to send an email."""

        un, pw, server = self.login()
        sender = un

        msg = MIMEMultipart()
        msg["From"] = sender
        msg["To"] = self.emailto
        msg["Cc"] = sender
        msg["Subject"] = self.subject

        body_content = MIMEText(self.body, self.type)
        self.emailto = self.parse_emailto(self.emailto)

        # if file doesn't have an attachment
        if self.attach is None or self.attach is "":
            msg.attach(body_content)
            server.sendmail(self.emailfrom, self.emailto + [msg["Cc"]], msg.as_string())

        # if file does have an attachment
        else:
            ctype, encoding = mimetypes.guess_type(self.attach)
            if ctype is None or encoding is not None:
                ctype = "application/octet-stream"

            maintype, subtype = ctype.split("/", 1)

            if maintype == "text":
                fp = open(self.attach)
                attachment = MIMEText(fp.read(), _subtype=subtype)
                fp.close()
            elif maintype == "image":
                fp = open(self.attach, "rb")
                attachment = MIMEImage(fp.read(), _subtype=subtype)
                fp.close()
            else:
                fp = open(self.attach, "rb")
                attachment = MIMEBase(maintype, subtype)
                attachment.set_payload(fp.read())
                fp.close()
                encoders.encode_base64(attachment)
            attachment.add_header("Content-Disposition", "attachment", filename=os.path.basename(self.attach))
            msg.attach(body_content)
            msg.attach(attachment)

            server.login(un, pw)

            server.sendmail(self.emailfrom, self.emailto + [msg["Cc"]], msg.as_string())

        print("Success!")
        server.quit()
