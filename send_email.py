import os
import mimetypes
from email.mime.multipart import MIMEMultipart
from email import encoders
from email.mime.base import MIMEBase
from email.mime.image import MIMEImage
from email.mime.text import MIMEText

def send(un, pw, server, emailto=None, subject=None, body=None, attach=None):
    #in case this has been set by sendParams
    if emailto == None:
        emailto = input("Who are we sending this email to?: ")
    if attach == None:
        attach = input("What document are we sending?: ")
    if subject == None:
        subject = input("What is the subject line?: ")
    if body == None:
        body = input("What is the body text?: ")
    
    #these have to be provided every time
    username = un
    password = pw
    emailfrom = un

    msg = MIMEMultipart()
    msg["From"] = un
    msg["To"] = emailto
    msg["Subject"] = subject

    body_content = MIMEText(body, 'plain')

    #if file doesn't have an attachment
    if (attach == None or attach == ""):
        server.sendmail(emailfrom, emailto, msg.as_string())
    #if file does have an attachment  
    else:
        ctype, encoding = mimetypes.guess_type(attach)
        if ctype is None or encoding is not None:
            ctype = "application/octet-stream"

        maintype, subtype = ctype.split("/", 1)

        if maintype == "text":
            fp = open(attach)
            # Note: we should handle calculating the charset
            attachment = MIMEText(fp.read(), _subtype=subtype)
            fp.close()
        elif maintype == "image":
            fp = open(attach, "rb")
            attachment = MIMEImage(fp.read(), _subtype=subtype)
            fp.close()
        else:
            fp = open(attach, "rb")
            attachment = MIMEBase(maintype, subtype)
            attachment.set_payload(fp.read())
            fp.close()
            encoders.encode_base64(attachment)
        attachment.add_header("Content-Disposition", "attachment", filename=os.path.basename(attach))
        msg.attach(body_content)
        msg.attach(attachment)

        server.login(username,password)
        server.sendmail(emailfrom, emailto, msg.as_string())
    print("Success!")
    server.quit()
