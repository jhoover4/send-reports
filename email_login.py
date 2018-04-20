import smtplib
from secret import secret_dict


def login(un=secret_dict["global_un"], pw=secret_dict["email"]):
    """Logs into email using the TPWF email exchange client.
    Microsoft exchange technically uses its own protocol called MAPI.
    """

    login = []
    login.append(un)
    login.append(pw)

    #SMTP login. AKA sending emails.
    def auth_login():
        auth_login.smtpObj = smtplib.SMTP(secret_dict["email_server"], 587) #connect to server
        type(auth_login.smtpObj)
        auth_login.smtpObj.ehlo()
        auth_login.smtpObj.starttls() #establishes secure connection
        auth_login.smtpObj.login(un, pw) #logs into email

        return auth_login.smtpObj

    server = auth_login()
    login.append(server)

    print("Connection to smtp is successful.\n")

    return login
