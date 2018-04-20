import sys
import os
import shutil
import glob
import mimetypes
from email.mime.multipart import MIMEMultipart
from email import encoders
from email.mime.base import MIMEBase
from email.mime.image import MIMEImage
from email.mime.text import MIMEText

import win32com.client
from pywintypes import com_error


from .email_login import login

# TODO:
# Need to work on extension
# REFACTOR!!
# Better functionality to grab correct file


class CreateFile:
    def find_trx_file(self):
        """"Finds file. Need to make this more abstract for recurring use."""
        try:
            all_trx_files = glob.glob(self.find_file_path + "*" + self.find_extension)  # get latest donations_by_trx csv
            latest_trx_file = max(all_trx_files, key=os.path.getctime)

            print("Latest downloaded file is " + latest_trx_file)
        except (FileNotFoundError, ValueError):
            print("No file found in '" + self.find_file_path + "' path.")
            sys.exit()

        return latest_trx_file

    def __init__(self, find_file_path, extension=".xlsx", find_extension =".csv",
                 **kwargs):
        """init must come after findTrxFile because it references function"""

        # super(Create_File, self).__init__()
        self.find_file_path = find_file_path
        self.exstention = extension
        self.find_extension = find_extension
        self.original_file = self.find_trx_file()

        for key, value in kwargs.items():
            setattr(self, key, value)

    def save_original_file(self, file_name, dest):
        """"Saves file from downloads folder to new location."""
        shutil.copy(file_name, dest)
        print("Saved " + file_name + " to " + dest)

    def delete_original_file(self, file_name):
        """"Deletes file from downloads folder"""
        try:
            os.remove(file_name)
            print("Deleted " + file_name + " file.")
        except (FileNotFoundError, ValueError):
            print("No file found named'" + file_name + "' to delete.")
            sys.exit()


class ExcelMacro:
    xl = None

    def open_excel(self):
        if self.xl is None: # if not already defined
            self.xl = win32com.client.gencache.EnsureDispatch("Excel.Application")
            self.xl.DisplayAlerts = False

    def close_excel(self):
        self.xl.DisplayAlerts = True
        self.xl.Application.Quit()
        self.xl = None

    def open_wkbk(self, file_name):
        try:
            wb = self.xl.Workbooks(file_name)
        except Exception as e:
            try:
                wb = self.xl.Workbooks.Open(file_name)
            except Exception as e:
                wb = None
                print("Trouble opening workbook")
                self.error_handling(e)

        return wb

    def close_wkbk(self, wb):
        try:
            wb.Close()
            wb = None
        except com_error:
            # need to find a way to distinguish pywinerrors here
            print("Macro must have already closed workbooks, moving on...")

    def run_macro(self):
        """"Runs macro on excel file. Saves to new location."""

        wb = self.open_wkbk(self.file_name)
        personalwkbk = self.xl.Workbooks.Open(
            "{0}\\AppData\\Roaming\\Microsoft\\Excel\\XLSTART\\PERSONAL.XLSB".format(os.path.expanduser('~')))
        self.xl.Application.Run('PERSONAL.XLSB!' + self.macro_name)
        if self.save_location != "":
            wb.SaveAs(Filename=self.save_location, FileFormat="51")  # 51 is xlOpenXMLWorkbook
        try:
            wb.Close()
            wb = None
            personalwkbk.Close()
            personalwkbk = None
        except com_error:
            # need to find a way to distinguish pywinerrors here
            print("Macro must have already closed workbooks, moving on...")

        self.close_excel()

    def __init__(self, file_name, macro_name, save_location=""):
        """Opens an excel instance on object creation"""

        self.open_excel()
        self.file_name = file_name
        self.macro_name = macro_name
        self.save_location = save_location

        self.run_macro()

    def save_as_pdf(self, path_to_pdf):
        """This can be refactored to print out a better pdf with exact specs."""
        self.open_excel()

        try:
            print(self.save_location + ".xlsx")

            wb = self.open_wkbk(self.save_location + ".xlsx")
            ws = wb.Worksheets[1]
            ws.PageSetup.Orientation = 1
            ws.PageSetup.FitToPagesTall = 1
            ws.PageSetup.FitToPagesWide = 1

            ws.ExportAsFixedFormat(0, path_to_pdf)
            self.close_wkbk(wb)

            print("Successfully created PDF.")
            self.close_excel()

        except Exception as e:
            print("Error prevented PDF from being created.")
            self.error_handling(e)

    def error_handling(self, error):
        print(error)

        print("Excel is closing because of an error.")
        self.close_excel()


class SendEmail:
    def __init__(self, emailto, subject, body, attach=None, emailfrom="jhoover@tpwf.org", type="plain", **kwargs):
        self.emailfrom = emailfrom
        self.emailto = emailto
        self.subject = subject
        self.body = body
        self.attach = attach
        self.type = type

        for key, value in kwargs.items():
            setattr(self, key, value)

    def send(self):
        """This was moved/copied from send_email.py. Not sure if I'll delete that module altogether and combine it with this one."""

        un, pw, server = login()
        sender = un

        msg = MIMEMultipart()
        msg["From"] = sender
        msg["To"] = self.emailto
        msg["Cc"] = sender
        msg["Subject"] = self.subject

        body_content = MIMEText(self.body, self.type)

        # translate emailto into list for sending
        if "," in self.emailto:
            self.emailto.replace(" ", "")
            self.emailto = self.emailto.split(",")
        else:
            self.emailto = [self.emailto]
        # if file doesn't have an attachment
        if (self.attach is None or self.attach == ""):
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
                # Note: we should handle calculating the charset
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
