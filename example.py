#!/usr/bin/python
import datetime

from send_reports_email import create_reports
from send_reports_email import send_email

date = datetime.date.today().strftime("%m%d%y")
save_location = r"F:\PWF Development\Planned Giving\Reports\Planned Giving Luminate Interest " + date
extension = ".xlsx"


def create_file():
    find_file_path = r"C:\Users\jhoover\Downloads\Planned"
    macro_name = "Mail_Merge.cleanMailMergePG"

    file = create_reports.CreateFile(find_file_path)
    create_reports.ExcelMacro(file.original_file, macro_name, save_location)
    file.delete_original_file(file.original_file)


def email_file():
    emailto = "mgregg@tpwf.org, nrosado@tpwf.org"
    subject = "Planned Giving Interest Update " + date
    body = "Hi Merrill,\r\rHere’s your excel for the week!\r\r-JordanBot"
    attach = save_location + extension

    email = send_email.SendEmail(emailto, subject, body, attach)
    email.send()


if __name__ == "__main__":
    create_file()
    email_file()