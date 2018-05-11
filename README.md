# Send Reports Through Email

Created to help sending emails created from excel reports. An example of how to use the module below.

## To run a macro on a downloaded file

```
import os
from send_reports_email import create_reports


save_location = r"~/save_test"
extension = ".xlsx"

find_file_path = os.path.expanduser("~") + r"\Downloads\"
macro_name = "Sample.sampleMacro"

file = create_reports.CreateFile(find_file_path)
create_reports.ExcelMacro(file.original_file, macro_name, save_location)
file.delete_original_file(file.original_file)

```

## To email that created file

```
from send_reports_email import send_email

emailto = "test@gmail.com"
subject = "Planned Giving Interest Update " + date
body = "Hello,\r\rYour file is ready!\r\r-JordanBot"
attach = save_location + extension

email = send_email.SendEmail(emailto, subject, body, attach)
email.send()

```