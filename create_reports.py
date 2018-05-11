import glob
import os
import shutil
import sys

import win32com.client
from pywintypes import com_error


# TODO: Work on extension handling.
# TODO: Better functionality to grab correct file.


class CreateFile:
    def __init__(self, find_file_path, extension=".xlsx", find_extension=".csv",
                 **kwargs):

        self.find_file_path = find_file_path
        self.exstention = extension
        self.find_extension = find_extension
        self.original_file = self.find_trx_file()

        for key, value in kwargs.items():
            setattr(self, key, value)

    def find_trx_file(self):
        """"Finds file. Need to make this more abstract for recurring use."""

        try:
            # get latest donations_by_trx csv
            all_trx_files = glob.glob(self.find_file_path + "*" + self.find_extension)
            latest_trx_file = max(all_trx_files, key=os.path.getctime)

            print("Latest downloaded file is " + latest_trx_file)
        except (FileNotFoundError, ValueError):
            print("No file found in '" + self.find_file_path + "' path.")
            sys.exit()

        return latest_trx_file

    @staticmethod
    def save_original_file(file_name, dest):
        """"Saves file from downloads folder to new location."""

        shutil.copy(file_name, dest)
        print("Saved " + file_name + " to " + dest)

    @staticmethod
    def delete_original_file(file_name):
        """"Deletes file from downloads folder"""
        try:
            os.remove(file_name)
            print("Deleted " + file_name + " file.")
        except (FileNotFoundError, ValueError):
            print("No file found named'" + file_name + "' to delete.")
            sys.exit()


class ExcelMacro:
    xl = None

    def __init__(self, file_name, macro_name, save_location=""):
        """Opens an excel instance on object creation."""

        self.open_excel()
        self.macro_name = macro_name
        self.save_location = save_location
        self.file_name = file_name

        self.run_macro()

    @staticmethod
    def find_macro_dependencies(partial_file_name, folder=""):
        """If dependencies are needed, find them in the downloads folder."""

        if folder is "":
            folder = os.path.expanduser("~") + "\\Downloads\\"

        file = glob.glob(folder + partial_file_name + "*")[0]

        if not file:
            raise FileNotFoundError("Base file not found.")
        else:
            return file

    def open_excel(self):
        if self.xl is None:  # if not already defined
            self.xl = win32com.client.gencache.EnsureDispatch("Excel.Application")
            self.xl.DisplayAlerts = False

    def close_excel(self):
        self.xl.DisplayAlerts = True
        self.xl.Application.Quit()
        self.xl = None

    def open_wkbk(self, file_name):
        try:
            wb = self.xl.Workbooks(file_name)
        except Exception:
            try:
                wb = self.xl.Workbooks.Open(file_name)
            except Exception as e:
                wb = None
                print("Trouble opening workbook")
                self.error_handling(e)

        return wb

    @staticmethod
    def close_wkbk(wb):
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
