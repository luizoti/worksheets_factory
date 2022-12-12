"""This is the main software file."""

import string
from os.path import abspath, join

import pandas as pd
from memory_profiler import profile
from pyexcelerate import Workbook


class WorkSheetsFactory:
    def __init__(self, work_sheet_name: str, sheet_diretory, sheet_file_name, work_sheet_header: list, work_sheet_rows: list[list]):
        self.work_sheet_name: str = work_sheet_name
        self.sheet_diretory: str = sheet_diretory
        self.sheet_file_name: str = sheet_file_name

        self.sheet_data = [work_sheet_header]
        self.sheet_data += work_sheet_rows

    @profile
    def pandas(self, rows_data, custon_file_name=None):
        """Generate a .xlsx with pandas."""
        print("Running pandas solution.")
        file_name = custon_file_name if custon_file_name else self.sheet_file_name
        sheet_file_path = join(self.sheet_diretory, file_name)
        if ".xlsx" not in sheet_file_path:
            sheet_file_path = f"{sheet_file_path}.xlsx"
        pd.DataFrame(rows_data).to_excel(sheet_file_path, index=False,
                                         header=False)

    @profile
    def pyexcelerate(self, rows_data, custon_file_name=None):
        """Generate a .xlsx with pyexcelerate."""
        print("Running pyexcelerate solution.")
        file_name = custon_file_name if custon_file_name else self.sheet_file_name
        sheet_file_path = join(self.sheet_diretory, file_name)
        if ".xlsx" not in sheet_file_path:
            sheet_file_path = f"{sheet_file_path}.xlsx"
        wb = Workbook()
        wb.new_sheet(self.work_sheet_name, data=rows_data)
        wb.save(sheet_file_path)

    def chunked(self, size_of_chunk=200000, pandas=False):
        for index, chunk in enumerate(
                [self.sheet_data[x:x + size_of_chunk] for x in range(0, len(self.sheet_data), size_of_chunk)]):
            sheet_file_path = abspath(f"{self.sheet_file_name}_{index}")
            if ".xlsx" not in sheet_file_path:
                sheet_file_path = f"{sheet_file_path}.xlsx"
            print(f"Criando: {sheet_file_path}")
            if pandas:
                self.pandas(custon_file_name=sheet_file_path, rows_data=chunk)
            else:
                self.pyexcelerate(custon_file_name=sheet_file_path, rows_data=chunk)


if __name__ == "__main__":
    sheet_name = "Sheet_name_test"
    work_dir = abspath(".")
    sheet_header = [x for x in string.ascii_lowercase]
    sheet_rows = [list(range(len(sheet_header)))] * 1048574
    # sheet_rows = [list(range(len(sheet_header)))] * 5

    # Colunas: 26 (alfabeto)
    # Linhas? 1048574
    # Running pandas' solution.
    # FUNCTION pandas TIME: 1673.4213s
    # Running pyexcelerate solution.
    # FUNCTION pyexcelerate TIME: 305.1016s

    wsf = WorkSheetsFactory(
        work_sheet_name=sheet_name,
        sheet_diretory=work_dir,
        sheet_file_name="output",
        work_sheet_header=sheet_header,
        work_sheet_rows=sheet_rows
    )
    # wsf.pandas()
    # wsf.pyexcelerate()
    wsf.chunked()
