# %%
import pandas as pd
import win32com.client as win32
from dataclasses import dataclass
from contextlib import contextmanager

import tempfile
import shutil
import os


@dataclass
class ExcelInteract:
    file_name: str
    visible: bool = False

    def __post_init__(self):
        self.__excel = win32.gencache.EnsureDispatch('Excel.Application')
        self.__excel.Visible = self.visible

    def close_excel_app(self):
        self.__excel.Application.Quit()

    def __create_temp_copy(self):
        tmp = tempfile.NamedTemporaryFile(delete=False)
        shutil.copy2(self.file_name, tmp.name)
        return tmp

    def value_from_worksheet(self, wb):
        ws = wb.Worksheets(1)

    @contextmanager
    def open(self):
        workbook = None
        try:
            tmp_file = self.__create_temp_copy()
            workbook = self.__excel.Workbooks.Open(tmp_file.name)
            yield workbook
        finally:
            tmp_file.close()
            if workbook:
                workbook.Close(SaveChanges=False)


# %%
eint = ExcelInteract(r"C:\Users\t_chr\Desktop\SAMPLE.xlsx", True)
with eint.open() as wb:
    ws = wb.Worksheets("Hoja1")
    sheet_names = [sheet.Name for sheet in wb.Sheets]
    print(sheet_names)
    ws.Cells(*eint.parse_cel_name("A2")).Value = 10
    ws.Range("A2:C2").Value = [5, 6, 3]
    print(list(map(list, ws.Range("A1:E5").Value)))
    print(pd.DataFrame(ws.Range("A1:E5").Value))
    ""

# %%
from cellnameparser import CellSelectorParser
cellSel = "A1:A5"
cellSelParser = CellSelectorParser()
cellSelParser.parse_selection(cellSel)