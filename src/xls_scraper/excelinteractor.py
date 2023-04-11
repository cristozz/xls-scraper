#%%
import win32com.client as win32
from dataclasses import dataclass
from contextlib import contextmanager

import tempfile, shutil, os

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
        shutil.copy2(self.file_name, tmp)
        return tmp
    
    def __col_title_to_number(self, col_title: str)->int:
        ans = 0
        for i in col_title:
            ans = ans * 26 + ord(i) - 64
        return ans

    def __parse_cel_name(self, cel_name: str)->tuple[str,int]:
        c = cel_name.rstrip('0123456789')
        r = cel_name[len(c):]
        return c, r

    def value_from_worksheet(self, wb):
        ws=wb.Worksheets(1)

    @contextmanager
    def open(self):
        workbook=None
        try:
            tmp_file = self.__create_temp_copy()
            workbook = self.__excel.Workbooks.Open(tmp_file.name)
            yield workbook
        finally:
            if workbook:
                workbook.Close(SaveChanges=False)


#%%
import pandas as pd
eint = ExcelInteract(r"C:\Users\t_chr\Desktop\SAMPLE.xlsx", True)
with eint.open() as wb:
    ws = wb.Worksheets("Hoja1")
    sheet_names = [sheet.Name for sheet in wb.Sheets]
    print(sheet_names)
    print(ws.Cells(1,1).Value)
    print(list(map(list, ws.Range("A1:E5").Value)))
    print(pd.DataFrame(ws.Range("A1:E5").Value))

    

