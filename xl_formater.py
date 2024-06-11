import openpyxl
from openpyxl.worksheet.table import Table, TableStyleInfo
from openpyxl.styles import Alignment
from openpyxl.styles.borders import Border, Side
from openpyxl.worksheet.datavalidation import DataValidation
from settings import Settings

class XlFormater():
    """
    Class to give .xlsx file the correct format after the dataframe has been added to it.\n\n

    Parameters:\n
        xlhandler\n 
        df_handler\n
        duplicates\n\n

    Attributes:\n
        self.settings = Settings()\n
        self.xl = xlhandler\n
        self.df_handler = df_handler\n
        self.df = df_handler.df\n
        self.duplicates = duplicates\n
        self.wbook = openpyxl.load_workbook(self.xl.filename)\n
        self.sheet = self.wbook[self.xl.sheet_name]\n\n
    Table settings:\n
        self.table_start_col_index = 1 + self.xl.start_col\n
        self.table_start_row_index = self.xl.skip_rows + 1\n
        self.table_start = f"{openpyxl.utils.get_column_letter(self.table_start_col_index)}{self.table_start_row_index}"\n
        self.table_end_col_index = self.df.shape[1] + self.xl.start_col\n
        self.table_end_row_index = len(self.df) + self.xl.skip_rows + 1\n
        self.table_end = f"{openpyxl.utils.get_column_letter(self.table_end_col_index)}{self.table_end_row_index}"\n
    
    """
    def __init__(self, xlhandler, df_handler, duplicates) -> None:
        self.settings = Settings()
        self.xl = xlhandler
        self.df_handler = df_handler
        self.df = df_handler.df
        self.duplicates = duplicates
        self.table_start_col_index = 1 + self.xl.start_col
        self.table_start_row_index = self.xl.skip_rows + 1
        self.table_start = f"{openpyxl.utils.get_column_letter(self.table_start_col_index)}{self.table_start_row_index}"
        self.table_end_col_index = self.df.shape[1] + self.xl.start_col
        self.table_end_row_index = len(self.df) + self.xl.skip_rows + 1
        self.table_end = f"{openpyxl.utils.get_column_letter(self.table_end_col_index)}{self.table_end_row_index}"
        self.wbook = openpyxl.load_workbook(self.xl.filename)
        self.sheet = self.wbook[self.xl.sheet_name]

        self._add_table()
        self._add_headings()
        self._add_lists()
        self._adjust_sizes()
        self._alignment()
        self._add_borders()
        self._data_validation()       
        self._hide_sheets()
        self.wbook.active = self.sheet
        self._save()

    def _add_table(self):
        self.table = Table(displayName=f"Positions_REV_{self.xl.revnum}", ref=f'{self.table_start}:{self.table_end}')
        style = TableStyleInfo(name="TableStyleMedium16", showFirstColumn=False,
                       showLastColumn=False, showRowStripes=True, showColumnStripes=False)
        self.table.tableStyleInfo = style
        self.sheet.add_table(self.table)

    def _add_headings(self):
        self.sheet['B2'].value = "PART NUMBER"
        self.sheet['C2'].value = self.xl.partnum
        self.sheet['B3'].value = "REVISION"
        self.sheet['C3'].value = self.xl.revnum

    def _add_lists(self):
        if len(self.duplicates) > 0:
            self.sheet['G2'].value = "WARNING! THE FOLLOWING DUPLICATES HAS BEEN FOUND:"
            self.sheet['H2'].value = str(self.duplicates)
        self.sheet['G3'].value = "NEW POSITION NUMBERS:"
        self.sheet['H3'].value = str(self.df_handler.new)
        self.sheet['G4'].value = "REMOVED POSITION NUMBERS:"
        self.sheet['H4'].value = str(self.df_handler.removed)

    def _adjust_sizes(self):
        MAX_LEN = 8
        LENGHT_LIMIT = 50
        for col in self.sheet.columns: 
            column = col[0].column_letter
            for cell in col:
                try:
                    cell_len = len(str(cell.value))
                    if cell_len > LENGHT_LIMIT:
                        cell_len = LENGHT_LIMIT
                    if cell_len > MAX_LEN:
                        MAX_LEN = cell_len
                except:
                    pass
            adjusted_width = (MAX_LEN + 2) * 1.2
            self.sheet.column_dimensions[column].width = adjusted_width

    def _alignment(self):
        #Align Position Number, Col 1
        for r in range(self.table_start_row_index + 1, self.sheet.max_row + 1):
            self.sheet.cell(r, self.table_start_col_index).alignment = Alignment(horizontal='right')

    def _add_borders(self):
        thin_side = Side(style= 'thin')
        thin_border = Border(left= thin_side, right= thin_side, top= thin_side, bottom= thin_side)
        for row in range(self.table_start_row_index, self.sheet.max_row + 1):
            for col in range(self.table_start_col_index, self.table_end_col_index + 1):
                self.sheet.cell(row= row, column= col).border = thin_border

    def _data_validation(self):
        for key, value in self.settings.data["Data Validation"].items():
            validation_string = ", ".join(value)
            col_index = self.df.columns.get_loc(key) + self.table_start_col_index
            dv = DataValidation(type= "list", formula1= f'"{validation_string}"', allow_blank= True)
            dv.error = "Invalid Entry"
            dv.errorTitle = "Invalid Entry"
            self.sheet.add_data_validation(dv)
            for r in range(self.table_start_row_index + 1, self.sheet.max_row + 1):
                dv.add(openpyxl.utils.get_column_letter(col_index) + str(r))

    def _hide_sheets(self):
        for sheet in self.wbook.get_sheet_names():
            if sheet != self.xl.sheet_name:
                self.wbook[sheet].views.sheetView[0].tabSelected = False
                self.wbook[sheet].sheet_state = "hidden"

    def _save(self):
        self.wbook.save(self.xl.filename)



if __name__ == "__main__":
    pass