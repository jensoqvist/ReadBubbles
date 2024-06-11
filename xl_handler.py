from os.path import isfile, join
import openpyxl
from dataframe_handler import DataFrameHandler


class XlHandler():
    """
    Class handling the .xlsx files created for the position numbers.\n\n

    Parameters:\n
        pdf_extractor: (PdfExtractor)\n
        pos_num_lenght: lenght of list of position numbers\n
        path: path to the .xlsx file\n

    Attributes:\n
        self.partnum = Part number from pdf_extractor\n
        self.revnum = Revision number from pdf_extractor\n
        self.sheet_name = Name of sheet to create in .xlsx"\n
        self.path = Path to .xlsx\n
        self.filename = Filename of .xlsx")\n
        self.pos_num_lenght = pos_num_lenght\n
        self.skip_rows = 5 number of rows to skip to where to look for dataframe\n 
        self.header_index = 0\n
        self.start_col = 1\n
        self.df_old = None\n
    """
    def __init__(self, pdf_extractor, pos_num_lenght, path) -> None:
        self.partnum = pdf_extractor.partnum
        self.revnum = pdf_extractor.revnum
        self.sheet_name = f"{self.partnum}_{self.revnum}"
        self.path = path
        self.filename = join(path, f"{self.partnum} Posnr.xlsx")
        self.pos_num_lenght = pos_num_lenght
        self.skip_rows = 5
        self.header_index = 0
        self.start_col = 1
        self.df_old = None

        self.get_old_dataframe()

    def check_file_exist(self):
        return isfile(self.filename)

    def create_new(self):
        wb = openpyxl.Workbook()
        wb['Sheet'].title = self.sheet_name
        wb.save(self.filename)

    def get_old_dataframe(self):
        if self.check_file_exist():
            for i in range(1, int(self.revnum)):
                self.partnum + "_" + str(int(self.revnum) - i)
                try:
                    self.df_old = DataFrameHandler().read_xl(self.filename, header= self.header_index, skiprows= self.skip_rows, dtype= str, sheet_name= self.partnum + "_" + str(int(self.revnum) - i))
                    self.df_old.drop(self.df_old.columns[self.df_old.columns.str.contains('unnamed',case = False)],axis = 1, inplace = True)
                    self.df_old.dropna(subset= ["Position Number"], inplace= True)
                    return
                except:
                    pass
        else:
            self.create_new()


if __name__ == "__main__":
    pass