import pandas as pd
from pos_number import PositionNumber

class DataFrameHandler():
    def __init__(self, df_old= None, position_nums = None, xl_handler= None) -> None:
        self.col_names = PositionNumber().list_keys()
        self.df_old = df_old
        self.position_numbers = position_nums
        self.xl = xl_handler
        self.df = self.create_new_data_frame()
        self.new = []
        self.removed = []
        self._fill_df()
        if self.df_old is not None:    
            self._compare_df()
            self._compare()
        
    def create_new_data_frame(self):
        return pd.DataFrame(columns= self.col_names)

    def read_xl(self, *args, **kwargs):
        return pd.read_excel(*args, **kwargs)

    def to_xl(self):
        with pd.ExcelWriter(self.xl.filename, engine= 'openpyxl', mode= 'a', if_sheet_exists= 'replace') as writer:
            self.df.to_excel(writer, index=False, startrow= self.xl.skip_rows, startcol= self.xl.start_col, sheet_name= self.xl.sheet_name)

    def _fill_df(self):
        if self.position_numbers is not None:
            self.df = pd.DataFrame(self.position_numbers, columns= self.col_names)

    def _compare_df(self):
        old_columns = self.df_old.columns
        for column in self.df.columns:
            if column not in old_columns:
                self.df_old.insert(self.df.columns.get_loc(column), column, None)
        for column in old_columns:
            if column not in self.df.columns:
                self.df_old.drop(columns= [column], inplace= True)

    def _compare(self):
        new = []
        removed = []
        for row in self.df["Position Number"].values:
            if row not in self.df_old["Position Number"].values:
                new.append(row)
            else:
                self.df[(self.df == row).any(axis= 1)] = self.df_old[(self.df_old == row).any(axis= 1)]
        for row in self.df_old["Position Number"].values:
            if row not in self.df["Position Number"].values:
                removed.append(row)
        self.new = new
        self.removed = removed


if __name__ == "__main__":
    pass