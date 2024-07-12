import pandas as pd
from pos_number import PositionNumber

class DataFrameHandler():
    """
    Class that creates a pandas dataframe over PositionNumbers found on a Bubble drawing.\n
    Checks if there is an old .xlsx and if so extracts the latest dataframe from it.\n\n

    If an old data frame is found, compares the position numbers in the newly created dataframe and the old.\n
    Keeps any old information added to the old dataframe.\n\n

    Attributes:\n
        self.col_names = Column names in the dataframe, determend by the keys defined in PositionNumber() \n
        self.df_old = The old dataframe if found \n
        self.position_numbers = Position Numbers found on bubble drawing\n
        self.xl = xl_handler, Handling of .xlsx file\n
        self.new = List of new position numbers\n
        self.removed = List of removed position numbers
    """
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
            self._compare_columns()
            self._add_manual()
            self._compare_rows()
        
    def create_new_data_frame(self):
        """
        Creation of new dataframe with columns defined in self.col_names
        """
        return pd.DataFrame(columns= self.col_names)

    def read_xl(self, *args, **kwargs):
        return pd.read_excel(*args, **kwargs)

    def to_xl(self):
        """
        Appends dataframe to .xlsx file in self.xl
        """
        with pd.ExcelWriter(self.xl.filename, engine= 'openpyxl', mode= 'a', if_sheet_exists= 'replace') as writer:
            self.df.to_excel(writer, index=False, startrow= self.xl.skip_rows, startcol= self.xl.start_col, sheet_name= self.xl.sheet_name)

    def _fill_df(self):
        """
        Fills the dataframe self.df with self.position_numbers
        """
        if self.position_numbers is not None:
            self.df = pd.DataFrame(self.position_numbers, columns= self.col_names)

    def _compare_columns(self):
        """
        Compares the entire Dataframe to an old Dataframe
        """
        old_columns = self.df_old.columns
        for column in self.df.columns:
            if column not in old_columns:
                self.df_old.insert(self.df.columns.get_loc(column), column, None)
        for column in old_columns:
            if column not in self.df.columns:
                self.df_old.drop(columns= [column], inplace= True)

    def _add_manual(self):
        self.df = pd.concat([self.df, self.df_old.loc[self.df_old["Manually Added"] == "Yes"]])            
        self.df = self.df.sort_values("Position Number")
        self.df = self.df.reset_index(drop= True)

    def _compare_rows(self):
        """
        Compare Position Numbers to old Position Numbers.\n
        Keeps information that might have been added in the old dataframe.
        """
        new = []
        removed = []
        for index, row in enumerate(self.df["Position Number"].values):
            if row not in self.df_old["Position Number"].values:
                new.append(row)
            else:
                self.df[(self.df == row).any(axis= 1)] = self.df_old[(self.df_old == row).any(axis= 1)].values.tolist()
        for index, row in enumerate(self.df_old["Position Number"].values):
            if row not in self.df["Position Number"].values:
                removed.append(row)


        self.new = new
        self.removed = removed
      

if __name__ == "__main__":
    pass