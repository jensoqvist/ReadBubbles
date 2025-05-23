import pandas as pd
from pos_number import PositionNumber
from df_add_gears import GearParameters

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
    def __init__(self, df_old= None, position_nums = None, xl_handler= None, run= False) -> None:
        self.col_names = PositionNumber().list_keys()
        self.df_old = df_old
        self.position_numbers = position_nums
        self.xl = xl_handler
        self.df = self.create_new_data_frame()
        self.new = []
        self.removed = []
        if run == True:
            self._run()
             
    def _run(self):
        self._fill_df()
        self._rename_old_columns() 
        self._check_for_gears()
        if self.df_old is not None:  
            self._compare_columns() 
            self._add_manual()
            self._compare_rows()     
        self._sort_df()  
        
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
            self.df["CMS Audit"] = pd.to_numeric(self.df["CMS Audit"], downcast= 'integer')

    def _check_for_gears(self):  
        gear= GearParameters(df= self.df, df_old= self.df_old)
        gear.run()
        self.df= gear.df
        self.df_old= gear.df_old

    def _rename_old_columns(self):
        if self.df_old is not None: 
            if "MSA3" in self.df_old.columns:
                self.df_old.rename(columns={"MSA3": "MSA2/3"}, inplace=True)
            if "GPS Specification" in self.df_old.columns:
                self.df_old.rename(columns={"GPS Specification": "Specification"}, inplace= True)

    def _compare_columns(self):
        """
        Compares the column of the Dataframe to the old Dataframe
        """
        old_columns = self.df_old.columns.to_list()
        columns = self.df.columns.to_list()
        for column in columns:
            if column not in old_columns:
                self.df_old.insert(self.df.columns.get_loc(column), column, None)
        for column in old_columns:
            if column not in self.df.columns:
                self.df_old.drop(columns= [column], inplace= True)
        self.df_old = self.df_old[columns]

    def _add_manual(self):
        self.df = pd.concat([self.df, self.df_old.loc[self.df_old["Manually Added"] == "Yes"]])            
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
                self._compare_pos(row, index)
        for index, row in enumerate(self.df_old["Position Number"].values):
            if row not in self.df["Position Number"].values:
                removed.append(row)
        self.new = new
        self.removed = removed

    def _compare_pos(self, row, index):
        new_shape= self.df[self.df["Position Number"] == row].shape
        old_shape= self.df_old[self.df_old["Position Number"] == row].shape
        if new_shape == old_shape:
            self.df[self.df["Position Number"] == row] = self.df_old[self.df_old["Position Number"] == row].values
        else:
            self.df.drop(index)
            for index, r in enumerate(self.df_old[self.df_old["Position Number"] == row].values):
                self.df.loc[self.df.index.max() + 1] = r

    def _sort_df(self):
        self.df = self.df.sort_values(["Gear ID", "Position Number"], ascending=[True, True], na_position="first")


      

if __name__ == "__main__":
    pass