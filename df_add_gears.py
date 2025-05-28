from settings import Settings

class GearParameters():
    """
    Class that adds gear parameters to a dataframe
    """
    def __init__(self, df, df_old) -> None:
        self.settings = Settings()
        self.df= df
        self.df_old= df_old
        self.columns= self.settings.data["Columns"]
        self.gear_count= None
        self.ids= []
        self.parameters = self.settings.data["Gear Settings"]["Parameters"]
        self.special = self.settings.data["Gear Settings"]["Special Char"]
      
    def run(self):
        self._count_gears()
        self._delete_old()
        self._set_gear_ids()
        for id in self.ids:
            self.add_parameters(id)
        print(f"Gear IDs: {self.ids}")

    def _count_gears(self):
        self.gear_count= self.df["Position Number"].value_counts().get("0200-0299")
        if self.gear_count == None:
            self.gear_count = 0
        print("GEAR COUNTS:" + str(self.gear_count))

    def _delete_old(self):
        self.df= self.df[self.df["Position Number"] != "0200-0299"]
        self.df= self.df[self.df["Position Number"] != "0299"]
        if self.df_old is not None:      
            index= self.df_old[(self.df_old["Position Number"] == "0299") & (self.df_old["Specification"] == "")].index
            self.df_old.drop(index, inplace= True)

    def _set_gear_ids(self):
        if self.df_old is not None: 
            self.ids = (self.df_old[(self.df_old["Gear ID"] != "")])["Gear ID"].unique().tolist()
        if self.gear_count > 0 and len(self.ids) != self.gear_count:
            if len(self.ids) > 0:
                self._delete_other_ids(self.ids)
                self.ids= []
            for i in range(self.gear_count):
                self.ids.append(input(f"Please input Gear ID {i + 1}: "))

    def _delete_other_ids(self, ids):
        if self.df_old is not None:
            ids.append("")
            for id in ids:
                for posnum, param in self.parameters.items():
                    index= self.df_old[(self.df_old["Position Number"] == posnum) & (self.df_old["Gear ID"] == id)].index
                    self.df_old.drop(index, inplace= True)

    def add_parameters(self, id):
        columns= self.columns
        columns["Type"]= "Gear Parameter"
        columns["Gear ID"]= id
        for posnum, param in self.parameters.items():
            columns["Position Number"] = posnum
            columns["Specification"] = param
            if self._check_classification(posnum):
                columns["Classification"]= "<M>"
                columns["Comment"]= "According to STD4567"
            else:
                columns["Classification"]= "<S>"
                columns["Comment"]= ""
            self.df = self.df._append(columns, ignore_index= True)

    def _check_classification(self, num):
        if "AK" in self.ids:
            if num in self.special["Gear"]:
                return True
        else:
            if num in self.special["Spline"]:
                return True

