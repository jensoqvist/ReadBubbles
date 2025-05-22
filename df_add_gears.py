from settings import Settings

class GearParameters():
    """
    Class that adds gear parameters to a dataframe
    """
    def __init__(self, df, gear_id) -> None:
        self.settings = Settings()
        self.df= df
        self.columns= self.settings.data["Columns"]
        self.id= gear_id
        self.parameters = self.settings.data["Gear Settings"]["Parameters"]
        self.special = self.settings.data["Gear Settings"]["Special Char"]
        self._add_parameters()


    def _add_parameters(self):
        columns= self.columns
        columns["Type"]= "Gear Parameter"
        columns["Gear ID"]= self.id
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
        if "AK" in self.id:
            if num in self.special["Gear"]:
                return True
        else:
            if num in self.special["Spline"]:
                return True
