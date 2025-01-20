

import re
from settings import Settings

class PositionNumber():
    """
    Class describing a position number found on drawing
    \n\n
    Attributes:\n
        posnum = The given position number on the drawing\n
        type = the type of specification deducted from the position number, based on DX standard for position numbers\n
        dic = Dictionary containing all the possible parameters a Position Number could have. Position number and Type hardcoded, the rest form settings.json
    """
    def __init__(self, posnum= None) -> None:
        self.posnum = posnum
        self.type = None
        self.settings = Settings()
        if posnum is not None:
            self.set_type()
        self.dic = {
            "Position Number": self.posnum,
            "Type": self.type,
        }
        self.dic.update(self.settings.data["Columns"])


    def set_type(self):
        for key, value in self.settings.data["PosTypes"].items():
            if re.match(value, self.posnum):
                self.type = key
        if self.type == None:
            self.type =  "NOT CATEGORISED"

    def list_values(self):
        """
        Returns a list of the values in the PositionNumber dictionary
        """
        return [*self.dic.values()]

    def list_keys(self):
        return [*self.dic.keys()]


if __name__ == "__main__":
    pass