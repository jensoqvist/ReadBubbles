import os
import sys

class FileHandler():
    def __init__(self) -> None:
        self.valid_path = False
        self.args_checked = False
        self.args = sys.argv
        self.fullpath = ""
        self.path = ""
        self.name = ""
        self.ext = ""
        self.dir = ""

        self.check_user_input()
              
    def check_user_input(self):
        while not self.valid_path:
            self.get_path()
            self.check_if_valid()

    def get_path(self):
        if len(self.args) > 1 and self.args_checked == False:
            self.fullpath = os.path.abspath(self.args[1])
            self.args_checked = True
        else: 
            self.fullpath = os.path.abspath(input("\nPlese input path to pdf: ").strip('"'))

    def split_path(self):
        split_name = os.path.split(self.fullpath)
        self.path = os.path.dirname(self.fullpath) + os.path.sep
        self.name = split_name[-1]
        self.ext = os.path.splitext(self.fullpath)[1]
        self.dir = os.listdir(self.path)

    def check_if_valid(self):
        self.split_path()
        if self.ext.lower() == ".pdf" and os.path.isfile(self.fullpath):
            self.valid_path = True
        else:
            print("Input path not valid, please input valid path to PDF!")



if __name__ == "__main__":
    pass


