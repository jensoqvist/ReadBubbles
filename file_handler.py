import os
import sys

class FileHandler():
    """
    Class handling input file from user.\n\n

    Checks if a path has been given as a system argument otherwise asks user to input a path.\n
    Checks if path is valid, is it a file and doeas it have extention .pdf.\n\n

    Attributes:\n
        self.valid_path = False\n
        self.args_checked = False\n
        self.args = system arguments\n
        self.fullpath = Full path from user\n
        self.path = Directory path\n
        self.name = Filename\n
        self.ext = File extention\n
        self.dir = List of files in self.path
    """
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
        """
        Loops get_path() and check_if_valid() until an valid path is given.
        """
        while not self.valid_path:
            self.get_path()
            self.check_if_valid()

    def get_path(self):
        """
        Checks if a system argument has been given, if it has sets self.args to True.\n
        Sets self.fullpath
        """
        if len(self.args) > 1 and self.args_checked == False:
            self.fullpath = os.path.abspath(self.args[1])
            self.args_checked = True
        else: 
            self.fullpath = os.path.abspath(input("\nPlease input path to pdf: ").strip('"'))

    def split_path(self):
        """
        Splits path into parts and sets:\n
        self.path\n
        self.name\n
        self.ext\n
        self.dir
        """
        split_name = os.path.split(self.fullpath)
        self.path = os.path.dirname(self.fullpath) + os.path.sep
        self.name = split_name[-1]
        self.ext = os.path.splitext(self.fullpath)[1]
        self.dir = os.listdir(self.path)

    def check_if_valid(self):
        """
        Checks if path is valid, is it a file and doeas it have extention .pdf.
        """
        self.split_path()
        if self.ext.lower() == ".pdf" and os.path.isfile(self.fullpath):
            self.valid_path = True
        else:
            print("Input path not valid, please input valid path to PDF!")



if __name__ == "__main__":
    pass


