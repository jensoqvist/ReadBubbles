'''
Handles the exctraction and cleaning of the text from the position number drawing PDF
\n\n
Contains:\n
Class PdfExctractor
'''

from pdfminer.high_level import extract_text
from pdfminer.layout import LAParams
import re


class PdfExctractor():
    '''
    Class that handles the exctraction and cleaning of the text from the position number drawing PDF\n\n

     Attributes:\n
        partnum = The Partnumber extracted from the position number drawing\n
        revnum = The Revision number extracted from the position number drawing\n
        pos_numbers_clean = A clean list of the position numbers on the position number drawing\n
        duplicates = Any duplicate position number found\n
        _pdf = the pdf file to exctract position numbers from\n
        _text = extracted text\n
        _pos_numbers = An list of the position numbers on the position number drawing before cleaning\n\n

    '''
    def __init__(self, filehandler) -> None:
        self.filehandler = filehandler
        self._pdf = self.filehandler.fullpath
        self.listdir = self.filehandler.dir
        self._text = self._extract_text(self._pdf)
        self.partnum = ""
        self.revnum = ""
        self._pos_numbers = []
        self.pos_numbers_clean = []
        self.duplicates = []
        self.check_for_multipdf()
        self._set_pos_numbers()
        self._cleaning()
        self._set_duplicates()

    def _extract_text(self, pdf):
        return extract_text(pdf, laparams= LAParams(detect_vertical= True,  boxes_flow= -0.5, line_overlap= 0.05))

    def check_for_multipdf(self):
        if re.search("\d{7}.*Sheet_\d\.pdf", self._pdf):
            for file in self.listdir:
                if re.search("\d{7}.*Sheet_\d\.pdf", file) and self._pdf != self.filehandler.path + file:
                    self._text += extract_text(self.filehandler.path + file)

    def find_rotated(self):
        '''
        Finds positions that are not strictly horizontal or vertical.
        '''
        rotated = re.findall("\n\n\d(?:\n{2}|\s)\d(?:\n{2}|\s)?\d(?:\n{2}|\s)\d", self._text, re.MULTILINE)
        rotated_clean = []
        for x in sorted(rotated):
            pos = x.replace("\n", "").replace(" ", "")
            if re.match("0(?:3|4)\d{2}", pos):
                rotated_clean.append(pos)
            elif re.match("2\d{3,3}", pos) and not re.match("2\d{4}", pos):
                rotated_clean.append(pos)
        return(rotated_clean)
        

    def _set_pos_numbers(self):
        '''
        Method that sets _pos_numbers, and calls _set_partnum, and _set_revnum
        '''
        last_line = ""
        numbers = []
        for line in self._text.split(): 
            if last_line == "Pos":
                self._set_partnum(line)
            elif re.match("Rev_", line):
                self._set_revnum(line)
            elif last_line == "ISO" and line == "6411":
                continue # Skip ISO 6411 - R
            if re.search('\d{4}', line):
                numbers.append(line)
            last_line = line
        self._pos_numbers = numbers
        self._pos_numbers += self.find_rotated()

    def _set_partnum(self, line):
        self.partnum = re.match('\d{7}', line).group(0)

    def _set_revnum(self, line):
        self.revnum = line[-1]

    def _set_duplicates(self):
        for pos in self.pos_numbers_clean:
            if self.pos_numbers_clean.count(pos) > 1 and pos != '0200-0299':
                self.pos_numbers_clean.remove(pos)
                self.duplicates.append(pos)

    def _cleaning(self):
        '''
        Set pos_numbers_clean from _pos_numbers after cleaning
        '''
        clean_numbers = []
        for pos in self._pos_numbers: 
            if re.match('[^\d]\d{4}', pos):
                pos = pos[1:]
            if pos == '0200-0299':
                clean_numbers.append(pos)
                continue
            elif re.search('\d{4}-\d{2}-\d{2}', pos):
                continue # Do not add date to clean list
            elif re.search('\d{7}', pos):
                continue # Gear data sheet, do not add
            elif re.search('SV\d{4}', pos):
                continue # SV
            elif re.search('TB-?\d{4}', pos):
                continue # No TB
            elif re.search('\d{4}-R', pos):
                continue # Thread
            elif re.search('\d{4}-\d{4}', pos): # EX. 0103-0106
                lst = [i for i in range(int(re.search('\d{4}', pos).group(0)), int(pos[-4:]) + 1)]
                for i in lst:
                    clean_numbers.append(str(i).zfill(4))
            elif re.search('\d{4}\.?\d?\/\d{4}\.?\d?', pos): #EX 0101/0102
                clean_numbers += pos.split("/")
            elif re.search('\d{4}.\d\/\d{4}.\d\/\d{4}.\d', pos): #EX 1101.1/1102.1/1103.1
                clean_numbers += pos.split("/")
            elif re.search('\d{4}.\d[-|\/]\d[^\d]', pos):     #EX 1101.1-3 (1101.1, 1101.2, 1101.3)  
                match = re.search('\d{4}.', pos).group(0)
                for i in range(int(pos[-3]), int(pos[-1]) + 1):
                    clean_numbers.append(match + str(i))
            elif re.search('.?\d{4}\.\d-\d{4}\.\d', pos): #EX 1101.1-1103.1 (1101.1, 1102.1, 1103.1)
                match = re.search('\d{4}.\d-\d{4}.\d', pos).group(0)
                if match[5] == match[-1]:
                    lst = [i for i in range(int(match[0:4]), int(pos[7:11]) + 1)]
                    for i in lst:
                        clean_numbers.append(str(i) + "." + match[5])
            else:
                clean_numbers.append(pos)
        self.pos_numbers_clean = clean_numbers        


if __name__ == "__main__":
    pass