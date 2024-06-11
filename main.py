"""
App that reads bubbles drawing .pdfs and creates standardised .xlsx files over positions found.
"""


from file_handler import FileHandler
from pdf_extract import PdfExctractor
from pos_numbers import PositionNumbers
from xl_handler import XlHandler
from dataframe_handler import DataFrameHandler
from xl_formater import XlFormater


def main():
    file_handler = FileHandler()
    print(f"Bubbles to Excel for {file_handler.name}")
    pdf_extractor = PdfExctractor(filehandler= file_handler)
    position_numbers_from_pdf = PositionNumbers(pdf_extractor.pos_numbers_clean)
    pos_num_lenght = position_numbers_from_pdf.lenght
    xlreader = XlHandler(pdf_extractor, pos_num_lenght, path= file_handler.path)
    df_handler = DataFrameHandler(df_old= xlreader.df_old, position_nums= position_numbers_from_pdf.position_numbers, xl_handler= xlreader)
    df_handler.to_xl()
    xl_formater = XlFormater(xlhandler= xlreader, df_handler= df_handler, duplicates= pdf_extractor.duplicates)
    print("Done!")
    input("Press Enter to exit...")


if __name__ == "__main__":
    main()