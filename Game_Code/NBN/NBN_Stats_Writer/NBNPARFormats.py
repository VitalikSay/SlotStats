from Basic_Code.Stats_Writer.BasicPARFormats import BasicPARFormats
import xlsxwriter as xl
import json as j
from Basic_Code.Stats_Writer.BasicGraphMaker import Colors


class NBNPARFormats(BasicPARFormats):
    def __init__(self, book: xl.Workbook):
        super().__init__(book)