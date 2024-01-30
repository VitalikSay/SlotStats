from Basic_Code.Stats_Writer.BasicPARFormats import BasicPARFormats
from Basic_Code.Stats_Writer.BasicPARSheet import BasicPARSheet
from Basic_Code.Utils.BasicPathHadler import BasicPathHandler
from Basic_Code.Basic_Structures.BasicSlot import BasicSlot
from Basic_Code.Basic_Structures.BasicStatistics import BasicStatistics
from Basic_Code.Basic_Calculator.BasicStatsCalculator import BasicStatsCalculator
import xlsxwriter as xl
import json as j


class BasicBook:
    def __init__(self, game_name_short: str):
        self._book = xl.Workbook

        self._par_formats = BasicPARFormats
        self._game_name_short = game_name_short
        self._path_handler = BasicPathHandler()
        self._file_name = "PAR_" + game_name_short + "_base.xlsx"
        self._save_book_path = self._path_handler.GetPARPath(game_name_short, self._file_name)

    def WriteExel(self):
        self._InitBook()

        self._WritePARSheet(self._game_name_short, '86')
        self._WritePARSheet(self._game_name_short, '90')
        self._WritePARSheet(self._game_name_short, '93')

        self._CloseBook()

    def _InitBook(self):
        self._book = xl.Workbook(self._save_book_path)
        self._par_formats = self.MakePARFormats()
        self._par_formats.InitFormats()

    def _WritePARSheet(self, game_name_short: str, version_name: str, simulation_tags: list=['usual']):
        json_stats_path = self._path_handler.GetJsonStatsPath(game_name_short,
                                                              version_name,
                                                              simulation_tags)
        data = j.load(open(json_stats_path, 'r'))
        slot = self.MakeSlot()
        slot.ReadSlot_Json(data)
        stats = self.MakeStatistics()
        stats.ReadStatistics(data)
        calculator = self.MakeCalculator(slot, stats)
        calculator.CalcStats()
        par_sheet = self.MakePARSheet(slot, stats, calculator)
        par_sheet.WritePARSheet()

    def MakeSlot(self):
        return BasicSlot()

    def MakeStatistics(self):
        return BasicStatistics()

    def MakeCalculator(self, slot: BasicSlot, stats: BasicStatistics):
        return BasicStatsCalculator(slot, stats)

    def MakePARSheet(self, slot: BasicSlot, stats: BasicStatistics, calculator: BasicStatsCalculator):
        return BasicPARSheet(self.GetBook(), self._par_formats, slot, stats, calculator)

    def MakePARFormats(self):
        return BasicPARFormats(self._book)

    def GetBook(self):
        return self._book

    def _CloseBook(self):
        #self._sheet.autofit()
        while True:
            try:
                self._book.close()
            except xl.exceptions.FileCreateError as e:
                decision = input("Exception caught in workbook.close(): %s\n"
                                 "Please close the file if it is open in Excel.\n"
                                 "Try to write file again? [Y/n]: " % e)
                if decision != 'n':
                    continue
            break

if __name__ == "__main__":
    book = BasicBook("NBN")
    book.WriteExel()
