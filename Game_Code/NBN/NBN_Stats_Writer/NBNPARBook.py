from Game_Code.NBN.NBN_Stats_Writer.NBNPARFormats import NBNPARFormats

from Game_Code.NBN.NBN_Structures.NBNSlot import NBNSlot
from Game_Code.NBN.NBN_Structures.NBNStatistics import NBNStatistics
from Game_Code.NBN.NBN_Calculator.NBNStatsCalculator import NBNStatsCalculator
from Game_Code.NBN.NBN_Stats_Writer.NBNPARFormats import NBNPARFormats
from Game_Code.NBN.NBN_Stats_Writer.NBNPARSheet import NBNPARSheet
from Basic_Code.Stats_Writer.BasicPARBook import BasicBook


class NBNBook(BasicBook):
    def __init__(self, game_name_short: str):
        super().__init__(game_name_short)
        self._file_name = "PAR_" + game_name_short + ".xlsx"
        self._save_book_path = self._path_handler.GetPARPath(game_name_short, self._file_name)

    def MakeSlot(self):
        return NBNSlot()

    def MakeStatistics(self):
        return NBNStatistics()

    def MakeCalculator(self, slot: NBNSlot, stats: NBNStatistics):
        return NBNStatsCalculator(slot, stats)

    def MakePARSheet(self, slot: NBNSlot, stats: NBNStatistics, calculator: NBNStatsCalculator):
        return NBNPARSheet(self.GetBook(), self._par_formats, slot, stats, calculator)

    def MakePARFormats(self):
        return NBNPARFormats(self.GetBook())

    def WriteExel(self):
        self._InitBook()

        self._WritePARSheet(self._game_name_short, 'usual_93', simulation_tags=['bet'])

        self._CloseBook()

if __name__ == "__main__":
    book = NBNBook("NBN")
    book.WriteExel()
