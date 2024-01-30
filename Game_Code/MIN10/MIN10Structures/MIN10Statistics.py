from Basic_Code.Basic_Structures.BasicStatistics import BasicStatistics
from Basic_Code.Utils.BasicTimer import timer
import json as j
import numpy as np

class MIN10Statistics(BasicStatistics):
    def __init__(self):
        super().__init__()
        self.__number_winlines_counter = dict()  #  Number of winlines counter
        self.__winline_start_symbol = dict()  # Winline start symbol
        self.__number_of_symbols_in_board = dict()  #  Number of symbols on board
        self.__symbol_counter_by_reels = []  #  Symbol by reels

    @timer
    def ReadStatistics(self, data: dict):
        self._total_spin_count = data["Total Spin Count"]
        self._winlines_counter = np.array(data["Winlines counter"])
        self._ReadSpinWin(data["Spin Win"])
        # self._ReadSpinWinBySection(data["Section Spin Win"])  # feature spin win
        # self._ReadAllWinByBaseReelsets(data["All Wins By Base Reelsets"])
        self._ReadReelsetStats(data["Reelset Stats"])
        self._ReadWinlineStats(data["Winline Stats"])
        self._ReadSymbolViewCounter(data["Symbol View Counter"])

        self.__ReadNumberofWinlinesCounter(data["Number of winlines counter"])
        self.__ReadWinlineStartSymbol(data["Winline start symbol"])
        self.__ReadNumberofSymbolsOnBoard(data["Number of symbols on board"])
        self.__ReadSymbolCounterByReels(data["Symbol by reels"])

    def __ReadNumberofWinlinesCounter(self, data: dict):
        self.__number_winlines_counter = dict()
        for item in data:
            self.__number_winlines_counter[item["number"]] = item["counter"]

    def __ReadWinlineStartSymbol(self, data: dict):
        self.__winline_start_symbol = dict()
        for item in data:
            self.__winline_start_symbol[item["symbol"]] = item["counter"]

    def __ReadNumberofSymbolsOnBoard(self, data):
        self.__number_of_symbols_in_board = dict()
        for item in data:
            self.__number_of_symbols_in_board[item["number"]] = item["counter"]

    def __ReadSymbolCounterByReels(self, data):
        self.__symbol_counter_by_reels = dict()
        for item in data:
            self.__symbol_counter_by_reels[item["reel"]] = item["counter"]

    def GetWinlinesCounter(self):
        return self.__number_winlines_counter

    def GetWinlineStartSymbolCounter(self):
        return self.__winline_start_symbol

    def GetNumberOfSymbolsInBoard(self):
        return self.__number_of_symbols_in_board

    def GetSymbolCounterByReels(self):
        return self.__symbol_counter_by_reels

if __name__ == "__main__":
    file = open(
        r"/Game_Source_Data/Json Stats/MIN10/93/MIN10_usual_93_no_gamble_stats.json")
    data = j.load(file)

    r = MIN10Statistics()
    r.ReadStatistics(data)
    print(r.GetTotalSpinCount())
    print()

