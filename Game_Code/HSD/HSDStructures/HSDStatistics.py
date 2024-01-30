from Basic_Code.Basic_Structures.BasicStatistics import BasicStatistics
from Basic_Code.Utils.BasicTimer import timer
import json as j
import numpy as np

class HSDStatistics(BasicStatistics):
    @timer
    def ReadStatistics(self, data: dict):
        self._total_spin_count = data["Total Spin Count"]
        self.__winlines_counter = np.array(data["Winlines counter"])
        self._BasicStatistics__ReadSpinWin(data["Spin Win"])
        # self._BasicStatistics__ReadSpinWinBySection(data["Section Spin Win"])  # feature spin win
        # self._BasicStatistics__ReadAllWinByBaseReelsets(data["All Wins By Base Reelsets"])
        # self._BasicStatistics__ReadReelsetStats(data["Reelset Stats"])
        self._BasicStatistics__ReadWinlineStats(data["Winline Stats"])
        self._BasicStatistics__ReadSymbolViewCounter(data["Symbol View Counter"])

if __name__ == "__main__":
    file = open(
        r"C:\Users\VitalijSaiganov\PycharmProjects\SlotHandler\Game_Source_Data\HSD\90\HSD_usual_90_no_gamble_stats.json")
    data = j.load(file)

    stat = HSDStatistics()
    stat.ReadStatistics(data)