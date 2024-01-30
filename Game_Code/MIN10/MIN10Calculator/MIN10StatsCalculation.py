from Basic_Code.Basic_Calculator.BasicStatsCalculator import BasicStatsCalculator
from Game_Code.MIN10.MIN10Structures.MIN10Slot import MIN10Slot
from Game_Code.MIN10.MIN10Structures.MIN10Statistics import MIN10Statistics
from Basic_Code.Utils.BasicTimer import timer
import json as j
import pandas as pd

class MIN10StatsCalculator(BasicStatsCalculator):
    def __init__(self, slot: MIN10Slot, stats: MIN10Statistics):
        super().__init__(slot, stats)

        self._start_symbol_counter = pd.DataFrame()
        self._number_winlines_counter = pd.DataFrame()

        self._start_symbol_columns = ["counter", "pulls_to_hit"]
        self._number_of_active_winlines_columns = ["counter", "pulls_to_hit"]

    @timer
    def CalcStats(self):
        self._HandleTotalSpinWin(self._bet)
        self._HandleSectionsSpinWins(self._section_all_spins_win, self._bet)
        self._CalculateReelsetsDF()
        self._CalculateSectionAllSpinsDF()
        self._CalculateSectionFeatureSpinsDF()
        self._CalculateAllWinReelsetsDF()
        self._CalculateTotalSpinWinDF()
        self._CalcConfidentIntervals()
        self._CalculateLineWins_By_Section()
        self._CalculateTopAward()
        # self._CalcRTPAllReelsets()
        # self._CalcRTPBaseReelsets()

        self._CalcWinlineStartWithSymbol()
        self._CalcNumberOfActiveWinlines()
        # self._CalcWinlineLen()
        self._CalcWinlineLenCounter()
        self._CalcNumberOfSymbolsOnBoard()
        self._CalcSymbolByReels()

    def _CalcWinlineStartWithSymbol(self):
        temp = pd.Series(stats.GetWinlineStartSymbolCounter())
        self._start_symbol_counter["counter"] = temp
        self._start_symbol_counter["pulls_to_hit"] = (stats.GetTotalSpinCount() / self._start_symbol_counter["counter"]).map("{:,.2f}".format)
        print(self._start_symbol_counter)

    def _CalcNumberOfActiveWinlines(self):
        temp = pd.Series(stats.GetWinlinesCounter())
        self._number_winlines_counter["counter"] = temp
        self._number_winlines_counter["pulls_to_hit"] = (stats.GetTotalSpinCount() / self._number_winlines_counter["counter"]).map("{:,.2f}".format)
        print(self._number_winlines_counter)

    def _CalcWinlineLen(self):
        temp = pd.Series(stats.GetWinlinesStat(0))
        self._number_winlines_counter["counter"] = temp
        self._number_winlines_counter["pulls_to_hit"] = (
                    stats.GetTotalSpinCount() / self._number_winlines_counter["counter"]).map("{:,.2f}".format)
        print(self._number_winlines_counter)

    def _CalcWinlineLenCounter(self):
        columns = ["counter", "pulls_to_hit"]

        winline_stats = stats.GetWinlinesStat(0)
        two_counter = winline_stats[winline_stats["winline_len"] == 2]["counter"].sum()
        three_counter = winline_stats[winline_stats["winline_len"] == 3]["counter"].sum()
        temp = pd.Series([two_counter, three_counter], index=[2, 3])
        temp_df = pd.DataFrame()
        temp_df["counter"] = temp
        temp_df["pulls_to_hit"] = (stats.GetTotalSpinCount() / temp_df["counter"]).map("{:,.2f}".format)
        print(temp_df)

    def _CalcNumberOfSymbolsOnBoard(self):
        temp = pd.Series(stats.GetNumberOfSymbolsInBoard())
        df = pd.DataFrame()
        df["counter"] = temp
        df["pulls_to_hit"] = (stats.GetTotalSpinCount() / df["counter"]).map("{:,.2f}".format)
        print(df)

    def _CalcSymbolByReels(self):
        temp = pd.Series(stats.GetSymbolCounterByReels())
        df = pd.DataFrame()
        df["counter"] = temp
        df["pulls_to_hit"] = (stats.GetTotalSpinCount() / df["counter"]).map("{:,.2f}".format)
        print(df)






if __name__ == "__main__":
    file = open(
        r"C:\Users\VitalijSaiganov\PycharmProjects\SlotHandler\Game_Source_Data\MIN10\93\MIN10_usual_93_no_gamble_stats.json")
    data = j.load(file)

    slot = MIN10Slot()
    stats = MIN10Statistics()

    slot.ReadSlot_Json(data)
    stats.ReadStatistics(data)

    calc = MIN10StatsCalculator(slot, stats)
    calc.CalcStats()

    pd.set_option('display.max_rows', None)
    pd.set_option('display.max_columns', None)
    pd.set_option('display.width', None)
    pd.set_option('display.max_colwidth', 100)

    print()
    print()
    print("RTP: ", calc.GetRTP())
    print("Hit fr: ", calc.GetHitFrPerc(), " 1 in ", 100 / calc.GetHitFrPerc())
    print("Avg mult: ", calc.GetAvgWinInCents() / slot.GetBet())
    winl = calc.GetSectionAllSpinWinsDF()[0]

    winl["pulls_to_hit"] = (stats.GetTotalSpinCount() / winl["total_counter"]).map("{:,.2f}".format)
    winl["rtp"] = winl["total_win"] / slot.GetBet() / stats.GetTotalSpinCount() * 100
    winl["avg_win"] = winl["total_win"] / winl["total_counter"]
    print(winl)




    print("Max win: x", calc.GetMaxWinInCents() / slot.GetBet(), "( 1 in ", stats.GetTotalSpinCount() / calc.GetMaxWinCount())



    print("Winline start with symbol: ")