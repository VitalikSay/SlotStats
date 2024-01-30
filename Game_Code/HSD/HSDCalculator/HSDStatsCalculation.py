from Basic_Code.Basic_Calculator.BasicStatsCalculator import BasicStatsCalculator
from Basic_Code.Basic_Structures.BasicSlot import BasicSlot
from Basic_Code.Basic_Structures.BasicStatistics import BasicStatistics

from Game_Code.HSD.HSDStructures.HSDSlot import HSDSlot
from Game_Code.HSD.HSDStructures.HSDStatistics import HSDStatistics
from Basic_Code.Utils.BasicTimer import timer
import json as j
import pandas as pd
import numpy as np

class HSDStatsCalculator(BasicStatsCalculator):
    def __init__(self, slot: BasicSlot, stats: BasicStatistics):
        self._total_spin_count = stats.GetTotalSpinCount()
        self._bet = slot.GetBet()

        self.__winline_stats = stats.GetWinlinesStat()
        self.__paytable = slot.GetPaytable()

        self._spin_win = self._ConvertSpinWinToArray(stats.GetTotalSpinWin())
        self.__reelsets_spin_win = stats.GetReelsetStats()
        self.__section_feature_spin_win = stats.GetSectionFeaturesSpinWin()
        self.__section_feature_count = stats.GetSectionFeatureCount()
        self.__section_all_spins_win = stats.GetSectionAllSpinWin()
        self.__section_all_spins_count = stats.GetSectionAllSpinsCount()
        self.__all_spin_win_by_base_reelsets = stats.GetAllWinByBaseReelsets()

        self._hitFreq_percent = 0  # Calculated in __HandleSpinWin()
        self._rtp = 0  # Calculated in __HandleSpinWin()
        self._max_win_in_cents = 0  # Calculated in __HandleSpinWin()
        self._max_win_counter = 0  # Calculated in __HandleSpinWin()
        self._global_std_in_percent = 0  # Calculated in __HandleSpinWin()
        self._global_volatility = ''
        self._avg_win_in_cents_with_zero_win = 0  # Calculated in __HandleSpinWin()
        self._avg_win_in_cents_without_zero_win = 0  # Calculated in __HandleSpinWin()

        self.__section_names = slot.GetSectionNames()
        self.__total_spin_win_df = pd.DataFrame()
        self.__sections_all_spins_spin_win_dfs = []  # List of df with spin win data about section
        self.__section_feature_spin_win_dfs = []
        self.__reelsets_spin_win_dfs = []  # List of df with spin win data about reelsets
        self.__all_win_by_base_reelsets_dfs = []  # One data frame for one reelset

        self.__df_stats = ["total_win", "total_counter", "max_win"]
        self.__win_ranges_names_detailed = ["[0, 0]",
                                            "[1, 0.5 * (Bet)]",
                                            "[0.5*(Bet)+1, Bet-1]",
                                            "[Bet, Bet]",
                                            "[Bet+1, 2*(Bet)-1]",
                                            "[2*Bet,3*(Bet)-1]",
                                            "[3*Bet, 5*(Bet)-1]",
                                            "[5*Bet, 10*(Bet)-1]",
                                            "[10*Bet, 20*(Bet)-1]",
                                            "[20*Bet, 30*(Bet)-1]",
                                            "[30*Bet, 50*(Bet)-1]",
                                            "[50*Bet, 100*(Bet)-1]",
                                            "[100*Bet, 200*(Bet)-1]",
                                            "[200*Bet, 500*(Bet)-1]",
                                            "[500*Bet, 1000*(Bet)-1]",
                                             "[>= 1000*Bet]"]
        self.__win_ranges_names = ["[0x, 0x]",
                                   "(0x, 0.5x]",
                                   "(0.5x, 1x)",
                                   "[1x, 1x]",
                                   "(1x, 2x)",
                                   "[2x, 3x)",
                                   "[3x, 5x)",
                                   "[5x, 10x)",
                                   "[10x, 20x)",
                                   "[20x, 30x)",
                                   "[30x, 50x)",
                                   "[50x, 100x)",
                                   "[100x, 200x)",
                                   "[200x, 500x)",
                                   "[500x, 1000x)",
                                   "[1000x, inf]"]
        self.__win_ranges = [[0, 0],
                             [1, np.floor(.5 * self._bet)],
                             [np.floor(.5 * self._bet) + 1, self._bet - 1],
                             [self._bet, self._bet],
                             [self._bet + 1, (2 * self._bet) - 1],
                             [2 * self._bet, (3 * self._bet) - 1],
                             [3 * self._bet, (5 * self._bet) - 1],
                             [5 * self._bet, (10 * self._bet) - 1],
                             [10 * self._bet, (20 * self._bet) - 1],
                             [20 * self._bet, (30 * self._bet) - 1],
                             [30 * self._bet, (50 * self._bet) - 1],
                             [50 * self._bet, (100 * self._bet) - 1],
                             [100 * self._bet, (200 * self._bet) - 1],
                             [200 * self._bet, (500 * self._bet) - 1],
                             [500 * self._bet, (1000 * self._bet) - 1],
                             [1000 * self._bet, np.Infinity]]

        self.__additional_winline_columns = ["probability", "pulls_to_hit", "paytable", "rtp", "rtp_from_common"]
        self._confidence_column_names = ["num_games", "left_border", "right_border"]

        self._confidences_percent = [0.95, 0.99]
        self._confident_num_games_mults = [1, 2, 3, 5, 8]
        self._confident_num_games = [1_000, 10_000, 100_000, 1_000_000, 10_000_000, 100_000_000, 1_000_000_000, 10_000_000_000, 100_000_000_000]
        self._num_games_for_confidence = []  # List of Number of games for which confident is calculated
        self._confidence = []  # List of DataFrames (for each percent)

        self.__top_award_number_of_games = [100_000, 500_000, 1_000_000, 5_000_000, 10_000_000, 14_000_000,
                                            20_000_000, 30_000_000, 50_000_000, 100_000_000, 1_000_000_000,
                                            10_000_000_000]
        self.__top_award_number_names = ["100 thousand", "500 thousand", "1 million", "5 million",
                                         "10 million", "14 million", "20 million", "30 million",
                                         "50 million", "100 million", "1 billiard", "10 billiard"]
        self.__top_award_columns = ["number_of_games", "top_award_cents", "more_or_equal_wins_count", "pulls_to_hit"]
        self.__top_award_df = pd.DataFrame()


    def _CalcConfidentIntervals(self):
        self._CalcIntervalsForConfidents()
        self._confidence = [pd.DataFrame(columns=self._confidence_column_names) for _ in range(len(self._confidences_percent))]
        for i, perc in enumerate(self._confidences_percent):
            row_counter = 0
            for num_games in self._num_games_for_confidence:
                self._confidence[i].loc[row_counter] = ['{:,}'.format(num_games)] + self._CalcOneConfidentInterval(perc, 0.861737, 0.815809, num_games)
                row_counter += 1

    @timer
    def CalcStats(self):
        self._HandleTotalSpinWin(self._bet)
        # self.__CalculateReelsetsDF()
        # self.__CalculateSectionAllSpinsDF()
        # self.__CalculateSectionFeatureSpinsDF()
        # self.__CalculateAllWinReelsetsDF()
        # self.__CalculateTotalSpinWinDF()
        self._CalcConfidentIntervals()
        # self.__CalculateLineWins_By_Section()
        # self.__CalculateTopAward()


if __name__ == "__main__":
    file = open(
        r"C:\Users\VitalijSaiganov\PycharmProjects\SlotHandler\Game_Source_Data\HSD\93\HSD_usual_93_no_gamble_stats.json")
    data = j.load(file)
    print(data["Reelset Stats"]['0']['31'])

    tot_count = 0
    tot_win = 0
    for str_reelset_index, win_list in data["Reelset Stats"]['0'].items():
        for val in win_list:
            tot_count += val['count']
            tot_win += val['win'] * val['count']
    print(tot_count)
    print("Base RTP: ", '{:.4f}'.format(tot_win / (10 * tot_count)*100))

    _avg_win_in_cents_with_zero_win = tot_win / tot_count
    print("avg win: ", _avg_win_in_cents_with_zero_win)
    dispersion_numerator = 0
    for str_reelset_index, win_list in data["Reelset Stats"]['0'].items():
        for val in win_list:
            tot_count += val['count']
            tot_win += val['win'] * val['count']
            dispersion_numerator += np.power(val['win'] - _avg_win_in_cents_with_zero_win, 2) * val['count']

    print("STD: ", np.power(dispersion_numerator / (tot_count - 1), .5)/10)



    slot = HSDSlot()
    slot.ReadSlot_Json(data)

    stat = HSDStatistics()
    stat.ReadStatistics(data)

    calculation = HSDStatsCalculator(slot, stat)
    calculation.CalcStats()
    print("RTP", calculation.GetRTP() * 100)
    print("STD", calculation.GetStdInPercent())
    # hit_fr_perc = calculation.GetHitFrPerc()
    # print("Hit fr perc: ", hit_fr_perc, "(1 in", 100 / hit_fr_perc)
    # print("Max Win: ", calculation.GetMaxWinInCents())
    # print("Avg win ", calculation.GetAvgWinInCents())

    pd.set_option('display.max_rows', None)
    pd.set_option('display.max_columns', None)
    pd.set_option('display.width', None)
    pd.set_option('display.max_colwidth', 100)

    re_1 = calculation.GetTotalConfIntervalsDF(1)

    print(re_1)

    # print(calculation.GetSectionSpinCount())