from Basic_Code.Basic_Calculator.BasicSpinWin import BasicSpinWin
from Basic_Code.Utils.BasicTimer import timer
import json as j
import numpy as np
import pandas as pd
from scipy.stats import norm
import typing as t

from Basic_Code.Basic_Structures.BasicSlot import BasicSlot
from Basic_Code.Basic_Structures.BasicStatistics import BasicStatistics


class BasicStatsCalculator:
    def __init__(self, slot: BasicSlot, stats: BasicStatistics):
        self._total_spin_count = stats.GetTotalSpinCount()
        self._bet = slot.GetBet()

        self._total_spin_win_struct = BasicSpinWin
        self._sections_spin_win_structs = []

        self._winline_stats = stats.GetWinlinesStat()
        self._paytable = slot.GetPaytable()

        self._spin_win = stats.GetTotalSpinWin()
        self._reelsets_spin_win = stats.GetReelsetStats()
        self._section_feature_spin_win = stats.GetSectionFeaturesSpinWin()
        self._section_feature_count = stats.GetSectionFeatureCount()
        self._section_all_spins_win = stats.GetSectionAllSpinWin()
        self._section_all_spins_count = stats.GetSectionAllSpinsCount()
        self._all_spin_win_by_base_reelsets = stats.GetAllWinByBaseReelsets()

        self._hitFreq_percent = 0  # Calculated in __HandleSpinWin()
        self._rtp = 0  # Calculated in __HandleSpinWin()
        self._max_win_in_cents = 0  # Calculated in __HandleSpinWin()
        self._max_win_counter = 0  # Calculated in __HandleSpinWin()
        self._global_std_in_bets = 0  # Calculated in __HandleSpinWin()
        self._global_volatility = ''
        self._avg_win_in_cents_with_zero_win = 0  # Calculated in __HandleSpinWin()
        self._avg_win_in_cents_without_zero_win = 0  # Calculated in __HandleSpinWin()

        self._reelsets = slot.GetReelsets()
        self._section_names = slot.GetSectionNames()
        self._total_spin_win_df = pd.DataFrame()
        self._sections_all_spins_spin_win_dfs = []  # List of df with spin win data about section
        self._section_feature_spin_win_dfs = []
        self._reelsets_spin_win_dfs = []  # List of df with spin win data about reelsets
        self._all_win_by_base_reelsets_dfs = []  # One data frame for one reelset

        self._df_stats = ["total_win", "total_counter", "max_win"]
        self._win_ranges_names_detailed = ["[0, 0]",
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
        self._win_ranges_names = ["[0x, 0x]",
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
                                  "[1000x, inf)"]
        self._win_ranges = [[0, 0],
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

        self._additional_winline_columns = ["probability", "pulls_to_hit", "paytable", "rtp", "rtp_from_common"]
        self._confidence_column_names = ["num_games", "left_border", "right_border"]

        self._confidences_percent = [0.95, 0.99]
        self._confident_num_games_mults = [1, 2, 3, 5, 8]
        self._confident_num_games = [1_000, 10_000, 100_000, 1_000_000, 10_000_000, 100_000_000, 1_000_000_000, 10_000_000_000, 100_000_000_000]
        self._num_games_for_confidence = []  # List of Number of games for which confident is calculated
        self._total_confidence = dict()  # Confidence DFs, key - confidence percent
        self._sections_confidence = dict()  # Section confidance (confidence, section_index)

        self._top_award_number_of_games = [5_000, 10_000, 50_000, 100_000, 500_000, 1_000_000, 5_000_000,
                                           10_000_000, 14_000_000, 20_000_000, 30_000_000, 50_000_000,
                                           100_000_000, 1_000_000_000, 10_000_000_000, 100_000_000_000]
        self._top_award_number_names = ["5 Thousand", "10 Thousand", "50 Thousand", "100 Thousand",
                                        "500 Thousand", "1 Million", "5 Million", "10 Million", "14 Million",
                                        "20 Million", "30 Million", "50 Million", "100 Million",
                                        "1 Billiard", "10 Billiard", "100 Billiard"]
        self._top_award_columns = ["number_of_games", "top_award_cents", "more_or_equal_wins_count", "pulls_to_hit"]
        self._top_award_df = pd.DataFrame()

        self._all_reelsets_section_dfs = []
        self._base_reelsets_df = pd.DataFrame

        self._reelsets_column_names = ['reelset_name', 'total_hits', 'win_hits', 'total_win', 'rtp', 'rtp_common', 'prob', 'hits_1_in', 'avg_win_zero', 'avg_win_no_zero', 'max_win', 'win_freq_1_in', 'std', 'global_std', 'reelset_weight']

    def _MakeSpinWin(self, spin_win: dict, bet, total_spin_count):
        return BasicSpinWin(spin_win, bet, total_spin_count)

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
        self._CalcRTPAllReelsets()
        self._CalcRTPBaseReelsets()

    def _CalcIntervalsForConfidents(self):
        stop = False
        for num_games in self._confident_num_games:
            for mult in self._confident_num_games_mults:
                games = num_games * mult
                if games <= self._total_spin_count:
                    self._num_games_for_confidence.append(games)
                else:
                    stop = True
                    break
            if stop:
                break

    def _ConvertSpinWinToArray(self, spin_win: dict):
        return np.array(list(spin_win.items()), dtype="uint64")

    def _CalcOneConfidentInterval(self, confidence: float, mean: float, scale: float, num_games: int):
        return list(norm.interval(alpha=confidence, loc=mean, scale=scale / np.sqrt(num_games)))

    def _CalcConfidentIntervals(self):
        self._CalcIntervalsForConfidents()
        for i, perc in enumerate(self._confidences_percent):
            cur_df = pd.DataFrame(columns=self._confidence_column_names)
            row_counter = 0
            for num_games in self._num_games_for_confidence:
                cur_df.loc[row_counter] = ['{:,}'.format(num_games)] + self._CalcOneConfidentInterval(perc, self._rtp, self._global_std_in_bets, num_games)
                row_counter += 1
            cur_df['left_border'].where(cur_df['left_border'] > 0, 0, inplace=True)
            self._total_confidence[perc] = cur_df

        for i, perc in enumerate(self._confidences_percent):
            for section_index in range(len(self._section_names)):
                cur_df = pd.DataFrame(columns=self._confidence_column_names)
                row_counter = 0
                for num_games in self._num_games_for_confidence:
                    cur_df.loc[row_counter] = ['{:,}'.format(num_games)] + \
                                              self._CalcOneConfidentInterval(perc,
                                                                             self._sections_spin_win_structs[section_index].GetRTP(),
                                                                             self._sections_spin_win_structs[section_index].GetSTD(),
                                                                             num_games)
                    row_counter += 1
                cur_df['left_border'].where(cur_df['left_border'] > 0, 0, inplace=True)
                self._sections_confidence[perc, section_index] = cur_df

    def _CalculateLineWins_By_Section(self):
        for section_df in self._winline_stats:
            section_index = section_df["section_index"][0]
            section_df["section_spin_count"] = [self._section_all_spins_count[section_index] for _ in range(section_df.shape[0])]
            paytable = []
            for df_index in range(section_df.shape[0]):
                section_id = section_df.iloc[df_index]["section_index"]
                symbol_id = section_df.iloc[df_index]["symbol_id"]
                winline_len = section_df.iloc[df_index]["winline_len"]
                mult = section_df.iloc[df_index]["multiplier"]
                paytable.append(mult * self._paytable.GetSymbolPay(symbol_id, winline_len))
            section_df["paytable"] = paytable
            section_df["probability"] = section_df["counter"] / section_df["section_spin_count"]
            section_df["pulls_to_hit"] = section_df["section_spin_count"] / section_df["counter"]
            section_df["rtp_perc"] = section_df["paytable"] * section_df["counter"] / self._total_spin_count / self._bet * 100
            tot_win = (section_df["paytable"] * section_df["counter"]).sum()
            section_df["rtp_from_common_perc"] = section_df["paytable"] * section_df["counter"] / tot_win * 100

    @staticmethod
    def CalculateCommonSpinWinDF(spin_win: dict, win_ranges: list, win_ranges_names: list, bet: int, total_spin_count: int):

        ranges_total_win = np.array([0 for _ in range(len(win_ranges))], dtype="uint64")
        ranges_total_counter = np.array([0 for _ in range(len(win_ranges))], dtype="uint64")
        ranges_max_win = np.array([0 for _ in range(len(win_ranges))], dtype="uint64")

        for win, counter in spin_win.items():
            current_win = win * counter
            for i, (left_border, right_border) in enumerate(win_ranges):
                if left_border <= win <= right_border:
                    ranges_total_win[i] += current_win
                    ranges_total_counter[i] += counter
                    if ranges_max_win[i] < win:
                        ranges_max_win[i] = win
                    break
        current_spin_win_df = pd.DataFrame(data=np.array([ranges_total_win,
                                                          ranges_total_counter,
                                                          ranges_max_win]).transpose(),
                                           index=win_ranges_names,
                                           columns=["total_win", "total_counter", "max_win"],
                                           dtype="int64")
        current_spin_win_df['rtp'] = current_spin_win_df['total_win'] / (bet * total_spin_count) * 100
        current_spin_win_df['rtp_big'] = current_spin_win_df['total_win'] / np.sum(current_spin_win_df['total_win']) * 100
        current_spin_win_df['avg_win'] = current_spin_win_df['total_win'] / current_spin_win_df['total_counter']
        current_spin_win_df['win_1_in'] = np.sum(current_spin_win_df['total_counter']) / current_spin_win_df['total_counter']
        current_spin_win_df['win_1_in_small'] = total_spin_count / current_spin_win_df['total_counter']
        current_spin_win_df['avg_win_with_zero'] = np.sum(current_spin_win_df['avg_win'] * current_spin_win_df['total_counter']) / np.sum(current_spin_win_df['total_counter'])
        current_spin_win_df['avg_win_no_zero'] = np.sum(
            current_spin_win_df['avg_win'] * current_spin_win_df['total_counter']) / np.sum(
            current_spin_win_df['total_counter'][current_spin_win_df['total_win'] > 0])
        current_spin_win_df['win_1_in_total_feature'] = np.sum(current_spin_win_df['total_counter']) / np.sum(current_spin_win_df['total_counter'][current_spin_win_df['total_win'] > 0])
        current_spin_win_df['win_1_in_total_base'] = total_spin_count / np.sum(current_spin_win_df['total_counter'][current_spin_win_df['total_win'] > 0])
        return current_spin_win_df

    def _CalculateReelsetsDF(self):
        self._reelsets_spin_win_dfs = [[] for _ in range(len(self._section_names))]

        for section_index, reelsets in self._reelsets_spin_win.items():
            self._reelsets_spin_win_dfs[section_index] = [0 for _ in range(len(reelsets))]
            for reelset_index, reelset in reelsets.items():
                reelset_ranges_total_win = np.array([0 for _ in range(len(self._win_ranges))], dtype="uint64")
                reelset_ranges_total_counter = np.array([0 for _ in range(len(self._win_ranges))], dtype="uint64")
                reelset_ranges_max_win = np.array([0 for _ in range(len(self._win_ranges))], dtype="uint64")

                for win, counter in reelset.items():
                    cur_win = win * counter
                    for i, (left_border, right_border) in enumerate(self._win_ranges):
                        if left_border <= win <= right_border:
                            reelset_ranges_total_win[i] += cur_win
                            reelset_ranges_total_counter[i] += counter
                            if win > reelset_ranges_max_win[i]:
                                reelset_ranges_max_win[i] = win
                            break
                current_reelset_df = pd.DataFrame(data=np.array([reelset_ranges_total_win,
                                                                 reelset_ranges_total_counter,
                                                                 reelset_ranges_max_win]).transpose(),
                                                  index=self._win_ranges_names,
                                                  columns=self._df_stats,
                                                  dtype="uint64")
                current_reelset_df['rtp'] = current_reelset_df['total_win'] / (self._bet * self._total_spin_count) * 100
                current_reelset_df['avg_win'] = current_reelset_df['total_win'] / current_reelset_df['total_counter']
                current_reelset_df['win_1_in'] = np.sum(current_reelset_df['total_counter']) / current_reelset_df['total_counter']
                self._reelsets_spin_win_dfs[section_index][reelset_index] = current_reelset_df

    def _CalculateAllWinReelsetsDF(self):
        self._all_win_by_base_reelsets_dfs = []

        for index, (reelset_index, spin_win) in enumerate(self._all_spin_win_by_base_reelsets.items()):
            reelset_ranges_total_win = np.array([0 for _ in range(len(self._win_ranges))], dtype="uint64")
            reelset_ranges_total_counter = np.array([0 for _ in range(len(self._win_ranges))], dtype="uint64")
            reelset_ranges_max_win = np.array([0 for _ in range(len(self._win_ranges))], dtype="uint64")

            for reelset_win, counter in self._all_spin_win_by_base_reelsets[reelset_index].items():
                cur_win = reelset_win * counter
                for i, (left_border, right_border) in enumerate(self._win_ranges):
                    if left_border <= reelset_win <= right_border:
                        reelset_ranges_total_win[i] += cur_win
                        reelset_ranges_total_counter[i] += counter
                        if reelset_win > reelset_ranges_max_win[i]:
                            reelset_ranges_max_win[i] = reelset_win
                        break
            current_section_df = pd.DataFrame(data=np.array([reelset_ranges_total_win,
                                                             reelset_ranges_total_counter,
                                                             reelset_ranges_max_win]).transpose(),
                                              index=self._win_ranges_names,
                                              columns=self._df_stats,
                                              dtype="int64")
            current_section_df['rtp'] = current_section_df['total_win'] / (self._bet * self._total_spin_count) * 100
            current_section_df['avg_win'] = current_section_df['total_win'] / current_section_df['total_counter']
            current_section_df['win_1_in'] = np.sum(current_section_df['total_counter']) / current_section_df['total_counter']
            self._all_win_by_base_reelsets_dfs.append(current_section_df)

    def _CalculateSectionAllSpinsDF(self):
        self._sections_all_spins_spin_win_dfs = []  # df for each section
        for section_index, section in enumerate(self._section_all_spins_win):
            section_ranges_total_win = np.array([0 for _ in range(len(self._win_ranges))], dtype="uint64")
            section_ranges_total_counter = np.array([0 for _ in range(len(self._win_ranges))], dtype="uint64")
            section_ranges_max_win = np.array([0 for _ in range(len(self._win_ranges))], dtype="uint64")

            for section_win, counter in self._section_all_spins_win[section].items():
                cur_win = section_win * counter
                for i, (left_border, right_border) in enumerate(self._win_ranges):
                    if left_border <= section_win <= right_border:
                        section_ranges_total_win[i] += cur_win
                        section_ranges_total_counter[i] += counter
                        if section_win > section_ranges_max_win[i]:
                            section_ranges_max_win[i] = section_win
                        break
            current_section_df = pd.DataFrame(data=np.array([section_ranges_total_win,
                                                             section_ranges_total_counter,
                                                             section_ranges_max_win]).transpose(),
                                              index=self._win_ranges_names,
                                              columns=self._df_stats,
                                              dtype="int64")
            current_section_df['rtp'] = current_section_df['total_win'] / (self._bet * self._total_spin_count) * 100
            current_section_df['avg_win'] = current_section_df['total_win'] / current_section_df['total_counter']
            current_section_df['win_1_in'] = np.sum(current_section_df['total_counter']) / current_section_df['total_counter']
            self._sections_all_spins_spin_win_dfs.append(current_section_df)

    def _CalculateSectionFeatureSpinsDF(self):
        self._section_feature_spin_win_dfs = []  # df for each section
        for section_index, section in enumerate(self._section_feature_spin_win):
            section_ranges_total_win = np.array([0 for _ in range(len(self._win_ranges))], dtype="uint64")
            section_ranges_total_counter = np.array([0 for _ in range(len(self._win_ranges))], dtype="uint64")
            section_ranges_max_win = np.array([0 for _ in range(len(self._win_ranges))], dtype="uint64")

            for section_win, counter in self._section_feature_spin_win[section].items():
                cur_win = section_win * counter
                for i, (left_border, right_border) in enumerate(self._win_ranges):
                    if left_border <= section_win <= right_border:
                        section_ranges_total_win[i] += cur_win
                        section_ranges_total_counter[i] += counter
                        if section_win > section_ranges_max_win[i]:
                            section_ranges_max_win[i] = section_win
                        break
            current_section_df = pd.DataFrame(data=np.array([section_ranges_total_win,
                                                             section_ranges_total_counter,
                                                             section_ranges_max_win]).transpose(),
                                              index=self._win_ranges_names,
                                              columns=self._df_stats,
                                              dtype="int64")
            current_section_df['rtp'] = current_section_df['total_win'] / (self._bet * self._total_spin_count) * 100
            current_section_df['avg_win'] = current_section_df['total_win'] / current_section_df['total_counter']
            current_section_df['win_1_in'] = np.sum(current_section_df['total_counter']) / current_section_df['total_counter']
            self._section_feature_spin_win_dfs.append(current_section_df)

    def _CalculateTotalSpinWinDF(self):
        self._total_spin_win_df = self.CalculateCommonSpinWinDF(self._spin_win,
                                                                self._win_ranges,
                                                                self._win_ranges_names,
                                                                self._bet,
                                                                self._total_spin_count)

    def _CalculateTopAward(self):
        for i, game_count in enumerate(self._top_award_number_of_games):
            if game_count > self._total_spin_count:
                self._top_award_number_names = self._top_award_number_names[:i]
                self._top_award_number_of_games = self._top_award_number_of_games[:i]

        top_award = np.array([0 for _ in range(len(self._top_award_number_of_games))], dtype="uint64")
        more_or_equal_wins_count = np.array([0 for _ in range(len(self._top_award_number_of_games))], dtype="uint64")
        this_win_happen_1_in = np.array([0 for _ in range(len(self._top_award_number_of_games))], dtype="float64")

        wins_remaining = self._total_spin_count
        prev_win = 0
        prev_win_or_more_count = 0
        prev_win_happen_1_in = self._total_spin_count
        found_top_award_win_index = 0
        max_win = 0
        max_win_counter = 0
        for win, count in self._spin_win.items():
            if found_top_award_win_index == len(self._top_award_number_of_games)-1:
                break
            if win > max_win:
                max_win = win
                max_win_counter = count
            this_or_more_win_count = count + (wins_remaining - count)
            wins_remaining -= count
            this_or_more_win_happen_1_in = self._total_spin_count / this_or_more_win_count
            for i, num_games in enumerate(self._top_award_number_of_games):
                if prev_win_happen_1_in <= num_games <= this_or_more_win_happen_1_in:
                    top_award[i] = prev_win
                    more_or_equal_wins_count[i] = prev_win_or_more_count
                    this_win_happen_1_in[i] = self._total_spin_count / prev_win_or_more_count
                    found_top_award_win_index = i
            prev_win_or_more_count = this_or_more_win_count
            prev_win_happen_1_in = this_or_more_win_happen_1_in
            prev_win = win
        if found_top_award_win_index < len(self._top_award_number_of_games)-1:
            for i in range(found_top_award_win_index+1, len(self._top_award_number_of_games)):
                top_award[i] = max_win
                more_or_equal_wins_count[i] = max_win_counter
                this_win_happen_1_in[i] = self._total_spin_count / max_win_counter

        self._top_award_df = pd.DataFrame(data=np.array([self._top_award_number_of_games,
                                                         top_award,
                                                         more_or_equal_wins_count,
                                                         this_win_happen_1_in]).transpose(),
                                          columns=self._top_award_columns,
                                          index=self._top_award_number_names)

    def _HandleTotalSpinWin(self, bet: int = 0):
        total_section_struct = self._MakeSpinWin(self._spin_win, self._bet, self._total_spin_count)
        total_section_struct.HahdleSpinWin()
        self._total_spin_win_struct = total_section_struct

        self._rtp = total_section_struct.GetRTP()
        self._max_win_in_cents = total_section_struct.GetMaxWinCents()
        self._max_win_counter = total_section_struct.GetMaxWinCount()
        self._avg_win_in_cents_with_zero_win = total_section_struct.GetAvgWinInCents(True)
        self._avg_win_in_cents_without_zero_win = total_section_struct.GetAvgWinInCents(False)
        self._hitFreq_percent = total_section_struct.GetWinFreq(globl=True, decimal=False)
        self._global_std_in_bets = total_section_struct.GetSTD()

        volats = {(0, 3): 'Low Profile',
                  (4, 7): 'Medium Profile',
                  (7, np.Infinity): 'High Profile'}

        for ranges, name in volats.items():
            if ranges[0] <= self._global_std_in_bets <= ranges[1]:
                self._global_volatility = name

    def _HandleSectionsSpinWins(self, sections_spin_win: t.Dict[int, t.Dict], bet: int):
        spin_wins_incorrect_order = list(sections_spin_win.values())
        keys = list(sections_spin_win.keys())
        spin_wins_in_correct_order = sorted(spin_wins_incorrect_order, key=lambda dct: keys[spin_wins_incorrect_order.index(dct)])

        for section_index, spin_win in enumerate(spin_wins_in_correct_order):
            cur_spin_struct = self._MakeSpinWin(spin_win, bet, self._total_spin_count)
            cur_spin_struct.HahdleSpinWin()
            self._sections_spin_win_structs.append(cur_spin_struct)

    def _CalcRTPAllReelsets(self):
        self._all_reelsets_section_dfs = [pd.DataFrame(columns=self._reelsets_column_names, index=range(len(self._reelsets.GetReelsets(i)))) for i in range(len(self._reelsets))]
        for section_index, reelsets_spin_wins in self._reelsets_spin_win.items():
            total_section_win = 0
            total_section_hit_counter = 0
            total_section_win_counter = 0
            cur_section_reelsets_df = self._all_reelsets_section_dfs[section_index]
            for i, (reelset_index, reelset_spin_win) in enumerate(reelsets_spin_wins.items()):
                total_reelset_win = 0
                total_reelset_hit_counter = 0
                total_reelset_win_counter = 0
                reelset_max_win = 0
                for win, count in reelset_spin_win.items():
                    total_reelset_win += win * count
                    total_section_win += win * count
                    total_reelset_hit_counter += count
                    total_section_hit_counter += count
                    if win != 0:
                        total_reelset_win_counter += count
                        total_section_win_counter += count
                    if win > reelset_max_win:
                        reelset_max_win = win

                avg_win = total_reelset_win / total_reelset_hit_counter
                dispersion_numerator = 0
                for win, count in reelset_spin_win.items():
                    dispersion_numerator += np.power(win - avg_win, 2) * count
                std = np.power(dispersion_numerator / (total_reelset_hit_counter - 1), .5) / self._bet

                cur_section_reelsets_df['reelset_name'].loc[reelset_index] = self._reelsets.GetReelsetName(section_index, reelset_index)
                cur_section_reelsets_df['total_hits'].loc[reelset_index] = total_reelset_hit_counter
                cur_section_reelsets_df['win_hits'].loc[reelset_index] = total_reelset_win_counter
                cur_section_reelsets_df['total_win'].loc[reelset_index] = total_reelset_win

                left_border, right_border = self._reelsets.GetReelsetRange(section_index, reelset_index)
                reelset_weight = right_border - left_border + 1
                cur_section_reelsets_df['prob'].loc[reelset_index] = reelset_weight / self._reelsets.GetTotalSectionWeight(section_index)
                cur_section_reelsets_df['max_win'].loc[reelset_index] = reelset_max_win
                cur_section_reelsets_df['std'].loc[reelset_index] = std
                reelset_range = self._reelsets.GetReelsetRange(section_index, reelset_index)
                cur_section_reelsets_df['reelset_weight'].loc[reelset_index] = reelset_range[1] - reelset_range[0] + 1
            dispersion_numerator_global = 0
            global_avg_win = cur_section_reelsets_df['total_win'].sum() / cur_section_reelsets_df['total_hits'].sum()
            for reelset_index, reelset_spin_win in reelsets_spin_wins.items():
                for win, counter in reelset_spin_win.items():
                    dispersion_numerator_global += np.power(win - global_avg_win, 2) * counter
            global_std = np.power(dispersion_numerator_global / (cur_section_reelsets_df['total_hits'].sum()-1), .5) / self._bet
            cur_section_reelsets_df['global_std'] = global_std

            cur_section_reelsets_df['rtp_common'] = cur_section_reelsets_df['total_win'] / cur_section_reelsets_df['total_hits'] / self._bet
            cur_section_reelsets_df['rtp'] = cur_section_reelsets_df['rtp_common'] * cur_section_reelsets_df['prob'] * (cur_section_reelsets_df['total_hits'].sum() / self._total_spin_count)
            cur_section_reelsets_df['hits_1_in'] = total_section_hit_counter / cur_section_reelsets_df['total_hits']
            cur_section_reelsets_df['avg_win_zero'] = cur_section_reelsets_df['total_win'] / cur_section_reelsets_df['total_hits']
            cur_section_reelsets_df['avg_win_no_zero'] = cur_section_reelsets_df['total_win'].div(cur_section_reelsets_df['win_hits']).replace(np.inf, 0)#cur_section_reelsets_df['total_win'] / cur_section_reelsets_df['win_hits']
            cur_section_reelsets_df['win_freq_1_in'] = cur_section_reelsets_df['total_hits'] / cur_section_reelsets_df['win_hits']


        #self._reelsets_column_names = ['reelset_name', 'total_hits', 'win_hits', 'total_win', 'rtp', 'rtp_common', 'prob', 'hits_1_in', 'avg_win_zero', 'avg_win_no_zero', 'max_win', 'win_freq_1_in', 'std', 'global_std']
        #name, RTP, RTP(common), probability, hits 1 in?, avg_win(cents), win_freq(1 in)

    def _CalcRTPBaseReelsets(self):
        cur_section_reelsets_df = pd.DataFrame(columns=self._reelsets_column_names,
                                               index=range(len(self._all_spin_win_by_base_reelsets)))
        total_win = 0
        total_hit_counter = 0
        total_win_counter = 0
        for i, (reelset_index, reelset_spin_win) in enumerate(self._all_spin_win_by_base_reelsets.items()):
            total_reelset_win = 0
            total_reelset_hit_counter = 0
            total_reelset_win_counter = 0
            reelset_max_win = 0
            for win, count in reelset_spin_win.items():
                total_reelset_win += win * count
                total_win += win * count
                total_reelset_hit_counter += count
                total_hit_counter += count
                if win != 0:
                    total_reelset_win_counter += count
                    total_win_counter += count
                if win > reelset_max_win:
                    reelset_max_win = win

            avg_win = total_reelset_win / total_reelset_hit_counter
            dispersion_numerator = 0
            for win, count in reelset_spin_win.items():
                dispersion_numerator += np.power(win - avg_win, 2) * count
            std = np.power(dispersion_numerator / (total_reelset_hit_counter - 1), .5) / self._bet

            cur_section_reelsets_df['reelset_name'].loc[reelset_index] = self._reelsets.GetReelsetName(0, reelset_index)
            cur_section_reelsets_df['total_hits'].loc[reelset_index] = total_reelset_hit_counter
            cur_section_reelsets_df['win_hits'].loc[reelset_index] = total_reelset_win_counter
            cur_section_reelsets_df['total_win'].loc[reelset_index] = total_reelset_win

            left_border, right_border = self._reelsets.GetReelsetRange(0, reelset_index)
            reelset_weight = right_border - left_border + 1
            cur_section_reelsets_df['prob'].loc[reelset_index] = reelset_weight / self._reelsets.GetTotalSectionWeight(0)
            cur_section_reelsets_df['max_win'].loc[reelset_index] = reelset_max_win
            cur_section_reelsets_df['std'].loc[reelset_index] = std
            reelset_range = self._reelsets.GetReelsetRange(0, reelset_index)
            cur_section_reelsets_df['reelset_weight'].loc[reelset_index] = reelset_range[1] - reelset_range[0] + 1
        dispersion_numerator_global = 0
        global_avg_win = cur_section_reelsets_df['total_win'].sum() / cur_section_reelsets_df['total_hits'].sum()
        for reelset_index, reelset_spin_win in self._all_spin_win_by_base_reelsets.items():
            for win, counter in reelset_spin_win.items():
                dispersion_numerator_global += np.power(win - global_avg_win, 2) * counter
        global_std = np.power(dispersion_numerator_global / (cur_section_reelsets_df['total_hits'].sum() - 1),
                              .5) / self._bet
        cur_section_reelsets_df['global_std'] = global_std
        cur_section_reelsets_df['rtp_common'] = cur_section_reelsets_df['total_win'] / cur_section_reelsets_df[
            'total_hits'] / self._bet
        cur_section_reelsets_df['rtp'] = cur_section_reelsets_df['rtp_common'] * cur_section_reelsets_df['prob'] * (cur_section_reelsets_df['total_hits'].sum() / self._total_spin_count)
        cur_section_reelsets_df['hits_1_in'] = total_hit_counter / cur_section_reelsets_df['total_hits']
        cur_section_reelsets_df['avg_win_zero'] = cur_section_reelsets_df['total_win'] / cur_section_reelsets_df[
            'total_hits']
        cur_section_reelsets_df['avg_win_no_zero'] = cur_section_reelsets_df['total_win'] / cur_section_reelsets_df[
            'win_hits']
        cur_section_reelsets_df['win_freq_1_in'] = cur_section_reelsets_df['total_hits'] / cur_section_reelsets_df[
            'win_hits']
        self._base_reelsets_df = cur_section_reelsets_df

    @staticmethod
    def CalcVectorVariant_DF(variants: np.array, join_ranges: list, total_game_spin_count: int):
        columns = ['count', 'pulls_to_hit_big', 'pulls_to_hit_small', 'percent_big', 'percent_small']
        index_main = np.arange(0, len(variants))
        index_merged = np.arange(0, len(join_ranges))
        join_ranges[-1][1] = len(variants)-1

        main_table = pd.DataFrame(columns=columns, index=index_main)
        merged_table = pd.DataFrame(columns=['left_var', 'right_var'] + columns, index=index_merged)

        main_table['count'] = variants
        main_table['pulls_to_hit_big'] = np.sum(main_table['count']) / main_table['count']
        main_table['pulls_to_hit_small'] = total_game_spin_count / main_table['count']
        main_table['percent_big'] = main_table['count'] / np.sum(main_table['count']) * 100
        main_table['percent_small'] = main_table['count'] / total_game_spin_count * 100
        main_table['avg_val_with_zero'] = np.sum(main_table.index * main_table['count']) / np.sum(main_table['count'])
        if 0 in main_table.index:
            main_table['avg_val_without_zero'] = np.sum(main_table.index * main_table['count']) / (np.sum(main_table['count'])-main_table['count'].loc[0])
        else:
            # same as with zero
            main_table['avg_val_without_zero'] = np.sum(main_table.index * main_table['count']) / (
                        np.sum(main_table['count']))
        main_table['any_variant_1_in_base_no_zero'] = total_game_spin_count / \
                                                      (np.sum(main_table['count']) - (
                                                          main_table['count'].loc[0] if 0 in main_table.index else 0))
        main_table['any_variant_1_in_base_with_zero'] = total_game_spin_count / np.sum(main_table['count'])
        main_table['any_variant_1_in_feature_no_zero'] = np.sum(main_table['count']) / \
                                                         (np.sum(main_table['count']) - (main_table['count'].loc[
                                                                                             0] if 0 in main_table.index else 0))
        main_table.drop(main_table[main_table['count'] == 0].index, inplace=True)

        drop_ind = []
        for i, (left_border, right_border) in enumerate(join_ranges):
            if left_border not in main_table.index:
                for new_left in range(left_border, right_border+1):
                    if new_left in main_table.index:
                        left_border = new_left
                        break
                if left_border not in main_table.index:
                    drop_ind.append(i)
                    continue
            if right_border not in main_table.index:
                for new_right in range(right_border, left_border-1, -1):
                    if new_right in main_table.index:
                        right_border = new_right
                        break
                if right_border not in main_table.index:
                    drop_ind.append(i)
                    continue
            merged_table['left_var'].iloc[i] = left_border
            merged_table['right_var'].iloc[i] = right_border
            merged_table['count'].iloc[i] = np.sum(main_table['count'].loc[left_border: right_border+1])
        merged_table.drop(drop_ind, inplace=True)
        merged_table['pulls_to_hit_big'] = np.sum(merged_table['count']) / merged_table['count']
        merged_table['pulls_to_hit_small'] = total_game_spin_count / merged_table['count']
        merged_table['percent_big'] = merged_table['count'] / np.sum(merged_table['count']) * 100
        merged_table['percent_small'] = merged_table['count'] / total_game_spin_count * 100
        merged_table.drop(merged_table[merged_table['count'] == 0].index, inplace=True)
        merged_table.reset_index(inplace=True)

        return main_table, merged_table

    @staticmethod
    def CalcMapVariant_DF(variants: dict, join_ranges: list, total_game_spin_count: int):
        columns = ['count', 'pulls_to_hit_big', 'pulls_to_hit_small', 'percent_big', 'percent_small']
        index_main = sorted(list(variants.keys()))
        index_merged = np.arange(0, len(join_ranges))
        join_ranges[-1][1] = index_main[-1]

        main_table = pd.DataFrame(columns=columns, index=index_main)
        merged_table = pd.DataFrame(columns=['left_var', 'right_var'] + columns, index=index_merged)

        main_table['count'] = [variants[key] for key in main_table.index]
        main_table['pulls_to_hit_big'] = np.sum(main_table['count']) / main_table['count']
        main_table['pulls_to_hit_small'] = total_game_spin_count / main_table['count']
        main_table['percent_big'] = main_table['count'] / np.sum(main_table['count']) * 100
        main_table['percent_small'] = main_table['count'] / total_game_spin_count * 100
        main_table['avg_val_with_zero'] = np.sum(main_table.index * main_table['count']) / np.sum(main_table['count'])
        if 0 not in main_table.index:
            # same as with zero
            main_table['avg_val_without_zero'] = np.sum(main_table.index * main_table['count']) / (np.sum(main_table['count']))
        else:
            main_table['avg_val_without_zero'] = np.sum(main_table.index * main_table['count']) / (
                np.sum(main_table['count']) - main_table['count'].loc[0])
        main_table['any_variant_1_in_base_no_zero'] = total_game_spin_count / \
                                                      (np.sum(main_table['count']) - (main_table['count'].loc[0] if 0 in main_table.index else 0))
        main_table['any_variant_1_in_base_with_zero'] = total_game_spin_count / np.sum(main_table['count'])
        main_table['any_variant_1_in_feature_no_zero'] = np.sum(main_table['count']) / \
        (np.sum(main_table['count']) - (main_table['count'].loc[0] if 0 in main_table.index else 0))
        main_table.drop(main_table[main_table['count'] == 0].index, inplace=True)

        drop_ind = []
        for i, (left_border, right_border) in enumerate(join_ranges):
            if left_border not in main_table.index:
                for new_left in range(left_border, right_border+1):
                    if new_left in main_table.index:
                        left_border = new_left
                        break
                if left_border not in main_table.index:
                    drop_ind.append(i)
                    continue
            if right_border not in main_table.index:
                for new_right in range(right_border, left_border-1, -1):
                    if new_right in main_table.index:
                        right_border = new_right
                        break
                if right_border not in main_table.index:
                    drop_ind.append(i)
                    continue
            merged_table['left_var'].iloc[i] = left_border
            merged_table['right_var'].iloc[i] = right_border
            join_count = 0
            for ind in range(left_border, right_border+1):
                if ind in main_table.index:
                    join_count += main_table['count'].loc[ind]
            merged_table['count'].iloc[i] = join_count
        merged_table.drop(drop_ind, inplace=True)
        for index in merged_table.index:
            if merged_table['count'].loc[index] != 0:
                merged_table['pulls_to_hit_big'].loc[index] = np.sum(merged_table['count']) / merged_table['count'].loc[index]
                merged_table['pulls_to_hit_small'].loc[index] = total_game_spin_count / merged_table['count'].loc[index]
            else:
                merged_table['pulls_to_hit_big'].loc[index] = np.inf
                merged_table['pulls_to_hit_small'].loc[index] = np.inf
        merged_table['percent_big'] = merged_table['count'] / np.sum(merged_table['count']) * 100
        merged_table['percent_small'] = merged_table['count'] / total_game_spin_count * 100
        merged_table.drop(merged_table[merged_table['count'] == 0].index, inplace=True)
        merged_table.reset_index(inplace=True)

        return main_table, merged_table

    # def _HandleSpinWin(self, spin_win: np.array, bet: int, spin_count: int = -1):
    #     total_counter = 0
    #     total_win = 0
    #     max_win, max_win_counter = spin_win[-1]
    #     for win, count in spin_win:
    #         total_counter += count
    #         total_win += win * count
    #     avg_win_in_cents_with_zero = total_win / total_counter
    #     avg_win_in_cents_no_zero = total_win / (total_counter - spin_win[0][1])
    #     hit_freq = (total_counter - spin_win[0][1]) / total_counter * 100
    #     rtp_decimal = total_win / bet / (spin_count if spin_count != -1 else total_counter)
    #
    #     dispersion_numerator = 0
    #     for win, count in spin_win:
    #         dispersion_numerator += np.power(win - avg_win_in_cents_with_zero, 2) * count
    #     std = np.power(dispersion_numerator / (total_counter - 1), .5) / bet
    #
    #     return rtp_decimal, max_win, max_win_counter, avg_win_in_cents_with_zero, avg_win_in_cents_no_zero, hit_freq, std

    def GetHitFrPerc(self):
        return self._hitFreq_percent

    def GetRTP(self):
        return self._rtp

    def GetMaxWinInCents(self):
        return self._max_win_in_cents

    def GetMaxWinCount(self):
        return self._max_win_counter

    def GetAvgWinInCents(self, zero_included=False):
        if zero_included:
            return self._avg_win_in_cents_with_zero_win
        else:
            return self._avg_win_in_cents_without_zero_win

    def GetStdInPercent(self):
        return self._global_std_in_bets

    def GetVolatility(self):
        return self._global_volatility

    def GetTotalSpinWinStruct(self):
        return self._total_spin_win_struct

    def GetSectionSpinWinStructs(self, section_index: int = -1) -> t.Union[t.List[BasicSpinWin], BasicSpinWin]:
        if section_index == -1:
            return self._sections_spin_win_structs
        return self._sections_spin_win_structs[section_index]

    def GetSectionAllSpinWinsDF(self, section_index: int = -1):
        if section_index == -1:
            return self._sections_all_spins_spin_win_dfs
        else:
            return self._sections_all_spins_spin_win_dfs[section_index]

    def GetSectionFeatureSpinWinsDF(self, section_index: int = -1):
        if section_index == -1:
            return self._section_feature_spin_win_dfs
        else:
            return self._section_feature_spin_win_dfs[section_index]

    def GetReelsetsSpinWinsDF(self, section_index: int = -1):
        if section_index == -1:
            return self._reelsets_spin_win_dfs
        else:
            return self._reelsets_spin_win_dfs[section_index]

    def GetTotalSpinWinDF(self):
        return self._total_spin_win_df

    def GetSectionAllSpinsSpinCount(self, section: int = -1):
        if section == -1:
            return self._section_all_spins_count
        else:
            return self._section_all_spins_count[section]

    def GetSectionFeaturesSpinCount(self, section: int = -1):
        if section == -1:
            return self._section_feature_count
        else:
            return self._section_feature_count[section]

    def GetTotalConfIntervalsDF(self, index: int = -1):
        if index == -1:
            return self._total_confidence
        else:
            return self._total_confidence[index]

    def GetSectionConfIntervalsDF(self, perc: float = -1, section_index: int = -1):
        if perc != -1 and section_index != -1:
            return self._sections_confidence[(perc, section_index)]
        return self._sections_confidence

    def GetConfidencePercents(self):
        return self._confidences_percent

    def GetLineWinsDF(self, section: int = -1):
        if section == -1:
            return self._winline_stats
        else:
            return self._winline_stats[section]

    def GetTopAwardDF(self):
        return self._top_award_df

    def GetAllWinByBaseReelsetDF(self, base_reelset_index: int = -1):
        if base_reelset_index == -1:
            return self._all_win_by_base_reelsets_dfs
        else:
            return self._all_win_by_base_reelsets_dfs[base_reelset_index]

    def GetAllReelsetsSectionDF(self, section_index: int = -1):
        if section_index != -1:
            return self._all_reelsets_section_dfs[section_index]
        return self._all_reelsets_section_dfs

    def GetBaseReelsetsDF(self):
        return self._base_reelsets_df



if __name__ == "__main__":
    file = open(r"C:\Users\VitalijSaiganov\PycharmProjects\SlotHandler\Game_Source_Data\NBN\86\NBN_usual_86_no_gamble_stats.json")
    data = j.load(file)

    slot = BasicSlot()
    slot.ReadSlot_Json(data)

    stat = BasicStatistics()
    stat.ReadStatistics(data)

    calculation = BasicStatsCalculator(slot, stat)
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

    re_1 = calculation.GetBaseReelsetsDF()
    print(re_1.iloc[:, 1:])



    #
    # for i, df in enumerate(re_1):
    #     print("SECTION: ", i)
    #     print(df)

    #print(calculation.GetSectionSpinCount())

