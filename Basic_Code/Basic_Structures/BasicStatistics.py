from Basic_Code.Utils.BasicTimer import timer
import numpy as np
import pandas as pd
import json
from collections import defaultdict

class BasicStatistics:
    """
    Methods:
        ReadStatistics()
        ReadSpinWin()
        ReadSymbolViewCounter()
        ReadWinlineStats()
        ReadReelsetStats()
    """
    def __init__(self):
        self._total_spin_count = -1

        self._total_spin_win = {}  # key: win, value: counter
        self._section_feature_spin_win = dict()
        self._all_section_spin_spin_win = dict()
        self._all_spin_win_by_base_reelsets = dict()
        self._reelsets_spin_win = dict()  # 0 - section_index, 1 - reelset_index, 2 - win_amount, val - counter

        self._symbol_view_counter = []  # List of DataFrames()  0 - section_index, 1 - stack_height, 2 - symbol_id, 3 - column index, val - counter
        self._winline_stats = []  # List of DataFrames()  # 0 - section_index, 1 - symbol_id, 2 - winline_length, 3 - multiplier, val - counter
        self._winlines_counter = []  # List of numbers

        self._view_counter_columns = ["section_index", "stack_height", "symbol_id", "reel_index", "counter"]
        self._winlines_columns = ["section_index", "symbol_id", "winline_len", "multiplier", "counter"]

        self._section_all_spins_count = dict()
        self._section_feature_count = dict()

    @timer
    def ReadStatistics(self, data: dict):
        self._total_spin_count = data["Total Spin Count"]
        self._winlines_counter = np.array(data["Winlines counter"])
        self._ReadSpinWin(data["Spin Win"])
        self._ReadSpinWinBySection(data["Section Spin Win"])  # feature spin win
        self._ReadAllWinByBaseReelsets(data["All Wins By Base Reelsets"])
        self._ReadReelsetStats(data["Reelset Stats"])
        self._ReadWinlineStats(data["Winline Stats"])
        self._ReadSymbolViewCounter(data["Symbol View Counter"])


    def _ReadSpinWin(self, spin_win: list):
        for item in spin_win:
            self._total_spin_win[item["win"]] = item["count"]

    def _ReadSpinWinBySection(self, section_spin_win: dict):
        for index, (str_section_index, section_spin_win) in enumerate(section_spin_win.items()):
            int_section_index = int(str_section_index)
            assert int_section_index == index
            current_section_spin_win = dict()
            current_section_feature_count = 0
            for item in section_spin_win:
                current_section_spin_win[item["win"]] = item["count"]
                current_section_feature_count += item["count"]
            self._section_feature_count[int_section_index] = current_section_feature_count
            self._section_feature_spin_win[int_section_index] = current_section_spin_win

    def _ReadAllWinByBaseReelsets(self, win_by_base_reels: dict):
        for index, (str_reelset_index, spin_win) in enumerate(win_by_base_reels.items()):
            int_reelset_index = int(str_reelset_index)
            #assert int_reelset_index == index, str(int_reelset_index) + " " + str(index)
            current_reelset_spin_win = dict()
            for item in spin_win:
                current_reelset_spin_win[item["win"]] = item["count"]
            self._all_spin_win_by_base_reelsets[int_reelset_index] = current_reelset_spin_win


    def _ReadSymbolViewCounter(self, symbol_view_counter: dict):
        self._symbol_view_counter = []
        for str_section_index, x in symbol_view_counter.items():
            int_section_index = int(str_section_index)
            rows_counter = 0
            section_df = pd.DataFrame(columns=self._view_counter_columns)
            for str_stack_height, y in x.items():
                int_stack_height = int(str_stack_height)
                for str_symbol_id, list_counter in y.items():
                    int_symbol_id = int(str_symbol_id)
                    for reel_index, counter in enumerate(list_counter):
                        # ["section_index", "stack_height", "symbol_id", "column_index", "counter"]
                        temp_list = []
                        temp_list += [int_section_index]
                        temp_list += [int_stack_height]
                        temp_list += [int_symbol_id]
                        temp_list += [reel_index]
                        temp_list += [counter]
                        section_df.loc[rows_counter] = np.array(temp_list)
                        rows_counter += 1
            self._symbol_view_counter.append(section_df)

    def _ReadWinlineStats(self, winline_stats: dict):
        self._winline_stats = []
        for str_section_index, x in winline_stats.items():
            int_section_index = int(str_section_index)
            row_counter = 0
            section_winline_df = pd.DataFrame(columns=self._winlines_columns)
            for str_symbol_id, y in x.items():
                int_symbol_id = int(str_symbol_id)
                for str_winline_len, z in y.items():
                    int_winline_len = int(str_winline_len)
                    for str_mult, counter in z.items():
                        int_mult = int(str_mult)
                        #["section_index", "symbol_id", "winline_len", "multiplier", "counter"]
                        temp_row = []
                        temp_row += [int_section_index]
                        temp_row += [int_symbol_id]
                        temp_row += [int_winline_len]
                        temp_row += [int_mult]
                        temp_row += [counter]
                        section_winline_df.loc[row_counter] = temp_row
                        row_counter += 1
            self._winline_stats.append(section_winline_df.sort_values(by=["symbol_id"], axis=0).reset_index(drop=True))

    def _ReadReelsetStats(self, reelset_stats: dict):
        for i, (str_section_index, x) in enumerate(reelset_stats.items()):
            current_section_all_spins_count = 0
            int_section_index = int(str_section_index)
            assert i == int_section_index, "Wrong section order"
            current_reelset_index_to_win_amount = {}
            current_section_spin_win = defaultdict(int)
            for str_reelset_index, y in x.items():
                int_reelset_index = int(str_reelset_index)
                current_win_to_counter = {}
                for item in y:
                    current_section_spin_win[item["win"]] += item["count"]
                    current_win_to_counter[item["win"]] = item["count"]
                    current_section_all_spins_count += item["count"]
                current_reelset_index_to_win_amount[int_reelset_index] = current_win_to_counter
            self._all_section_spin_spin_win[int_section_index] = current_section_spin_win
            self._section_all_spins_count[int_section_index] = current_section_all_spins_count
            self._reelsets_spin_win[int_section_index] = current_reelset_index_to_win_amount

    def _ReadBasicMap(self, data: list, key_name: str, val_name: str):
        res = dict()
        for value_index, value in enumerate(data):
            res[value[key_name]] = value[val_name]
        return res

    def _ReadBasicVector(self, data: list, index_name: str, val_name: str):
        res = [0 for _ in range(len(data))]
        for val_index, value in enumerate(data):
            res[value[index_name]] = value[val_name]
        return np.array(res, dtype='uint64')

    def GetTotalSpinWin(self):
        return self._total_spin_win

    def GetTotalSpinCount(self):
        return self._total_spin_count

    def GetSymbolViewCounter(self, section_index: int = -1):
        if section_index == -1:
            return self._symbol_view_counter
        else:
            return self._symbol_view_counter[section_index]

    def GetWinlinesStat(self, section_index: int = -1):
        if section_index == -1:
            return self._winline_stats
        else:
            return self._winline_stats[section_index]

    def GetWinlineCounter(self):
        return self._winlines_counter

    def GetReelsetStats(self, section_index: int = -1):
        if section_index == -1:
            return self._reelsets_spin_win
        else:
            return self._reelsets_spin_win[section_index]

    def GetSectionFeaturesSpinWin(self, section_index: int = -1):
        if section_index == -1:
            return self._section_feature_spin_win
        else:
            return self._section_feature_spin_win[section_index]

    def GetSectionAllSpinWin(self, section_index: int = -1):
        if section_index == -1:
            return self._all_section_spin_spin_win
        else:
            return self._all_section_spin_spin_win[section_index]

    def GetSectionAllSpinsCount(self, section: int = -1):
        if section == -1:
            return self._section_all_spins_count
        else:
            return self._section_all_spins_count[section]

    def GetSectionFeatureCount(self, section: int = -1):
        if section == -1:
            return self._section_feature_count
        else:
            return self._section_feature_count[section]

    def GetAllWinByBaseReelsets(self, reelset_index: int = -1):
        if reelset_index == -1:
            return self._all_spin_win_by_base_reelsets
        else:
            return self._all_spin_win_by_base_reelsets[reelset_index]


if __name__ == "__main__":
    file = open(
        r"C:\Users\VitalijSaiganov\PycharmProjects\SlotHandler\Game_Source_Data\NBN\86\NBN_usual_86_no_gamble_stats.json")
    data = json.load(file)

    f = BasicStatistics()
    f.ReadStatistics(data)
    pd.set_option('display.max_rows', None)
    pd.set_option('display.max_columns', None)
    pd.set_option('display.width', None)
    pd.set_option('display.max_colwidth', 100)

    base_df = f.GetSectionAllSpinWin()
    print(base_df)


