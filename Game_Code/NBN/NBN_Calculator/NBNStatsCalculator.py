from Basic_Code.Basic_Calculator.BasicStatsCalculator import BasicStatsCalculator

import json as j
import numpy as np
import pandas as pd

from Game_Code.NBN.NBN_Structures.NBNSlot import NBNSlot
from Game_Code.NBN.NBN_Structures.NBNStatistics import NBNStatistics


class NBNStatsCalculator(BasicStatsCalculator):
    def __init__(self, slot: NBNSlot, stats: NBNStatistics):
        super().__init__(slot, stats)

        self._one_regular_free_spin_win = stats.GetOne_Regular_FreeSpinWin()
        self._one_special_free_spin_win = stats.GetOne_Special_FreeSpinWin()

        self._scatter_counter_base = stats.GetScatterCountBase()
        self._scatter_counter_base_respin = stats.GetScatterCountBaseRespin()

        self._number_of_wilds_base = stats.GetWildCountBase()
        self._number_of_wilds_free = stats.GetWildCountFree()
        self._number_of_wilds_special = stats.GetWildCountSpecial()

        self._wild_positions_counter_base = stats.GetWildPositionsCounterBase()
        self._wild_positions_counter_free = stats.GetWildPositionsCounterFree()
        self._wild_positions_counter_special = stats.GetWildPositionsCounterSpecial()

        self._base_win_with_wilds = stats.GetBaseWinWithWilds()
        self._base_win_no_wilds = stats.GetBaseWinNoWilds()

        self._one_regular_free_spin_win_df = pd.DataFrame
        self._one_special_free_spin_win_df = pd.DataFrame

        self._scatter_counter_base_main_df = pd.DataFrame
        self._scatter_counter_base_merged_df = pd.DataFrame
        self._scatter_counter_base_respin_main_df = pd.DataFrame
        self._scatter_counter_base_respin_merged_df = pd.DataFrame

        self._number_of_wilds_base_main_df = pd.DataFrame
        self._number_of_wilds_base_merged_df = pd.DataFrame
        self._number_of_wilds_free_main_df = pd.DataFrame
        self._number_of_wilds_free_merged_df = pd.DataFrame
        self._number_of_wilds_special_main_df = pd.DataFrame
        self._number_of_wilds_special_merged_df = pd.DataFrame

        self._wild_positions_counter_base_main_df = pd.DataFrame
        self._wild_positions_counter_base_merged_df = pd.DataFrame
        self._wild_positions_counter_free_main_df = pd.DataFrame
        self._wild_positions_counter_free_merged_df = pd.DataFrame
        self._wild_positions_counter_special_main_df = pd.DataFrame
        self._wild_positions_counter_special_merged_df = pd.DataFrame

        self._base_win_with_wild_df = pd.DataFrame
        self._base_win_no_wild_df = pd.DataFrame

    def CalcStats(self):
        super().CalcStats()

        self._CalcOneFreeSpinWin_Distribution_DF()
        self._CalcScatterCounterDistribution_DF()
        self._CalcNumberOfWilds_DF()
        self._CalcWildPositions_DF()
        self._CalcBaseWinDistribution_DF()

    def _CalcOneFreeSpinWin_Distribution_DF(self):
        self._one_regular_free_spin_win_df = BasicStatsCalculator.CalculateCommonSpinWinDF(self._one_regular_free_spin_win,
                                                                                           self._win_ranges,
                                                                                           self._win_ranges_names,
                                                                                           self._bet,
                                                                                           self._total_spin_count)
        self._one_special_free_spin_win_df = BasicStatsCalculator.CalculateCommonSpinWinDF(
            self._one_special_free_spin_win,
            self._win_ranges,
            self._win_ranges_names,
            self._bet,
            self._total_spin_count)

    def _CalcScatterCounterDistribution_DF(self):
        self._scatter_counter_base_main_df, self._scatter_counter_base_merged_df = BasicStatsCalculator.CalcVectorVariant_DF(self._scatter_counter_base,
                                                                                                                             [[0, 2], [3, 5]],
                                                                                                                             self._total_spin_count)
        self._scatter_counter_base_respin_main_df, self._scatter_counter_base_respin_merged_df = BasicStatsCalculator.CalcVectorVariant_DF(
            self._scatter_counter_base_respin,
            [[0, 2], [3, 5]],
            self._total_spin_count)

    def _CalcNumberOfWilds_DF(self):
        self._number_of_wilds_base_main_df, self._number_of_wilds_base_merged_df = BasicStatsCalculator.CalcVectorVariant_DF(self._number_of_wilds_base,
                                                                                                                                [[0, 0], [1, 2],
                                                                                                                                 [3, 5], [6, 8],
                                                                                                                                 [9, 12], [13, 15]],
                                                                                                                                self._total_spin_count)
        self._number_of_wilds_free_main_df, self._number_of_wilds_free_merged_df = BasicStatsCalculator.CalcVectorVariant_DF(
            self._number_of_wilds_free,
            [[0, 0], [1, 2],
             [3, 5], [6, 8],
             [9, 12], [13, 15]],
            self._total_spin_count)
        self._number_of_wilds_special_main_df, self._number_of_wilds_special_merged_df = BasicStatsCalculator.CalcVectorVariant_DF(
            self._number_of_wilds_special,
            [[7, 9]],
            self._total_spin_count)

    def _CalcWildPositions_DF(self):
        self._wild_positions_counter_base_main_df, self._wild_positions_counter_base_merged_df = BasicStatsCalculator.CalcVectorVariant_DF(self._wild_positions_counter_base,
                                                                                                                                           [[0,2], [3, 5],
                                                                                                                                            [6, 8], [9, 11],
                                                                                                                                            [12, 14]],
                                                                                                                                           self._total_spin_count)
        self._wild_positions_counter_free_main_df, self._wild_positions_counter_free_merged_df = BasicStatsCalculator.CalcVectorVariant_DF(
            self._wild_positions_counter_free,
            [[0, 2], [3, 5],
             [6, 8], [9, 11],
             [12, 14]],
            self._total_spin_count)
        self._wild_positions_counter_special_main_df, self._wild_positions_counter_special_merged_df = BasicStatsCalculator.CalcVectorVariant_DF(
            self._wild_positions_counter_special,
            [[0, 2], [3, 5],
             [6, 8], [9, 11],
             [12, 14]],
            self._total_spin_count)

    def _CalcBaseWinDistribution_DF(self):
        self._base_win_with_wild_df = BasicStatsCalculator.CalculateCommonSpinWinDF(self._base_win_with_wilds,
                                                                                    self._win_ranges,
                                                                                    self._win_ranges_names,
                                                                                    self._bet,
                                                                                    self._total_spin_count)
        self._base_win_no_wild_df = BasicStatsCalculator.CalculateCommonSpinWinDF(self._base_win_no_wilds,
                                                                                  self._win_ranges,
                                                                                  self._win_ranges_names,
                                                                                  self._bet,
                                                                                  self._total_spin_count)

    def GetOneFreeSpinRegularWin_DF(self):
        return self._one_regular_free_spin_win_df

    def GetOneFreeSpinSpecialWin_DF(self):
        return self._one_special_free_spin_win_df

    def GetScatterCountBase_DFs(self):
        return self._scatter_counter_base_main_df, self._scatter_counter_base_merged_df

    def GetScatterCountBaseRespin_DFs(self):
        return self._scatter_counter_base_respin_main_df, self._scatter_counter_base_respin_merged_df

    def GetNumberOfWildsBase_DFs(self):
        return self._number_of_wilds_base_main_df, self._number_of_wilds_base_merged_df

    def GetNumberOfWildsFree_DFs(self):
        return self._number_of_wilds_free_main_df, self._number_of_wilds_free_merged_df

    def GetNumberOfWildsSpecial_DFs(self):
        return self._number_of_wilds_special_main_df, self._number_of_wilds_special_merged_df

    def GetWildPositionsCounterBase_DFs(self):
        return self._wild_positions_counter_base_main_df, self._wild_positions_counter_base_merged_df

    def GetWildPositionsCounterFree_DFs(self):
        return self._wild_positions_counter_free_main_df, self._wild_positions_counter_free_merged_df

    def GetWildPositionsCounterSpecial_DFs(self):
        return self._wild_positions_counter_special_main_df, self._wild_positions_counter_special_merged_df

    def GetBaseWinWithWilds_DF(self):
        return self._base_win_with_wild_df

    def GetBaseWinNoWilds_DF(self):
        return self._base_win_no_wild_df

    def GetFreeScatterCountBase_DFs(self):
        return self._scatter_counter_base_main_df, self._scatter_counter_base_merged_df

    def GetFreeScatterCountBaseRespin_DFs(self):
        return self._scatter_counter_base_respin_main_df, self._scatter_counter_base_respin_merged_df

    def GetWildCountBase_DFs(self):
        return self._number_of_wilds_base_main_df, self._number_of_wilds_base_merged_df

    def GetWildCountFree_DFs(self):
        return self._number_of_wilds_free_main_df, self._number_of_wilds_free_merged_df

    def GetWildCountSpecial_DFs(self):
        return self._number_of_wilds_special_main_df, self._number_of_wilds_special_merged_df

    def GetWildPositionCountBase_DFs(self):
        return self._wild_positions_counter_base_main_df, self._wild_positions_counter_base_merged_df

    def GetWildPositionCountFree_DFs(self):
        return self._wild_positions_counter_free_main_df, self._wild_positions_counter_free_merged_df

    def GetWildPositionCountSpecial_DFs(self):
        return self._wild_positions_counter_special_main_df, self._wild_positions_counter_special_merged_df

    def GetBaseWinNoWild_DF(self):
        return self._base_win_no_wild_df

    def GetBaseWinWithWild_DF(self):
        return self._base_win_with_wild_df

#     Stats:
# 1. One regular free spin win + distribution ЭТО
# 2. One special free spin win + distribution ЭТО
# 3. Free spins total section win avg + distribution without Feature-Respin
# 4. Free spins total section win avg + distribution with Feature Respin
# 5. Scatter counter base (freq + trigger freq) ЭТО
# 6. Scatter counter base re-spin (freq + trigger freq) ЭТО
# 7. Number of wilds distributions + avg + freq base ЭТО
# 8. Number of wilds distributions + avg + freq free ЭТО
# 9. Number of wilds distributions + avg + freq special ЭТО
# 10. Wild position counter base ЭТО
# 11. Wild position counter free ЭТО
# 12. Wild position counter special ЭТО
# 13. Wild pattern counter base
# 14. Wild pattern counter free
# 15. Wild pattern counter special
# 16. Board full of wilds base
# 17. Board full of wilds free
# 18. scatters by reels base
# 19. scatters by reels base respin
# 20. Base win with wilds avg + distribution ЭТО
# 21. Base win without wilds avg + distrivution ЭТО

if __name__ == "__main__":
    file = open(r"C:\Users\VitalijSaiganov\PycharmProjects\SlotHandler\Game_Source_Data\NBN\86\NBN_usual_86_no_gamble_stats.json")
    data = j.load(file)

    slot = NBNSlot()
    slot.ReadSlot_Json(data)

    stat = NBNStatistics()
    stat.ReadStatistics(data)

    calculation = NBNStatsCalculator(slot, stat)

    pd.set_option('display.max_rows', None)
    pd.set_option('display.max_columns', None)
    pd.set_option('display.width', None)
    pd.set_option('display.max_colwidth', 100)

    calculation.CalcStats()

    row_count = 3
    col_count = 5

    def on_one_row(cell_index: int):
        for row in range(row_count):
            if row * col_count <= cell_index < (row + 1) * col_count:
                return row

    def on_one_col(cell_index: int):
        return cell_index % col_count

    r = calculation.GetWildPositionCountBase_DFs()[0]

    print(r.groupby(on_one_col).sum())

    print("RTP", calculation.GetRTP() * 100)
    print("STD", calculation.GetStdInPercent())

