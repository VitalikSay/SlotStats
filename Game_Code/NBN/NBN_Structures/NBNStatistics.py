from Basic_Code.Basic_Structures.BasicStatistics import BasicStatistics
from Basic_Code.Basic_Structures.Features.SLStatistics import SLStatistics

import numpy as np
import pandas as pd
import json
from collections import defaultdict


class NBNStatistics(SLStatistics):
    def __init__(self):
        super().__init__()

        self._one_free_spin_regular_win = dict()
        self._one_free_spin_special_win = dict()

        self._free_spins_total_section_win_without_feature_respin = dict()
        self._free_spins_total_section_win_with_feature_respin = dict()

        self._scatter_counter_base = np.array  # index - scatter count in board; value - situations counter
        self._scatter_counter_base_respin = np.array  # index - scatter count in board; value - situations counter

        self._wild_counter_base = np.array
        self._wild_positions_counter_base = np.array
        self._wild_pattern_counter_base = np.array

        self._board_full_of_wilds_counter_base = -1
        self._board_full_of_wilds_counter_free = -1
        self._spectacular_link_counter_from_base = -1

        self._base_and_base_respin_win_with_wilds = dict()  # only wins when wilds fell
        self._base_win_no_wilds = dict()  # only wins when wilds didn't fall

    def ReadStatistics(self, data: dict):
        super().ReadStatistics(data)
        # self._one_free_spin_regular_win = super()._ReadBasicMap(data['One Free Spin regular spin win'],
        #                                                         key_name='win',
        #                                                         val_name='counter')
        # self._one_free_spin_special_win = super()._ReadBasicMap(data['One Free Spin special spin win'],
        #                                                         key_name='win',
        #                                                         val_name='counter')

        self._scatter_counter_base = super()._ReadBasicVector(data['Scatter Counter Base'],
                                                              index_name='number_of_scatters',
                                                              val_name='counter')
        self._scatter_counter_base_respin = super()._ReadBasicVector(data['Scatter Counter Base Re-Spin'],
                                                                     index_name='number_of_scatters',
                                                                     val_name='counter')
        self._wild_counter_base = super()._ReadBasicVector(data['Wild Counter Base'],
                                                           index_name='number_of_wilds',
                                                           val_name='counter')
        self._wild_positions_counter_base = super()._ReadBasicVector(data['Wild positions counter Base'],
                                                                     index_name='wild_position_index',
                                                                     val_name='counter')
        self._wild_pattern_counter_base = super()._ReadBasicVector(data['Wild Pattern Counter Base'],
                                                                   index_name='pattern_index',
                                                                   val_name='counter')

        self._board_full_of_wilds_counter_base = data['Board Full of Wilds Base']
        self._board_full_of_wilds_counter_free = data['Board Full of Wilds Free']
        self._spectacular_link_counter_from_base = data['Common Counter from Base']

        self._base_and_base_respin_win_with_wilds = super()._ReadBasicMap(data['Base and Base Respin Win with wilds'],
                                                                          key_name='win',
                                                                          val_name='count')
        self._base_win_no_wilds = super()._ReadBasicMap(data['Base win without wilds'],
                                                        key_name='win',
                                                        val_name='count')


    def GetOne_Regular_FreeSpinWin(self):
        return self._one_free_spin_regular_win

    def GetOne_Special_FreeSpinWin(self):
        return self._one_free_spin_special_win

    def GetScatterCountBase(self):
        return self._scatter_counter_base

    def GetScatterCountBaseRespin(self):
        return self._scatter_counter_base_respin

    def GetWildCountBase(self):
        return self._wild_counter_base

    def GetWildPositionsCounterBase(self):
        return self._wild_positions_counter_base

    def GetWildPatternCounterBase(self):
        return self._wild_pattern_counter_base

    def GetBoardFullOfWildsCounterBase(self):
        return self._board_full_of_wilds_counter_base

    def GetBoardFullOfWildsCounterFree(self):
        return self._board_full_of_wilds_counter_free

    def GetSLCounterFromBase(self):
        return self._spectacular_link_counter_from_base

    def GetBaseWinWithWilds(self):
        return self._base_and_base_respin_win_with_wilds

    def GetBaseWinNoWilds(self):
        return self._base_win_no_wilds

if __name__ == "__main__":
    file = open(
        r"C:\Users\VitalijSaiganov\PycharmProjects\SlotHandler\Game_Source_Data\NBN\86\NBN_usual_86_no_gamble_stats.json")
    data = json.load(file)

    f = NBNStatistics()
    f.ReadStatistics(data)
    pd.set_option('display.max_rows', None)
    pd.set_option('display.max_columns', None)
    pd.set_option('display.width', None)
    pd.set_option('display.max_colwidth', 100)

    print(f._wild_pattern_counter_base)

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


