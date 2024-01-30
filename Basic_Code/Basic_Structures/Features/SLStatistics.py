from Basic_Code.Basic_Structures.BasicStatistics import BasicStatistics
from Basic_Code.Basic_Calculator.BasicStatsCalculator import BasicStatsCalculator
import numpy as np
import pandas as pd
import json as j


class SLStatistics(BasicStatistics):
    def __init__(self):
        super().__init__()
        self._bet = -1
        self._bet_mult = -1

        self._base_game_spin_count = -1
        self._sl_features_count = -1
        self._grand_pot_counter = -1
        self._mini_pot_counter = -1

        self._sl_scatter_mult_win = dict()
        self._sl_scatter_counter = np.array
        self._one_scat_mult = dict()
        self._two_scat_mult = dict()
        self._three_scat_mult = dict()
        self._scatter_counter_by_reels = np.array

        self._sl_respin_counter = dict()
        self._sl_symbols_in_last_board_counter = dict()
        self._sl_scatter_type_counter = dict()
        self._sl_filled_cells_counter = dict()
        self._sl_win_mult = dict()
        self._sl_win_cents = dict()

        self._sl_upgrade_on_board_counter = np.array
        self._sl_upgrade_triggered_counter = np.array

        self._sl_respin_on_which_elimination_occured = dict()
        self._sl_numbers_of_cells_filled_by_spinner = dict()
        self._sl_numbers_of_cells_filled_by_whistle = dict()

        self._win_counter_by_respin = dict()
        self._win_counter_before_and_after_elimination = np.array
        self._number_of_filled_cells_after_one_respin = np.array

        self._pots_frequencies_df = pd.DataFrame

        self._average_scatter_win = -1
        self._max_simulation_scatter_win = -1
        self._scatter_win_df = pd.DataFrame

        self._scatter_drop_main_df = pd.DataFrame
        self._scatter_drop_merged_df = pd.DataFrame

        self._respin_count_main_df = pd.DataFrame
        self._respin_count_merged_df = pd.DataFrame

        self._filled_cells_count_main_df = pd.DataFrame
        self._filled_cells_count_merged_df = pd.DataFrame

        self._sl_spin_win_df = pd.DataFrame

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
        self._spinner_index = 0
        self._whistle_index = 1
        self._elimination_index = 2

        self._upgrade_freq_df = pd.DataFrame

        self._number_of_filled_cells_main_df = pd.DataFrame
        self._number_of_filled_cells_merged_df = pd.DataFrame
        self._total_number_of_respins = -1
        self._respin_win_1_in = -1


    def ReadStatistics(self, data: dict):
        super().ReadStatistics(data)
        self._bet = data['Current Bet']
        self._bet_mult = data['feature Bet Multiplier']

        sl_data = data['CF Stats']
        self._base_game_spin_count = sl_data['Base Game Spin Count']
        self._sl_features_count = sl_data['Common Game Counter']
        self._grand_pot_counter = sl_data['Grand Pot Counter']
        self._mini_pot_counter = sl_data['Mini Pot Counter']

        self._sl_scatter_mult_win = super()._ReadBasicMap(sl_data['Common Scatter Mult Win'],
                                                          key_name='win_mult',
                                                          val_name='counter')
        self._sl_scatter_counter = super()._ReadBasicVector(sl_data['CF Scatter Counter'],
                                                            index_name='number_of_scatters',
                                                            val_name='counter')
        self._one_scat_mult = super()._ReadBasicMap(sl_data['CF One scatter mult'],
                                                    key_name='mult',
                                                    val_name='counter')
        self._two_scat_mult = super()._ReadBasicMap(sl_data['CF Two scatters mult'],
                                                    key_name='mult',
                                                    val_name='counter')
        self._three_scat_mult = super()._ReadBasicMap(sl_data['CF Three scatter mult'],
                                                      key_name='mult',
                                                      val_name='counter')

        self._scatter_counter_by_reels = super()._ReadBasicVector(sl_data['CF Scatters by reels'],
                                                                  index_name='reel_index',
                                                                  val_name='counter')

        self._sl_respin_counter = super()._ReadBasicMap(sl_data['CF Respin counter'],
                                                        key_name='number_of_spins',
                                                        val_name='counter')
        self._sl_symbols_in_last_board_counter = self._ReadBasicMapSL(sl_data['CF Symbol in last board counter'],
                                                                      key_name='symbol_id',
                                                                      val_name='counter')
        self._sl_scatter_type_counter = super()._ReadBasicMap(sl_data['CF Scatter Type Counter'],
                                                              key_name='symbol_id',
                                                              val_name='counter')
        self._sl_filled_cells_counter = super()._ReadBasicMap(sl_data['CF filled cells counter'],
                                                              key_name='number_of_cells',
                                                              val_name='counter')
        self._sl_win_mult = super()._ReadBasicMap(sl_data['CF win mult'],
                                                  key_name='win_mult',
                                                  val_name='counter')
        for win_mult, counter in self._sl_win_mult.items():
            self._sl_win_cents[win_mult * self._bet] = counter

        self._sl_upgrade_on_board_counter = super()._ReadBasicVector(sl_data['CF upgrade on board'],
                                                                     index_name='upgrade_index',
                                                                     val_name='counter')
        self._sl_upgrade_triggered_counter = super()._ReadBasicVector(sl_data['CF upgrade triggered'],
                                                                      index_name='upgrade_index',
                                                                      val_name='counter')

        self._sl_respin_on_which_elimination_occured = super()._ReadBasicMap(sl_data['CF spin of which elimination happened'],
                                                                             key_name='spin',
                                                                             val_name='counter')
        self._sl_numbers_of_cells_filled_by_spinner = super()._ReadBasicMap(sl_data['CF number of cells filled by spinner'],
                                                                            key_name='number_of_cells',
                                                                            val_name='counter')
        self._sl_numbers_of_cells_filled_by_whistle = super()._ReadBasicMap(
            sl_data['CF number of cells filled by whistle'],
            key_name='number_of_cells',
            val_name='counter')

        self._win_counter_by_respin = super()._ReadBasicMap(sl_data['Win counter by respin'],
                                                            key_name='available_respin',
                                                            val_name='counter')
        self._win_counter_before_and_after_elimination = super()._ReadBasicVector(sl_data['Win counter before/after elimination'],
                                                                                  index_name='before/after',
                                                                                  val_name='counter')
        self._number_of_filled_cells_after_one_respin = super()._ReadBasicVector(sl_data['Number of filled cells after spin'],
                                                                                 index_name='number_of_cells',
                                                                                 val_name='counter')
        self.CalcSLStats()

    def CalcSLStats(self):
        self._CalcWinRanges(self._bet)
        self._CalcPotsFrequencies_DF()
        self._CalcScattersWin_DF()
        self._scatter_drop_main_df, self._scatter_drop_merged_df = BasicStatsCalculator.CalcVectorVariant_DF(self._sl_scatter_counter,
                                                                                                             [[0, 0], [1, 3], [4, 6], [7, 9], [10, 12]],
                                                                                                             super().GetTotalSpinCount())
        self._respin_count_main_df, self._respin_count_merged_df = BasicStatsCalculator.CalcMapVariant_DF(self._sl_respin_counter,
                                                                                                          [[0, 7], [8, 9], [10, 12], [13, 15],
                                                                                                           [16, 20], [21, 25], [26, 30], [31, 35], [36, 40]],
                                                                                                          super().GetTotalSpinCount())
        self._total_number_of_respins = np.sum(self._respin_count_main_df['count'] * self._respin_count_main_df.index)
        self._filled_cells_count_main_df, self._filled_cells_count_merged_df = BasicStatsCalculator.CalcMapVariant_DF(self._sl_filled_cells_counter,
                                                                                                                      [[0, 5], [6, 8], [9, 11],
                                                                                                                       [12, 15], [16, 20], [21, 25],
                                                                                                                       [26, 30], [31, 35]],
                                                                                                                      super().GetTotalSpinCount())
        self._sl_spin_win_df = BasicStatsCalculator.CalculateCommonSpinWinDF(self._sl_win_cents,
                                                                             self._win_ranges,
                                                                             self._win_ranges_names,
                                                                             self._bet,
                                                                             super().GetTotalSpinCount())
        self._CalcUpgradeFreq_DF()

        self._number_of_filled_cells_main_df, self._number_of_filled_cells_merged_df = self._CalcNumberOfFilledCells_DF(self._number_of_filled_cells_after_one_respin,
                                                                                                                                 [[1, 1], [2, 2], [3, 5],
                                                                                                                                  [6, 8]],
                                                                                                                                 super().GetTotalSpinCount())

    def _CalcWinRanges(self, bet):
        self._win_ranges = [[0, 0],
                            [1, np.floor(.5 * bet)],
                            [np.floor(.5 * bet) + 1, bet - 1],
                            [bet, bet],
                            [bet + 1, (2 * bet) - 1],
                            [2 * bet, (3 * bet) - 1],
                            [3 * bet, (5 * bet) - 1],
                            [5 * bet, (10 * bet) - 1],
                            [10 * bet, (20 * bet) - 1],
                            [20 * bet, (30 * bet) - 1],
                            [30 * bet, (50 * bet) - 1],
                            [50 * bet, (100 * bet) - 1],
                            [100 * bet, (200 * bet) - 1],
                            [200 * bet, (500 * bet) - 1],
                            [500 * bet, (1000 * bet) - 1],
                            [1000 * bet, np.Infinity]]

    def _CalcPotsFrequencies_DF(self):
        columns = ['pot_name', 'pot_counter', 'multiplier', 'pulls_to_hit_sl', 'pulls_to_hit_base']
        rows = ['Mini', 'Minor', 'Major', 'Grand']
        mults = [50, 100, 250, 500]
        pots_counter = [self._sl_symbols_in_last_board_counter.get(30, 0),
                        self._sl_symbols_in_last_board_counter.get(31, 0),
                        self._sl_symbols_in_last_board_counter.get(32, 0),
                        self._sl_symbols_in_last_board_counter.get(33, 0)]

        self._pots_frequencies_df = pd.DataFrame(columns=columns, index=[*range(len(rows))])
        self._pots_frequencies_df['pot_name'] = rows
        self._pots_frequencies_df['multiplier'] = mults
        self._pots_frequencies_df['pot_counter'] = pots_counter
        self._pots_frequencies_df['pulls_to_hit_sl'] = self._sl_features_count / self._pots_frequencies_df['pot_counter']
        self._pots_frequencies_df['pulls_to_hit_base'] = super().GetTotalSpinCount() / self._pots_frequencies_df['pot_counter']

    def _CalcScattersWin_DF(self):
        win_ranges_names = ["[4x, 4x]",
                            "[5x, 5x]",
                            "[6x, 6x]",
                            "[7x, 7x]",
                            "[8x, 8x]",
                            "[9x, 9x]",
                            "[10x, 10x]",
                            "[11x, 13x]",
                            "[14x, 17x]",
                            "[18x, 23x]",
                            "[24x, 35x]",
                            "[36x, 80x]",
                            "[81x, 150x]",
                            "[151x, 200x]",
                            "[201x, 300x]",
                            "[301x, inf)"]
        win_ranges = [[4, 4],
                      [5, 5],
                      [6, 6],
                      [7, 7],
                      [8, 8],
                      [9, 9],
                      [10, 10],
                      [11, 13],
                      [14, 17],
                      [18, 23],
                      [24, 35],
                      [36, 80],
                      [81, 150],
                      [151, 250],
                      [251, 300],
                      [301, np.inf]]

        columns = ['total_win', 'total_counter', 'max_win', 'rtp', 'rtp_big', 'avg_win', 'win_1_in']

        range_total_win = np.array([0 for _ in range(len(win_ranges_names))], dtype='uint64')
        range_counter = np.array([0 for _ in range(len(win_ranges_names))], dtype='uint64')
        range_max_win = np.array([0 for _ in range(len(win_ranges_names))], dtype='uint64')

        for win_mult, counter in self._sl_scatter_mult_win.items():
            current_win = win_mult * counter
            for i, (left_border, right_border) in enumerate(win_ranges):
                if left_border <= win_mult <= right_border:
                    range_total_win[i] += current_win
                    range_counter[i] += counter
                    if range_max_win[i] < win_mult:
                        range_max_win[i] = win_mult
                    break
        range_total_win *= self._bet
        range_max_win *= self._bet
        self._scatter_win_df = pd.DataFrame(data=np.array([range_total_win,
                                                           range_counter,
                                                           range_max_win]).transpose(),
                                            index=win_ranges_names,
                                            columns=columns[:3],
                                            dtype='uint64')
        self._scatter_win_df['rtp'] = self._scatter_win_df['total_win'] / (self._bet * super().GetTotalSpinCount()) * 100
        self._scatter_win_df['rtp_big'] = self._scatter_win_df['total_win'] / np.sum(self._scatter_win_df['total_win']) * 100
        self._scatter_win_df['avg_win'] = self._scatter_win_df['total_win'] / self._scatter_win_df['total_counter']
        self._scatter_win_df['win_1_in'] = np.sum(self._scatter_win_df['total_counter']) / self._scatter_win_df['total_counter']
        self._scatter_win_df['total_avg_win'] = np.sum(self._scatter_win_df['total_win']) / \
            np.sum(self._scatter_win_df['total_counter'])
        self._scatter_win_df['win_1_in_small'] = self._total_spin_count / self._scatter_win_df['total_counter']
        self._scatter_win_df['avg_win_with_zero'] = np.sum(self._scatter_win_df['total_win']) / np.sum(self._scatter_win_df['total_counter'])
        self._scatter_win_df['avg_win_no_zero'] = self._scatter_win_df['avg_win_with_zero']
        if 0 in self._scatter_win_df.index:
            self._scatter_win_df['avg_win_no_zero'] = np.sum(self._scatter_win_df['total_win']) / (np.sum(
                self._scatter_win_df['total_counter']) - self._scatter_win_df['total_counter'].loc[0])
        self._scatter_win_df['win_1_in_total_feature'] = np.sum(self._scatter_win_df['total_counter']) / np.sum(self._scatter_win_df['total_counter'])
        if 0 in self._scatter_win_df.index:
            self._scatter_win_df['win_1_in_total_feature'] = np.sum(self._scatter_win_df['total_counter']) / (np.sum(
                self._scatter_win_df['total_counter']) - self._scatter_win_df['total_counter'].loc[0])

        self._scatter_win_df['win_1_in_total_base'] = self._total_spin_count / np.sum(
            self._scatter_win_df['total_counter'])
        if 0 in self._scatter_win_df.index:
            self._scatter_win_df['win_1_in_total_base'] = self._total_spin_count / (np.sum(
                self._scatter_win_df['total_counter']) - self._scatter_win_df['total_counter'].loc[0])

        self._average_scatter_win = np.sum(self._scatter_win_df['avg_win'] * self._scatter_win_df['total_counter'] / np.sum(self._scatter_win_df['total_counter']))
        self._max_simulation_scatter_win = self._scatter_win_df['max_win'].iloc[-1]

    def _CalcUpgradeFreq_DF(self):
        columns = ['upgrade_name', 'on_board_count', 'trigger_count', '1_in_board', '1_in_triggered_in_board', '1_in_triggered_feature', 'percent_of_trigger_when_in_board']
        index = np.arange(0, len(self._sl_upgrade_on_board_counter))

        self._upgrade_freq_df = pd.DataFrame(columns=columns, index=index)
        self._upgrade_freq_df['upgrade_name'].iloc[0] = 'Spinner'
        self._upgrade_freq_df['upgrade_name'].iloc[1] = 'Whistle'
        self._upgrade_freq_df['upgrade_name'].iloc[2] = 'Elimination'

        self._upgrade_freq_df['on_board_count'] = self._sl_upgrade_on_board_counter
        self._upgrade_freq_df['trigger_count'] = self._sl_upgrade_triggered_counter
        self._upgrade_freq_df['1_in_board'] = self._sl_features_count / self._upgrade_freq_df['on_board_count']
        self._upgrade_freq_df['1_in_triggered_in_board'] = self._upgrade_freq_df['on_board_count'] / self._upgrade_freq_df['trigger_count']
        self._upgrade_freq_df['1_in_triggered_feature'] = self._sl_features_count / self._upgrade_freq_df['trigger_count']
        self._upgrade_freq_df['percent_of_trigger_when_in_board'] = self._upgrade_freq_df['trigger_count'] / self._upgrade_freq_df['on_board_count'] * 100

    def _CalcNumberOfFilledCells_DF(self, variants: list, join_ranges: list, total_game_spin_count: int):
        columns = ['count', 'pulls_to_hit_big', 'pulls_to_hit_small', 'percent_big', 'percent_small']
        index_main = np.arange(1, len(variants)+1)
        index_merged = np.arange(0, len(join_ranges))
        join_ranges[-1][1] = len(variants) - 1

        main_table = pd.DataFrame(columns=columns, index=index_main)
        merged_table = pd.DataFrame(columns=['left_var', 'right_var'] + columns, index=index_merged)

        main_table['count'] = variants
        main_table['pulls_to_hit_big'] = np.sum(main_table['count']) / main_table['count']
        main_table['pulls_to_hit_small'] = total_game_spin_count / main_table['count']
        main_table['percent_big'] = main_table['count'] / np.sum(main_table['count']) * 100
        main_table['percent_small'] = main_table['count'] / total_game_spin_count * 100
        main_table['avg_val_with_zero'] = np.sum(main_table.index * main_table['count']) / (
            np.sum(main_table['count']))
        main_table['avg_val_without_zero'] = main_table['avg_val_with_zero']
        if 0 in main_table.index:
            main_table['avg_val_without_zero'] = np.sum(main_table.index * main_table['count']) / (
                        np.sum(main_table['count']) - main_table['count'].loc[0])
        main_table['any_variant_1_in_base_with_zero'] = self._total_spin_count / np.sum(main_table['count'])
        main_table['any_variant_1_in_base_no_zero'] = main_table['any_variant_1_in_base_with_zero']
        if 0 in main_table.index:
            main_table['any_variant_1_in_base_no_zero'] = self._total_spin_count / np.sum(main_table['count'] - main_table['count'].loc[0])
        main_table['any_variant_1_in_feature_no_zero'] = self._sl_features_count / np.sum(main_table['count'])
        if 0 in main_table.index:
            main_table['any_variant_1_in_feature_no_zero'] = self._sl_features_count / np.sum(main_table['count'] - main_table['count'].loc[0])
        main_table.drop(main_table[main_table['count'] == 0].index, inplace=True)

        for i, (left_border, right_border) in enumerate(join_ranges):
            merged_table['left_var'].iloc[i] = left_border
            merged_table['right_var'].iloc[i] = right_border
            merged_table['count'].iloc[i] = np.sum(main_table['count'].iloc[left_border: right_border + 1])
        merged_table['pulls_to_hit_big'] = np.sum(merged_table['count']) / merged_table['count']
        merged_table['pulls_to_hit_small'] = total_game_spin_count / merged_table['count']
        merged_table['percent_big'] = merged_table['count'] / np.sum(merged_table['count']) * 100
        merged_table['percent_small'] = merged_table['count'] / total_game_spin_count * 100
        merged_table.drop(merged_table[merged_table['count'] == 0].index, inplace=True)
        merged_table.reset_index(inplace=True)

        self._respin_win_1_in = self._total_number_of_respins / np.sum(main_table['count'])
        return main_table, merged_table


    def GetBaseGameSpinCount(self):
        return self._base_game_spin_count

    def GetSLFeaturesCount(self):
        return self._sl_features_count

    def GetMiniPotCounter(self):
        return self._mini_pot_counter

    def GetGrandPotCounter(self):
        return self._grand_pot_counter

    def GetSLScatterMultWin(self):
        return self._sl_scatter_mult_win

    def GetSLScatterCounter(self):
        return self._sl_scatter_counter

    def GetSLOneScatterMult(self):
        return self._one_scat_mult

    def GetSLTwoScattersMult(self):
        return self._two_scat_mult

    def GetSLThreeScattersMult(self):
        return self._three_scat_mult

    def GetSLScattersByReelsCounter(self):
        return self._scatter_counter_by_reels

    def GetSLRespinCounter(self):
        return self._sl_respin_counter

    def GetSymbolsInLastBoardCounter(self):
        return self._sl_symbols_in_last_board_counter

    def GetSLScattersTypeCounter(self):
        return self._sl_scatter_type_counter

    def GetSLFilledCellsCounter(self):
        return self._sl_filled_cells_counter

    def GetWinMult(self):
        return self._sl_win_mult

    def _ReadBasicMapSL(self, data: dict, key_name: str, val_name: str):
        res = dict()
        for value_index, value in enumerate(data):
            key = value[key_name]
            if isinstance(key, str):
                key = int(key)
            res[key] = value[val_name]
        return res

    def GetPotsFrequenciesDF(self):
        return self._pots_frequencies_df

    def GetScatterWin_DF(self):
        return self._scatter_win_df

    def GetScatterAvgWin(self):
        return self._average_scatter_win

    def GetScatterMaxWin(self):
        return self._max_simulation_scatter_win

    def GetSLScatterCounter_DFs(self):
        return self._scatter_drop_main_df, self._scatter_drop_merged_df

    def GetSLRespinCounter_DFs(self):
        return self._respin_count_main_df, self._respin_count_merged_df

    def GetSLFilledCells_DFs(self):
        return self._filled_cells_count_main_df, self._filled_cells_count_merged_df

    def GetSLFilledCellsOneSpin_DFs(self):
        return self._number_of_filled_cells_main_df, self._number_of_filled_cells_merged_df

    def GetSLSpinWin_DF(self):
        return self._sl_spin_win_df

    def GetSLScattersWin_DF(self):
        return self._scatter_win_df

    def GetSLUpgradesFreq_DF(self):
        return self._upgrade_freq_df

    def GetSLPotsFreq_DF(self):
        return self._pots_frequencies_df


if __name__ == "__main__":
    file = open(
        r"C:\Users\VitalijSaiganov\PycharmProjects\SlotHandler\Game_Source_Data\NBN\86\NBN_usual_86_no_gamble_stats.json")
    data = j.load(file)

    f = SLStatistics()
    f.ReadStatistics(data)

    pd.set_option('display.max_rows', None)
    pd.set_option('display.max_columns', None)
    pd.set_option('display.width', None)
    pd.set_option('display.max_colwidth', 100)
    f.CalcSLStats()

    print()
    main = f.GetSLFilledCellsOneSpin_DFs()
    print(main)

    # base_df = f.GetScatterWin_DF()
    # print(base_df)
    # print(f.GetScatterAvgWin())
    # print(f.GetScatterMaxWin())




    # Stats:
    # 1. SL frequency, 1 in base spins ЭТО
    # 2. Pots frequency mini, minor, major, grand  ЭТО
    #
    # 3. SL scatters avg win + distribution  ЭТО
    # 4. SL scatter counter  ЭТО
    # 5. SL one scat mult avg + distribution
    # 6. SL two scatters mult avg + distribution
    # 7. SL three scatters avg mult + distribution
    # 8. SL scatters distribution by reels
    #
    # 9. SL respin avg counter + distribution  ЭТО
    # 10. SL distribution of symbols on last board
    # 11. SL distribution of scatter type counter
    # 12. SL filled cells counter avg + distribution  ЭТО
    # 13. SL win mult distribution  ЭТО
    # 14. SL upgrade on board + upgrade triggered  ЭТО
    # 15. SL avg respin on which elimination occured
    # 16. SL distribution of number of cells filled by spinner
    # 17. SL distribution of number of cells filled by whistle
    # 18. SL win distribution by respin (1,2,3)
    # 19. Win distribution before/after elimination
    # 20. number of filled cells after spin  ЭТО


