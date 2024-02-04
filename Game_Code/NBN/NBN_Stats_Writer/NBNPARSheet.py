from Basic_Code.Stats_Writer.BasicPARSheet import BasicPARSheet
from Game_Code.NBN.NBN_Structures.NBNSlot import NBNSlot
from Game_Code.NBN.NBN_Structures.NBNStatistics import NBNStatistics
from Game_Code.NBN.NBN_Calculator.NBNStatsCalculator import NBNStatsCalculator
from Game_Code.NBN.NBN_Stats_Writer.NBNPARFormats import NBNPARFormats
from Basic_Code.Stats_Writer.BasicGraphMaker import Colors
import xlsxwriter as xl
import json as j
import numpy as np
import pandas as pd


class NBNPARSheet(BasicPARSheet):
    def __init__(self, book: xl.Workbook, formats: NBNPARFormats, slot: NBNSlot, stats: NBNStatistics, calculation: NBNStatsCalculator):
        super().__init__(book, formats, slot, stats, calculation)
        self._slot = slot
        self._stats = stats
        self._calcualtion = calculation

        self._allocated_rows_for_content = 55

    def WriteCustomGameStatistics(self):
        self._WriteWildFeatureStats()
        self._WriteFreeSpinStats()
        self._WriteSLStats()

    def _WriteFreeSpinStats(self):
        self.WriteHeader("Free Spins Statistics", include_in_content=True, print_counter=True)

        #self._WriteFreeSpinsOneRegularWin()
        #self._WriteFreeSpinsOneSpecialWin()
        self.WriteFeatureRespinSettings(0)
        self._WriteFreeScattersCountBase()
        self._WriteFreeScattersCountBaseRespin()



    # def _WriteFreeSpinsOneRegularWin(self):
    #     self.WriteInnerHeader("Regular Free Spin and Re-Spin Win Distribution", include_in_content=True,
    #                           print_counter=True)
    #     self.WriteBasicSpinWin(self._calcualtion.GetOneFreeSpinRegularWin_DF(),
    #                            graph_name='one_regular_free_win',
    #                            color_index=1,
    #                            print_avg_win_no_zero=True,
    #                            print_avg_win_with_zero=True,
    #                            print_win_1_in_feature=True,
    #                            print_win_1_in_base=True,
    #                            print_win_freq_base_col=True,
    #                            print_max_range_win_col=True)
    #
    # def _WriteFreeSpinsOneSpecialWin(self):
    #     self.WriteInnerHeader("Special Free Spin and Re-Spin Win Distribution", include_in_content=True,
    #                           print_counter=True)
    #     self.WriteBasicSpinWin(self._calcualtion.GetOneFreeSpinSpecialWin_DF(),
    #                            graph_name='one_special_free_win',
    #                            color_index=2,
    #                            print_avg_win_no_zero=True,
    #                            print_avg_win_with_zero=True,
    #                            print_win_1_in_feature=True,
    #                            print_win_1_in_base=True,
    #                            print_win_freq_base_col=True,
    #                            print_max_range_win_col=True)

    def _WriteFreeScattersCountBase(self):
        self.WriteInnerHeader("Free Spins Scatters Count Base", include_in_content=True, print_counter=True)

        scatter_count_main_df, scatter_count_merged_df = self._calcualtion.GetFreeScatterCountBase_DFs()
        scatter_count_labels = {'avg_variant_with_zero': 'Average number of scatters:',
                                'avg_variant_no_zero': 'Average number of scatters:',
                                'any_variant_1_in_base_with_zero': 'Pulls to hit any scatters:',
                                'any_variant_1_in_base_no_zero': 'Pulls to hit any scatters:',
                                'any_variant_1_in_feature_no_zero': 'Pulls to hit any scatters:',
                                'plus_variants_1_in_base': 'Pulls to hit {:d}+ scatters:',
                                'plus_variants_1_in_feature': 'Pulls to hit {:d}+ scatters:',
                                'picture_col': 'Number of Scatters in View Distribution',
                                'variant_col': 'Scatters Count',
                                'pulls_to_hit_big': 'Pulls to Hit',
                                'pulls_to_hit_big_merged': 'Merged Ranges',
                                'pulls_to_hit_small': 'Pulls to Hit',
                                'pulls_to_hit_small_merged': 'Merged Ranges',
                                'percent_big': 'Percent',
                                'percent_big_merged': 'Merged Ranges',
                                'percent_small': 'Percent',
                                'percent_small_merged': 'Merged Ranges',
                                'graph_name': 'free_scatter_base_count'
                                }

        self.WriteBasicVariant(main_df=scatter_count_main_df,
                               merged_df=scatter_count_merged_df,
                               variant_df_col_name='count',
                               txt_labels=scatter_count_labels,
                               variant_is_mult=False,
                               color_index=0,
                               n=(3,),
                               include_zero_in_total=False,
                               print_n_plus_1_in_base=True,
                               print_n_plus_1_in_feature=False,
                               print_avg_variant_with_zero=False,
                               print_avg_variant_no_zero=True,
                               print_any_variant_1_in_base_no_zero=True,
                               print_any_variant_1_in_base_with_zero=False,
                               print_any_variant_1_in_feature_no_zero=False,
                               print_pulls_to_hit_big_col=False,
                               print_pulls_to_hit_small_col=True,
                               print_percent_big_col=False,
                               print_percent_small_col=True)

    def _WriteFreeScattersCountBaseRespin(self):
        self.WriteInnerHeader("Free Spins Scatters Count Base Re-Spin", include_in_content=True, print_counter=True)

        scatter_count_main_df, scatter_count_merged_df = self._calcualtion.GetFreeScatterCountBaseRespin_DFs()
        scatter_count_labels = {'avg_variant_with_zero': 'Average number of scatters:',
                                'avg_variant_no_zero': 'Average number of scatters:',
                                'any_variant_1_in_base_with_zero': 'Pulls to hit any scatters:',
                                'any_variant_1_in_base_no_zero': 'Pulls to hit any scatters:',
                                'any_variant_1_in_feature_no_zero': 'Pulls to hit any scatters:',
                                'plus_variants_1_in_base': 'Pulls to hit {:d}+ scatters:',
                                'plus_variants_1_in_feature': 'Pulls to hit {:d}+ scatters:',
                                'picture_col': 'Number of Scatters in View Distribution',
                                'variant_col': 'Scatters Count',
                                'pulls_to_hit_big': 'Pulls to Hit',
                                'pulls_to_hit_big_merged': 'Merged Ranges',
                                'pulls_to_hit_small': 'Pulls to Hit',
                                'pulls_to_hit_small_merged': 'Merged Ranges',
                                'percent_big': 'Percent',
                                'percent_big_merged': 'Merged Ranges',
                                'percent_small': 'Percent',
                                'percent_small_merged': 'Merged Ranges',
                                'graph_name': 'free_scatter_base_respin_count'
                                }
        self.WriteBasicVariant(main_df=scatter_count_main_df,
                               merged_df=scatter_count_merged_df,
                               variant_df_col_name='count',
                               txt_labels=scatter_count_labels,
                               variant_is_mult=False,
                               color_index=0,
                               n=(3,),
                               include_zero_in_total=False,
                               print_n_plus_1_in_base=True,
                               print_n_plus_1_in_feature=True,
                               print_avg_variant_with_zero=False,
                               print_avg_variant_no_zero=True,
                               print_any_variant_1_in_base_no_zero=True,
                               print_any_variant_1_in_base_with_zero=False,
                               print_any_variant_1_in_feature_no_zero=True,
                               print_pulls_to_hit_big_col=True,
                               print_pulls_to_hit_small_col=True,
                               print_percent_big_col=True,
                               print_percent_small_col=True)

    def _WriteWildFeatureStats(self):
        self.WriteHeader("Wild Feature Statistics", include_in_content=True, print_counter=True)

        self._WriteNumberofWildsDistribution()
        self._WriteWildPositionsDistribution()

    def _WriteNumberofWildsDistribution(self):
        self.WriteInnerHeader('Number of Wilds Distribution Base Game', include_in_content=True, print_counter=True)

        wild_count_base_main_df, wild_count_base_merged_df = self._calcualtion.GetWildCountBase_DFs()
        wild_count_base_labels = {'avg_variant_with_zero': 'Average number of Wilds:',
                                  'avg_variant_no_zero': 'Average number of Wild:',
                                  'any_variant_1_in_base_with_zero': 'Pulls to hit any Wilds:',
                                  'any_variant_1_in_base_no_zero': 'Pulls to hit any Wilds:',
                                  'any_variant_1_in_feature_no_zero': 'Pulls to hit any Wilds:',
                                  'plus_variants_1_in_base': 'Pulls to hit {:d}+ Wilds:',
                                  'plus_variants_1_in_feature': 'Pulls to hit {:d}+ Wilds:',
                                  'picture_col': 'Number of Wilds in View Distribution',
                                  'variant_col': 'Wilds Count',
                                  'pulls_to_hit_big': 'Pulls to Hit',
                                  'pulls_to_hit_big_merged': 'Merged Ranges',
                                  'pulls_to_hit_small': 'Pulls to Hit',
                                  'pulls_to_hit_small_merged': 'Merged Ranges',
                                  'percent_big': 'Percent',
                                  'percent_big_merged': 'Merged Ranges',
                                  'percent_small': 'Percent',
                                  'percent_small_merged': 'Merged Ranges',
                                  'graph_name': 'wild_count_base'
                                  }
        self.WriteBasicVariant(main_df=wild_count_base_main_df,
                               merged_df=wild_count_base_merged_df,
                               variant_df_col_name='count',
                               txt_labels=wild_count_base_labels,
                               variant_is_mult=False,
                               color_index=1,
                               n=(3, 5, 7, 10, 15),
                               include_zero_in_total=False,
                               print_n_plus_1_in_base=True,
                               print_n_plus_1_in_feature=False,
                               print_avg_variant_with_zero=False,
                               print_avg_variant_no_zero=True,
                               print_any_variant_1_in_base_no_zero=True,
                               print_any_variant_1_in_base_with_zero=False,
                               print_any_variant_1_in_feature_no_zero=False,
                               print_pulls_to_hit_big_col=True,
                               print_pulls_to_hit_small_col=False,
                               print_percent_big_col=True,
                               print_percent_small_col=False)

        # self.WriteInnerHeader('Number of Wilds Distribution Free Game', include_in_content=True, print_counter=True)
        #
        # wild_count_free_main_df, wild_count_free_merged_df = self._calcualtion.GetWildCountFree_DFs()
        # wild_count_free_labels = {'avg_variant_with_zero': 'Average number of Wilds:',
        #                           'avg_variant_no_zero': 'Average number of Wild:',
        #                           'any_variant_1_in_base_with_zero': 'Pulls to hit any Wilds:',
        #                           'any_variant_1_in_base_no_zero': 'Pulls to hit any Wilds:',
        #                           'any_variant_1_in_feature_no_zero': 'Pulls to hit any Wilds:',
        #                           'plus_variants_1_in_base': 'Pulls to hit {:d}+ Wilds:',
        #                           'plus_variants_1_in_feature': 'Pulls to hit {:d}+ Wilds:',
        #                           'picture_col': 'Number of Wilds in View Distribution',
        #                           'variant_col': 'Wilds Count',
        #                           'pulls_to_hit_big': 'Pulls to Hit',
        #                           'pulls_to_hit_big_merged': 'Merged Ranges',
        #                           'pulls_to_hit_small': 'Pulls to Hit',
        #                           'pulls_to_hit_small_merged': 'Merged Ranges',
        #                           'percent_big': 'Percent',
        #                           'percent_big_merged': 'Merged Ranges',
        #                           'percent_small': 'Percent',
        #                           'percent_small_merged': 'Merged Ranges',
        #                           'graph_name': 'wild_count_free'
        #                           }
        # self.WriteBasicVariant(main_df=wild_count_free_main_df,
        #                        merged_df=wild_count_free_merged_df,
        #                        variant_df_col_name='count',
        #                        txt_labels=wild_count_free_labels,
        #                        variant_is_mult=False,
        #                        color_index=1,
        #                        n=(3, 5, 7, 10, 15),
        #                        include_zero_in_total=False,
        #                        print_n_plus_1_in_base=False,
        #                        print_n_plus_1_in_feature=True,
        #                        print_avg_variant_with_zero=False,
        #                        print_avg_variant_no_zero=True,
        #                        print_any_variant_1_in_base_no_zero=False,
        #                        print_any_variant_1_in_base_with_zero=False,
        #                        print_any_variant_1_in_feature_no_zero=True,
        #                        print_pulls_to_hit_big_col=True,
        #                        print_pulls_to_hit_small_col=False,
        #                        print_percent_big_col=True,
        #                        print_percent_small_col=False)
        #
        # self.WriteInnerHeader('Number of Wilds Distribution Special Free Spin', include_in_content=True, print_counter=True)
        #
        # wild_count_special_main_df, wild_count_special_merged_df = self._calcualtion.GetWildCountSpecial_DFs()
        # wild_count_special_labels = {'avg_variant_with_zero': 'Average number of Wilds:',
        #                              'avg_variant_no_zero': 'Average number of Wild:',
        #                              'any_variant_1_in_base_with_zero': 'Pulls to hit any Wilds:',
        #                              'any_variant_1_in_base_no_zero': 'Pulls to hit any Wilds:',
        #                              'any_variant_1_in_feature_no_zero': 'Pulls to hit any Wilds:',
        #                              'plus_variants_1_in_base': 'Pulls to hit {:d}+ Wilds:',
        #                              'plus_variants_1_in_feature': 'Pulls to hit {:d}+ Wilds:',
        #                              'picture_col': 'Number of Wilds in View Distribution',
        #                              'variant_col': 'Wilds Count',
        #                              'pulls_to_hit_big': 'Pulls to Hit',
        #                              'pulls_to_hit_big_merged': 'Merged Ranges',
        #                              'pulls_to_hit_small': 'Pulls to Hit',
        #                              'pulls_to_hit_small_merged': 'Merged Ranges',
        #                              'percent_big': 'Percent',
        #                              'percent_big_merged': 'Merged Ranges',
        #                              'percent_small': 'Percent',
        #                              'percent_small_merged': 'Merged Ranges',
        #                              'graph_name': 'wild_count_special'
        #                              }
        # self.WriteBasicVariant(main_df=wild_count_special_main_df,
        #                        merged_df=wild_count_special_merged_df,
        #                        variant_df_col_name='count',
        #                        txt_labels=wild_count_special_labels,
        #                        variant_is_mult=False,
        #                        color_index=1,
        #                        n=(7,),
        #                        include_zero_in_total=False,
        #                        print_n_plus_1_in_base=False,
        #                        print_n_plus_1_in_feature=True,
        #                        print_avg_variant_with_zero=False,
        #                        print_avg_variant_no_zero=True,
        #                        print_any_variant_1_in_base_no_zero=False,
        #                        print_any_variant_1_in_base_with_zero=False,
        #                        print_any_variant_1_in_feature_no_zero=True,
        #                        print_pulls_to_hit_big_col=True,
        #                        print_pulls_to_hit_small_col=False,
        #                        print_percent_big_col=True,
        #                        print_percent_small_col=False)


    def _WriteWildPositionsDistribution(self):
        self.WriteInnerHeader('Wilds Positions Base Game', include_in_content=True, print_counter=True)

        str_row_colors = Colors.get_vector_of_scales(Colors.hexes[3][4], Colors.hexes[3][4], self._slot.GetBoardHeight())
        str_col_colors = Colors.get_vector_of_scales(Colors.hexes[1][5], Colors.hexes[1][8],
                                                     self._slot.GetBoardWidth())
        self.WriteBasicBoardDistribution(self._calcualtion.GetWildPositionCountBase_DFs()[0],
                                         row_count=self._slot.GetBoardHeight(),
                                         col_count=self._slot.GetBoardWidth(),
                                         graph_name='wild_positions_base',
                                         col_graph_color_index=1,
                                         row_graph_color_index=1,
                                         cells_scale_start=Colors.hexes[3][5],
                                         cells_scale_end=Colors.hexes[1][5],
                                         bar_plot_col_colors=str_col_colors,
                                         bar_plot_row_colors=str_row_colors)

        # self.WriteInnerHeader('Wilds Positions Free Game', include_in_content=True, print_counter=True)
        #
        # str_row_colors = Colors.get_vector_of_scales(Colors.hexes[3][4], Colors.hexes[3][4],
        #                                              self._slot.GetBoardHeight())
        # str_col_colors = Colors.get_vector_of_scales(Colors.hexes[1][5], Colors.hexes[1][8],
        #                                              self._slot.GetBoardWidth())
        # self.WriteBasicBoardDistribution(self._calcualtion.GetWildPositionCountFree_DFs()[0],
        #                                  row_count=self._slot.GetBoardHeight(),
        #                                  col_count=self._slot.GetBoardWidth(),
        #                                  graph_name='wild_positions_free',
        #                                  col_graph_color_index=1,
        #                                  row_graph_color_index=1,
        #                                  cells_scale_start=Colors.hexes[3][5],
        #                                  cells_scale_end=Colors.hexes[1][5],
        #                                  bar_plot_col_colors=str_col_colors,
        #                                  bar_plot_row_colors=str_row_colors)
        #
        # self.WriteInnerHeader('Wilds Positions Special Free Spin', include_in_content=True, print_counter=True)
        #
        # str_row_colors = Colors.get_vector_of_scales(Colors.hexes[3][4], Colors.hexes[3][4],
        #                                              self._slot.GetBoardHeight())
        # str_col_colors = Colors.get_vector_of_scales(Colors.hexes[1][5], Colors.hexes[1][8],
        #                                              self._slot.GetBoardWidth())
        # self.WriteBasicBoardDistribution(self._calcualtion.GetWildPositionCountSpecial_DFs()[0],
        #                                  row_count=self._slot.GetBoardHeight(),
        #                                  col_count=self._slot.GetBoardWidth(),
        #                                  graph_name='wild_positions_special',
        #                                  col_graph_color_index=1,
        #                                  row_graph_color_index=1,
        #                                  cells_scale_start=Colors.hexes[3][5],
        #                                  cells_scale_end=Colors.hexes[1][5],
        #                                  bar_plot_col_colors=str_col_colors,
        #                                  bar_plot_row_colors=str_row_colors)

    def _WriteSLStats(self):
        self.WriteHeader("Spectacular Link Statistics", include_in_content=True, print_counter=True)

        self._WriteSLMainInfo()
        self._WriteSLScattersCountStats()
        self._WriteSLScattersWinStats()
        self._WriteSLRespinCountStats()
        self._WriteSLFilledCellsStats()
        self._WriteSLWinStats()
        self._WriteSLFilledCellsOneSpin()

    def _WriteSLMainInfo(self):
        self.WriteInnerHeader('Spectacular Link Main Info', include_in_content=True, print_counter=True)
        info_table_info_width = 5
        info_table_val_width = 2
        info_table_start_col = self.GetCentreStartCol(self.GetInnerHeaderWidth(),
                                                      info_table_info_width + info_table_val_width,
                                                      self.GetInnerHeaderStartColumn())
        color_index = 0

        sl_freq_info_txt = [self._text_formats['text_big_bold'], 'Spectacular Link Triggered:',
                            self._text_formats['text_smallest'], '\n(1 in ... base spins)']
        sl_freq_val_txt = [self._text_formats['text_smallest'], '1 in ',
                           self._text_formats['text_big'], '{:.3f}'.format(self._spin_count / self._stats.GetSLFeaturesCount())]
        self.PrintInfoRow(start_col=info_table_start_col,
                          info_width=info_table_info_width,
                          val_width=info_table_val_width,
                          inf_txt_segments=sl_freq_info_txt,
                          val_txt_segments=sl_freq_val_txt,
                          bg_format_info=self._info_row_formats['info_even_info_'+str(color_index)],
                          bg_format_val=self._info_row_formats['info_even_val_'+str(color_index)],
                          row_height_in_pixels=self._big_cell_pixel_height)

        sl_avg_respins_info_txt = [self._text_formats['text_big_bold'], 'Average number of Re-Spins:']
        sl_avg_respins_val_txt = [self._text_formats['text_big'], '{:.3f}'.format(self._stats.GetSLRespinCounter_DFs()[0]['avg_val_with_zero'].iloc[0])]
        self.PrintInfoRow(start_col=info_table_start_col,
                          info_width=info_table_info_width,
                          val_width=info_table_val_width,
                          inf_txt_segments=sl_avg_respins_info_txt,
                          val_txt_segments=sl_avg_respins_val_txt,
                          bg_format_info=self._info_row_formats['info_odd_info_' + str(color_index)],
                          bg_format_val=self._info_row_formats['info_odd_val_' + str(color_index)],
                          row_height_in_pixels=self._big_cell_pixel_height)

        max_cell_count = self._stats.GetSLFilledCells_DFs()[0].index[-1]
        sl_avg_cells_info_txt = [self._text_formats['text_big_bold'], 'Average number of Filled Cells:',
                                 self._text_formats['text_smallest'], '\n(total {:d} cells)'.format(max_cell_count)]
        sl_avg_cells_val_txt = [self._text_formats['text_big'], '{:.3f}'.format(self._stats.GetSLFilledCells_DFs()[0]['avg_val_with_zero'].iloc[0])]
        self.PrintInfoRow(start_col=info_table_start_col,
                          info_width=info_table_info_width,
                          val_width=info_table_val_width,
                          inf_txt_segments=sl_avg_cells_info_txt,
                          val_txt_segments=sl_avg_cells_val_txt,
                          bg_format_info=self._info_row_formats['info_even_info_' + str(color_index)],
                          bg_format_val=self._info_row_formats['info_even_val_' + str(color_index)],
                          row_height_in_pixels=self._big_cell_pixel_height)

        sl_avg_win_info_txt = [self._text_formats['text_big_bold'], 'Average Win Amount:']
        sl_avg_win_val_txt = [self._text_formats['text_big'], '{:.3f}x'.format(self._stats.GetSLSpinWin_DF()['avg_win_no_zero'].iloc[0] / self._base_bet),
                              self._text_formats['text_smallest'], ' bets']
        self.PrintInfoRow(start_col=info_table_start_col,
                          info_width=info_table_info_width,
                          val_width=info_table_val_width,
                          inf_txt_segments=sl_avg_win_info_txt,
                          val_txt_segments=sl_avg_win_val_txt,
                          bg_format_info=self._info_row_formats['info_odd_info_' + str(color_index)],
                          bg_format_val=self._info_row_formats['info_odd_val_' + str(color_index)],
                          row_height_in_pixels=self._big_cell_pixel_height)

        sl_avg_scat_win_info_txt = [self._text_formats['text_big_bold'], 'Average SL Scatters Win Amount:']
        sl_avg_scat_win_val_txt = [self._text_formats['text_big'], '{:.3f}x'.format(self._stats.GetSLScattersWin_DF()['total_avg_win'].iloc[0] / self._base_bet),
                                   self._text_formats['text_smallest'], ' bets']
        self.PrintInfoRow(start_col=info_table_start_col,
                          info_width=info_table_info_width,
                          val_width=info_table_val_width,
                          inf_txt_segments=sl_avg_scat_win_info_txt,
                          val_txt_segments=sl_avg_scat_win_val_txt,
                          bg_format_info=self._info_row_formats['info_even_info_' + str(color_index)],
                          bg_format_val=self._info_row_formats['info_even_val_' + str(color_index)],
                          row_height_in_pixels=self._big_cell_pixel_height)

        self._current_row += 1

        self.WriteSmallHeader('SL Upgrades Frequencies')
        self._WriteSLUpgradesFreq()
        self.WriteSmallHeader('SL Pots Frequencies')
        #self._WriteSLPotFreq()

    def _WriteSLUpgradesFreq(self):
        up_name_width = 2
        up_in_board_width = 2
        up_trigger_in_board_width = 2
        up_trigger_in_feature_width = 2
        percent_width = 2
        start_col = self.GetCentreStartCol(self.GetSmallHeaderWidth(),
                                           up_name_width +
                                           up_in_board_width +
                                           up_trigger_in_board_width +
                                           up_trigger_in_feature_width +
                                           percent_width,
                                           self.GetSmallHeaderStartColumn())
        color_index = 1
        cur_col = start_col

        self._sheet.set_row_pixels(self._current_row, 60)
        self._sheet.merge_range(self._current_row, cur_col,
                                self._current_row, cur_col+up_name_width - 1,
                                '', self._rtp_distribution_formats['border_description_'+str(color_index)])
        self._sheet.write_rich_string(self._current_row, cur_col,
                                      self._text_formats['text_regular_bold'], 'Upgrade Name',
                                      self._text_formats['text_regular_bold'], ' ',
                                      self._rtp_distribution_formats['border_description_'+str(color_index)])

        cur_col += up_name_width

        self._sheet.merge_range(self._current_row, cur_col,
                                self._current_row, cur_col + up_in_board_width-1,
                                '', self._rtp_distribution_formats['border_description_' + str(color_index)])
        self._sheet.write_rich_string(self._current_row, cur_col,
                                      self._text_formats['text_regular_bold'], 'Upgrade in View',
                                      self._text_formats['text_smallest'], '\n(1 in ... SL features)',
                                      self._rtp_distribution_formats['border_description_' + str(color_index)])

        cur_col += up_in_board_width

        self._sheet.merge_range(self._current_row, cur_col,
                                self._current_row, cur_col + up_trigger_in_board_width-1,
                                '', self._rtp_distribution_formats['border_description_' + str(color_index)])
        self._sheet.write_rich_string(self._current_row, cur_col,
                                      self._text_formats['text_regular_bold'], 'Upgrade Triggered',
                                      self._text_formats['text_smallest'], '\n(when in view)',
                                      self._rtp_distribution_formats['border_description_' + str(color_index)])

        cur_col += up_trigger_in_board_width

        self._sheet.merge_range(self._current_row, cur_col,
                                self._current_row, cur_col + up_trigger_in_feature_width-1,
                                '', self._rtp_distribution_formats['border_description_' + str(color_index)])
        self._sheet.write_rich_string(self._current_row, cur_col,
                                      self._text_formats['text_regular_bold'], 'Upgrade Triggered',
                                      self._text_formats['text_smallest'], '\n(1 in ... SL features)',
                                      self._rtp_distribution_formats['border_description_' + str(color_index)])

        cur_col += up_trigger_in_feature_width

        self._sheet.merge_range(self._current_row, cur_col,
                                self._current_row, cur_col + percent_width-1,
                                '', self._rtp_distribution_formats['border_description_' + str(color_index)])
        self._sheet.write_rich_string(self._current_row, cur_col,
                                      self._text_formats['text_regular_bold'], 'Trigger Probability',
                                      self._text_formats['text_smallest'], '\n(when in view)',
                                      self._rtp_distribution_formats['border_description_' + str(color_index)])

        self._current_row += 1

        upgrade_df = self._stats.GetSLUpgradesFreq_DF()
        for i, upgrade in enumerate(upgrade_df.index):
            cur_format = self._rtp_distribution_formats['border_even_'+str(color_index)] if i % 2 == 0 else \
                self._rtp_distribution_formats['border_odd_' + str(color_index)]

            cur_col = start_col
            self._sheet.set_row_pixels(self._current_row, self.GetNormalCellPixelHeight())

            self._sheet.merge_range(self._current_row, cur_col,
                                    self._current_row, cur_col + up_name_width-1,
                                    '', cur_format)
            self._sheet.write_rich_string(self._current_row, cur_col,
                                          self._text_formats['text_regular_bold'], upgrade_df['upgrade_name'].loc[upgrade],
                                          self._text_formats['text_regular_bold'], ' ',
                                          cur_format)

            cur_col += up_name_width

            self._sheet.merge_range(self._current_row, cur_col,
                                    self._current_row, cur_col + up_in_board_width-1,
                                    '', cur_format)
            in_board = '{:,.3f}'.format(upgrade_df['1_in_board'].loc[upgrade])
            self._sheet.write_rich_string(self._current_row, cur_col,
                                          self._text_formats['text_smallest'], '1 in ',
                                          self._text_formats['text_regular'], in_board,
                                          cur_format)

            cur_col += up_in_board_width

            self._sheet.merge_range(self._current_row, cur_col,
                                    self._current_row, cur_col + up_trigger_in_board_width-1,
                                    '', cur_format)
            triggered_in_board = '{:,.3f}'.format(upgrade_df['1_in_triggered_in_board'].loc[upgrade])
            self._sheet.write_rich_string(self._current_row, cur_col,
                                          self._text_formats['text_smallest'], '1 in ',
                                          self._text_formats['text_regular'], triggered_in_board,
                                          cur_format)

            cur_col += up_trigger_in_board_width

            self._sheet.merge_range(self._current_row, cur_col,
                                    self._current_row, cur_col + up_trigger_in_feature_width-1,
                                    '', cur_format)
            triggered_in_feature = '{:,.3f}'.format(upgrade_df['1_in_triggered_feature'].loc[upgrade])
            self._sheet.write_rich_string(self._current_row, cur_col,
                                          self._text_formats['text_smallest'], '1 in ',
                                          self._text_formats['text_regular'], triggered_in_feature,
                                          cur_format)

            cur_col += up_trigger_in_feature_width

            self._sheet.merge_range(self._current_row, cur_col,
                                    self._current_row, cur_col + percent_width-1,
                                    '', cur_format)
            triggered_in_feature = '{:,.3f}'.format(upgrade_df['percent_of_trigger_when_in_board'].loc[upgrade])
            self._sheet.write_rich_string(self._current_row, cur_col,
                                          self._text_formats['text_regular'], triggered_in_feature,
                                          self._text_formats['text_regular'], '%',
                                          cur_format)

            self._current_row += 1
        self._current_row += 2

    def _WriteSLPotFreq(self):
        name_width = 4
        pulls_to_hit_feature = 3
        pulls_to_hit_base = 3
        start_col = self.GetCentreStartCol(self.GetSmallHeaderWidth(),
                                           name_width + pulls_to_hit_base + pulls_to_hit_feature,
                                           self.GetSmallHeaderStartColumn())

        color_index = 3

        cur_col = start_col
        self._sheet.set_row_pixels(self._current_row, self.GetBigCellPixelHeight())
        self._sheet.merge_range(self._current_row, cur_col,
                                self._current_row, cur_col + name_width - 1,
                                '', self._rtp_distribution_formats['border_description_'+str(color_index)])
        self._sheet.write_rich_string(self._current_row, cur_col,
                                      self._text_formats['text_big_bold'], 'Pot Name',
                                      self._text_formats['text_big_bold'], ' ',
                                      self._rtp_distribution_formats['border_description_'+str(color_index)])

        cur_col += name_width

        self._sheet.merge_range(self._current_row, cur_col,
                                self._current_row, cur_col + pulls_to_hit_feature - 1,
                                '', self._rtp_distribution_formats['border_description_' + str(color_index)])
        self._sheet.write_rich_string(self._current_row, cur_col,
                                      self._text_formats['text_big_bold'], 'Pulls to Hit',
                                      self._text_formats['text_smallest'], '\n(1 in ... SL features)',
                                      self._rtp_distribution_formats['border_description_' + str(color_index)])
        cur_col += pulls_to_hit_feature

        self._sheet.merge_range(self._current_row, cur_col,
                                self._current_row, cur_col + pulls_to_hit_base - 1,
                                '', self._rtp_distribution_formats['border_description_' + str(color_index)])
        self._sheet.write_rich_string(self._current_row, cur_col,
                                      self._text_formats['text_big_bold'], 'Pulls to Hit',
                                      self._text_formats['text_smallest'], '\n(1 in ... base spins)',
                                      self._rtp_distribution_formats['border_description_' + str(color_index)])

        self._current_row += 1

        pots_df = self._stats.GetSLPotsFreq_DF()
        for i, pot in enumerate(pots_df.index):
            cur_format = self._rtp_distribution_formats['border_even_'+str(color_index)] if i % 2 == 0 else \
                self._rtp_distribution_formats['border_odd_' + str(color_index)]
            self._sheet.set_row_pixels(self._current_row, self.GetNormalCellPixelHeight())

            cur_col = start_col
            self._sheet.merge_range(self._current_row, cur_col,
                                    self._current_row, cur_col+name_width-1,
                                    '', cur_format)
            self._sheet.write_rich_string(self._current_row, cur_col,
                                          self._text_formats['text_regular_bold'], pots_df['pot_name'].loc[pot] + ' x' + str(pots_df['multiplier'].loc[pot]),
                                          self._text_formats['text_regular_bold'], ' ',
                                          cur_format)
            cur_col += name_width
            self._sheet.merge_range(self._current_row, cur_col,
                                    self._current_row, cur_col + pulls_to_hit_feature - 1,
                                    '', cur_format)
            val_sl = '{:.3f}'.format(pots_df['pulls_to_hit_sl'].loc[pot])
            self._sheet.write_rich_string(self._current_row, cur_col,
                                          self._text_formats['text_smallest'], '1 in ',
                                          self._text_formats['text_regular'], val_sl,
                                          cur_format)
            cur_col += pulls_to_hit_feature
            self._sheet.merge_range(self._current_row, cur_col,
                                    self._current_row, cur_col + pulls_to_hit_base - 1,
                                    '', cur_format)
            val = '{:.3f}'.format(pots_df['pulls_to_hit_base'].loc[pot])
            self._sheet.write_rich_string(self._current_row, cur_col,
                                          self._text_formats['text_smallest'], '1 in ',
                                          self._text_formats['text_regular'], val,
                                          cur_format)
            self._current_row += 1
        self._current_row += 3


    def _WriteSLScattersCountStats(self):
        self.WriteInnerHeader("SL Scatters Count", include_in_content=True, print_counter=True)

        scatter_count_main_df, scatter_count_merged_df = self._stats.GetSLScatterCounter_DFs()
        scatter_count_labels = {'avg_variant_with_zero': 'Average number of scatters:',
                                'avg_variant_no_zero': 'Average number of scatters:',
                                'any_variant_1_in_base_with_zero': 'Pulls to hit any scatters:',
                                'any_variant_1_in_base_no_zero': 'Pulls to hit any scatters:',
                                'any_variant_1_in_feature_no_zero': 'Pulls to hit any scatters:',
                                'plus_variants_1_in_base': 'Pulls to hit {:d}+ scatters:',
                                'plus_variants_1_in_feature': 'Pulls to hit {:d}+ scatters:',
                                'picture_col': 'Number of Scatters in View Distribution',
                                'variant_col': 'Scatters Count',
                                'pulls_to_hit_big': 'Pulls to Hit',
                                'pulls_to_hit_big_merged': 'Merged Ranges',
                                'pulls_to_hit_small': 'Pulls to Hit',
                                'pulls_to_hit_small_merged': 'Merged Ranges',
                                'percent_big': 'Percent',
                                'percent_big_merged': 'Merged Ranges',
                                'percent_small': 'Percent',
                                'percent_small_merged': 'Merged Ranges',
                                'graph_name': 'sl_scatter_count'
                                }
        self.WriteBasicVariant(main_df=scatter_count_main_df,
                               merged_df=scatter_count_merged_df,
                               variant_df_col_name='count',
                               txt_labels=scatter_count_labels,
                               variant_is_mult=False,
                               color_index=0,
                               n=(4, 7, 10),
                               include_zero_in_total=False,
                               print_n_plus_1_in_base=True,
                               print_n_plus_1_in_feature=False,
                               print_avg_variant_with_zero=False,
                               print_avg_variant_no_zero=True,
                               print_any_variant_1_in_base_no_zero=True,
                               print_any_variant_1_in_base_with_zero=False,
                               print_any_variant_1_in_feature_no_zero=False,
                               print_pulls_to_hit_big_col=False,
                               print_pulls_to_hit_small_col=True,
                               print_percent_big_col=False,
                               print_percent_small_col=True)

    def _WriteSLRespinCountStats(self):
        self.WriteInnerHeader("SL Re-Spin Count", include_in_content=True, print_counter=True)

        respin_count_main_df, respin_count_merged_df = self._stats.GetSLRespinCounter_DFs()
        respin_count_labels = {'avg_variant_with_zero': 'Average number of Re-spins:',
                               'avg_variant_no_zero': 'Average number of Re-spins:',
                               'any_variant_1_in_base_with_zero': 'Pulls to hit any Re-spins:',
                               'any_variant_1_in_base_no_zero': 'Pulls to hit any Re-spins:',
                               'any_variant_1_in_feature_no_zero': 'Pulls to hit any Re-spins:',
                               'plus_variants_1_in_base': 'Pulls to hit {:d}+ Re-spins:',
                               'plus_variants_1_in_feature': 'Pulls to hit {:d}+ Re-spins:',
                               'picture_col': 'Number of Re-spins Distribution',
                               'variant_col': 'Re-spins Count',
                               'pulls_to_hit_big': 'Pulls to Hit',
                               'pulls_to_hit_big_merged': 'Merged Ranges',
                               'pulls_to_hit_small': 'Pulls to Hit',
                               'pulls_to_hit_small_merged': 'Merged Ranges',
                               'percent_big': 'Percent',
                               'percent_big_merged': 'Merged Ranges',
                               'percent_small': 'Percent',
                               'percent_small_merged': 'Merged Ranges',
                               'graph_name': 'sl_respin_count'
                               }
        self.WriteBasicVariant(main_df=respin_count_main_df,
                               merged_df=respin_count_merged_df,
                               variant_df_col_name='count',
                               txt_labels=respin_count_labels,
                               variant_is_mult=False,
                               color_index=1,
                               n=(10, 15, 20, 30),
                               include_zero_in_total=False,
                               print_n_plus_1_in_base=False,
                               print_n_plus_1_in_feature=True,
                               print_avg_variant_with_zero=False,
                               print_avg_variant_no_zero=True,
                               print_any_variant_1_in_base_no_zero=False,
                               print_any_variant_1_in_base_with_zero=False,
                               print_any_variant_1_in_feature_no_zero=False,
                               print_pulls_to_hit_big_col=True,
                               print_pulls_to_hit_small_col=False,
                               print_percent_big_col=True,
                               print_percent_small_col=False)

    def _WriteSLFilledCellsStats(self):
        self.WriteInnerHeader("SL Number of Filled Cells Distribution", include_in_content=True, print_counter=True)

        filled_cells_count_main_df, filled_cells_count_merged_df = self._stats.GetSLFilledCells_DFs()
        filled_cells_count_labels = {'avg_variant_with_zero': 'Average number of Filled Cells:',
                                     'avg_variant_no_zero': 'Average number of Filled Cells:',
                                     'any_variant_1_in_base_with_zero': 'Pulls to hit any Filled Cells:',
                                     'any_variant_1_in_base_no_zero': 'Pulls to hit any Filled Cells:',
                                     'any_variant_1_in_feature_no_zero': 'Pulls to hit any Filled Cells:',
                                     'plus_variants_1_in_base': 'Pulls to hit {:d}+ Filled Cells:',
                                     'plus_variants_1_in_feature': 'Pulls to hit {:d}+ Filled Cells:',
                                     'picture_col': 'Number of Filled Cells Distribution',
                                     'variant_col': 'Filled Cells Count',
                                     'pulls_to_hit_big': 'Pulls to Hit',
                                     'pulls_to_hit_big_merged': 'Merged Ranges',
                                     'pulls_to_hit_small': 'Pulls to Hit',
                                     'pulls_to_hit_small_merged': 'Merged Ranges',
                                     'percent_big': 'Percent',
                                     'percent_big_merged': 'Merged Ranges',
                                     'percent_small': 'Percent',
                                     'percent_small_merged': 'Merged Ranges',
                                     'graph_name': 'sl_filled_cells_count'
                                     }
        self.WriteBasicVariant(main_df=filled_cells_count_main_df,
                               merged_df=filled_cells_count_merged_df,
                               variant_df_col_name='count',
                               txt_labels=filled_cells_count_labels,
                               variant_is_mult=False,
                               color_index=3,
                               n=(10, 15, 20, 25, 30),
                               include_zero_in_total=False,
                               print_n_plus_1_in_base=False,
                               print_n_plus_1_in_feature=True,
                               print_avg_variant_with_zero=False,
                               print_avg_variant_no_zero=True,
                               print_any_variant_1_in_base_no_zero=False,
                               print_any_variant_1_in_base_with_zero=False,
                               print_any_variant_1_in_feature_no_zero=False,
                               print_pulls_to_hit_big_col=True,
                               print_pulls_to_hit_small_col=False,
                               print_percent_big_col=True,
                               print_percent_small_col=False)

    def _WriteSLScattersWinStats(self):
        self.WriteInnerHeader("SL Scatters Win Distribution", print_counter=True, include_in_content=True)

        self.WriteBasicSpinWin(spin_win_df=self._stats.GetSLScattersWin_DF(),
                               graph_name='sl_scatters_win',
                               color_index=4,
                               print_avg_win_no_zero=True,
                               print_avg_win_with_zero=False,
                               print_win_1_in_feature=False,
                               print_win_1_in_base=False,
                               print_win_freq_base_col=True,
                               print_max_range_win_col=True)

    def _WriteSLWinStats(self):
        self.WriteInnerHeader("SL Total Win Distribution", print_counter=True, include_in_content=True)

        self.WriteBasicSpinWin(spin_win_df=self._stats.GetSLSpinWin_DF(),
                               graph_name='sl_total_win',
                               color_index=5,
                               print_avg_win_no_zero=True,
                               print_avg_win_with_zero=False,
                               print_win_1_in_feature=True,
                               print_win_1_in_base=True,
                               print_win_freq_base_col=True,
                               print_max_range_win_col=True)

    def _WriteSLFilledCellsOneSpin(self):
        self.WriteInnerHeader("SL Number of Filled Cells on One Re-Spin", include_in_content=True, print_counter=True)

        filled_cells_count_main_df, filled_cells_count_merged_df = self._stats.GetSLFilledCellsOneSpin_DFs()
        filled_cells_count_labels = {'avg_variant_with_zero': 'Average number of Filled Cells:',
                                     'avg_variant_no_zero': 'Average number of Filled Cells:',
                                     'any_variant_1_in_base_with_zero': 'Pulls to hit any Filled Cells:',
                                     'any_variant_1_in_base_no_zero': 'Pulls to hit any Filled Cells:',
                                     'any_variant_1_in_feature_no_zero': 'Pulls to hit any Filled Cells:',
                                     'plus_variants_1_in_base': 'Pulls to hit {:d}+ Filled Cells:',
                                     'plus_variants_1_in_feature': 'Pulls to hit {:d}+ Filled Cells:',
                                     'picture_col': 'Number of Filled Cells Distribution',
                                     'variant_col': 'Filled Cells Count',
                                     'pulls_to_hit_big': 'Pulls to Hit',
                                     'pulls_to_hit_big_merged': 'Merged Ranges',
                                     'pulls_to_hit_small': 'Pulls to Hit',
                                     'pulls_to_hit_small_merged': 'Merged Ranges',
                                     'percent_big': 'Percent',
                                     'percent_big_merged': 'Merged Ranges',
                                     'percent_small': 'Percent',
                                     'percent_small_merged': 'Merged Ranges',
                                     'graph_name': 'sl_filled_cells_one_respin'
                                     }
        self.WriteBasicVariant(main_df=filled_cells_count_main_df,
                               merged_df=filled_cells_count_merged_df,
                               variant_df_col_name='count',
                               txt_labels=filled_cells_count_labels,
                               variant_is_mult=False,
                               color_index=0,
                               n=(2, 3, 4, 5, 6),
                               include_zero_in_total=False,
                               print_n_plus_1_in_base=False,
                               print_n_plus_1_in_feature=True,
                               print_avg_variant_with_zero=False,
                               print_avg_variant_no_zero=True,
                               print_any_variant_1_in_base_no_zero=False,
                               print_any_variant_1_in_base_with_zero=False,
                               print_any_variant_1_in_feature_no_zero=False,
                               print_pulls_to_hit_big_col=True,
                               print_pulls_to_hit_small_col=False,
                               print_percent_big_col=True,
                               print_percent_small_col=False)
