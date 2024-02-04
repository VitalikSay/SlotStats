import xlsxwriter as xl
import json as j
import numpy as np
import pandas as pd
import xlsxwriter.exceptions
from collections import defaultdict
from xlsxwriter.utility import xl_rowcol_to_cell
from Basic_Code.Stats_Writer.BasicGraphMaker import Colors

from Basic_Code.Basic_Structures.BasicSlot import BasicSlot
from Basic_Code.Basic_Structures.BasicStatistics import BasicStatistics
from Basic_Code.Basic_Calculator.BasicStatsCalculator import BasicStatsCalculator
from Basic_Code.Stats_Writer.BasicGraphMaker import BasicGraphMaker
from Basic_Code.Utils.BasicTimer import timer
from Basic_Code.Stats_Writer.BasicPARFormats import BasicPARFormats
from Basic_Code.Basic_Calculator.BasicSpinWin import BasicSpinWin


class BasicPARSheet:
    def __init__(self, book: xl.Workbook, formats: BasicPARFormats, slot: BasicSlot, stats: BasicStatistics, calculation: BasicStatsCalculator):
        self._graphMaker = self._MakeGraphMaker(slot, stats, calculation)

        self._book = book
        self._par_formats = formats
        self._sheet = xl.worksheet
        self._sheet_name = slot.GetGameName() + ' ' + slot.GetVersion()

        self._game_short_name = slot.GetGameName()
        self._game_full_name = slot.GetGameName(False)
        self._game_version_name = slot.GetVersion()
        self._line_wins_txt = slot.GetLineWins()
        self._features_list = slot.GetFeaturesList()
        self._board_shape = [slot.GetBoardHeight(), slot.GetBoardWidth()]
        self._std = calculation.GetStdInPercent()
        self._volatility = calculation.GetVolatility()
        self._spin_count = stats.GetTotalSpinCount()
        self._max_win_in_cents = calculation.GetMaxWinInCents()
        self._max_win_count = calculation.GetMaxWinCount()
        self._simulation_id = slot.GetSimulationId()
        self._feature_respin_strategy = slot.GetFeatureRespinStrategy()
        self._feature_respin_mults = slot.GetFeatureRespinMults()

        self._winlines = slot.GetWinlines()
        self._paytable = slot.GetPaytable()
        self._reelsets = slot.GetReelsets()

        self._rtp = calculation.GetRTP()
        self._win_frequency = calculation.GetHitFrPerc()
        self._total_spin_win_df = calculation.GetTotalSpinWinDF()
        self._base_bet = slot.GetBet()
        self._mult_bet = slot.GetBet(False)

        self._current_row = 0
        self._current_header = -1
        self._current_inner_header = -1
        self._headers_content = []
        self._inner_headers_content = defaultdict(list)
        self._reelsets_spin_win_section_content = []
        self._reelsets_spin_win_reelsets_content = defaultdict(list)
        self._reelsets_structure_section_content = []
        self._reelsets_structure_reelsets_content = defaultdict(list)
        self._allocated_rows_for_content = 35

        self._global_content_row_col = [-1, -1]
        self._reelsets_spin_win_content_row_col = [-1, -1]
        self._reelsets_structure_content_row_col = [-1, -1]

        self._total_spin_win_df = calculation.GetTotalSpinWinDF()
        self._section_feature_spin_win_dfs = calculation.GetSectionFeatureSpinWinsDF()
        self._section_all_spin_wins = calculation.GetSectionAllSpinWinsDF()
        self._section_names = slot.GetSectionNames()
        self._line_wins_dfs = calculation.GetLineWinsDF()
        self._confidence_percents = calculation.GetConfidencePercents()
        self._total_confidence_dfs = calculation.GetTotalConfIntervalsDF()
        self._sections_confidence_dfs = calculation.GetSectionConfIntervalsDF()
        self._total_spin_win_struct = calculation.GetTotalSpinWinStruct()
        self._sections_spin_win_structs = calculation.GetSectionSpinWinStructs()
        self._top_award_df = calculation.GetTopAwardDF()
        self._all_reelsets_section_dfs = calculation.GetAllReelsetsSectionDF()
        self._base_reelsets_df = calculation.GetBaseReelsetsDF()
        self._base_reelsets_spin_win_dfs = calculation.GetAllWinByBaseReelsetDF()
        self._all_reelsets_spin_win_dfs = calculation.GetReelsetsSpinWinsDF()

        self._header_width = 14
        self._header_start_col = 0
        self._inner_header_width = 12
        self._inner_header_start_col = 1
        self._small_header_width = 10
        self._small_header_start_col = 2

        self._small_cell_pixel_height = 20
        self._normal_cell_pixel_height = 30
        self._big_cell_pixel_height = 45

        self._text_formats = dict()
        self._summary_formats = dict()
        self._paytable_formats = dict()
        self._header_formats = dict()
        self._text_formats = dict()
        self._winline_formats = dict()
        self._rtp_distribution_formats = dict()
        self._line_wins_formats = dict()
        self._confidence_formats = dict()
        self._top_award_formats = dict()
        self._reelsets_table_formats = dict()
        self._reelsets_spin_win_formats = dict()
        self._reelsets_structure_formats = dict()
        self._content_formats = dict()
        self._variant_formats = dict()
        self._info_row_formats = dict()

    def _MakeGraphMaker(self, slot: BasicSlot, stats: BasicStatistics, calculation: BasicStatsCalculator):
        return BasicGraphMaker(slot, stats, calculation)

    @timer
    def WritePARSheet(self):
        self._InitSheet()

        self._WriteGameSummary()
        self._AllocateSpaceGlobalContent()
        self._WritePayTable()
        self._WriteWinlines()
        self._WriteRTPDistributions()
        self._WriteWinlinesWins()
        self._WriteConfidenceIntervals()
        self._WriteWinAmountStatistics()
        self.WriteCustomGameStatistics()
        self._WriteReelsetsPivotTables()
        self._WriteReelsetsSpinWinTables()
        self._WriteReelsetsStructure()
        self._WriteContent()

    def _InitSheet(self):
        self._sheet = self._book.add_worksheet(self._sheet_name)
        self._sheet.ignore_errors({'two_digit_text_year': 'A1:XFD1048576'})
        self._sheet.ignore_errors({'number_stored_as_text': 'A1:XFD1048576'})
        self._graphMaker.MakePARPlots()
        self._InitFormats()

    def _InitFormats(self):
        self._summary_formats = self._par_formats.GetSummaryFormats()
        self._paytable_formats = self._par_formats.GetPaytableFormats()
        self._header_formats = self._par_formats.GetHeaderFormats()
        self._text_formats = self._par_formats.GetFontFormats()
        self._winline_formats = self._par_formats.GetWinlineFormats()
        self._rtp_distribution_formats = self._par_formats.GetRTPDistributionFormats()
        self._line_wins_formats = self._par_formats.GetLineWinsFormats()
        self._confidence_formats = self._par_formats.GetCondideneceFormats()
        self._top_award_formats = self._par_formats.GetTopAwardFormats()
        self._reelsets_table_formats = self._par_formats.GetReelsetsTableFormats()
        self._reelsets_spin_win_formats = self._par_formats.GetReelsetsSpinWinFormats()
        self._reelsets_structure_formats = self._par_formats.GetReelsetsStructureFormats()
        self._content_formats = self._par_formats.GetContentFormats()
        self._variant_formats = self._par_formats.GetVariantFormats()
        self._info_row_formats = self._par_formats.GetInfoRowFormats()

    def WriteHeader(self, text: str, print_counter=True, include_in_content=True):
        self._current_inner_header = -1

        start_col = self._header_start_col
        width = self._header_width
        button_width = 2

        if print_counter:
            self._current_header += 1
            text = str(self._current_header+1) + ". " + text
        if include_in_content:
            self._headers_content.append((text, self._current_row))
        self._sheet.set_row_pixels(self._current_row, 90)
        self._sheet.merge_range(self._current_row, start_col,
                                self._current_row, start_col+width - 1,
                                text, self._header_formats["header"])
        if print_counter and include_in_content and self._current_header > 0:
            self._sheet.merge_range(self._current_row, start_col + width,
                                    self._current_row, start_col + width + button_width - 1,
                                    '', self._content_formats['content_button_big'])
            link_to = "internal:'" + self._sheet_name + "'!" + xl_rowcol_to_cell(self._global_content_row_col[0], 1)
            self._sheet.write_url(self._current_row, start_col+width,
                                  link_to,
                                  string='Go To\nContent',
                                  tip='Click Here for Quick Jump',
                                  cell_format=self._content_formats['content_button_big'])
        self._current_row += 2

    def WriteInnerHeader(self, text, print_counter=True, include_in_content=True):
        start_col = self._inner_header_start_col
        width = self._inner_header_width
        button_width = 2

        if print_counter:
            self._current_inner_header += 1
            text = str(self._current_header+1) + "." + str(self._current_inner_header+1) + ". " + text
        if include_in_content:
            self._inner_headers_content[self._current_header].append((text, self._current_row))
        self._sheet.set_row_pixels(self._current_row, 60)
        self._sheet.merge_range(self._current_row, start_col,
                                self._current_row, start_col + width - 1,
                                text, self._header_formats["inner_header"])
        if self._current_inner_header > 0:
            self._sheet.merge_range(self._current_row, start_col+width,
                                    self._current_row, start_col+width+button_width-1,
                                    '', self._content_formats['content_button_small'])
            link_to = "internal:'" + self._sheet_name + "'!" + xl_rowcol_to_cell(self._global_content_row_col[0], 1)
            self._sheet.write_url(self._current_row, start_col+width,
                                  link_to,
                                  string='Go To\nContent',
                                  tip='Click Here for Quick Jump',
                                  cell_format=self._content_formats['content_button_small'])
        self._current_row += 2

    def WriteSmallHeader(self, text):
        self._sheet.set_row_pixels(self._current_row, self._big_cell_pixel_height)
        self._sheet.merge_range(self._current_row, self._small_header_start_col,
                                self._current_row, self._small_header_start_col + self._small_header_width - 1,
                                text, self._header_formats['small_header'])
        self._current_row += 2

    def _WriteGameSummary(self):
        self._current_row += 1
        self.WriteHeader("Game Summary", print_counter=False, include_in_content=False)

        start_col = 1
        game_name_width = 5
        board_size_width = 3
        std_width = 3
        line_wins_width = 3
        rtp_width = 4
        frequency_width = 4
        game_features_width = 5
        rtp_distribution_width = 6

        self._sheet.merge_range(self._current_row, start_col, self._current_row, start_col+game_name_width-1, "",
                                self._summary_formats['border_cell_description'])
        self._sheet.write_rich_string(self._current_row, start_col,
                                      self._text_formats['text_regular'], 'Game',
                                      self._text_formats['text_regular'], ' Name',
                                      self._summary_formats['border_cell_description'])
        self._sheet.merge_range(self._current_row, start_col+game_name_width,
                                self._current_row, start_col+game_name_width+board_size_width-1, "",
                                self._summary_formats['border_cell_description'])
        self._sheet.write_rich_string(self._current_row, start_col+game_name_width,
                                      self._text_formats['text_regular'], 'Board',
                                      self._text_formats['text_regular'], ' Size',
                                      self._summary_formats['border_cell_description'])
        self._sheet.merge_range(self._current_row, start_col+game_name_width+board_size_width,
                                self._current_row, start_col+game_name_width+board_size_width+std_width-1,
                                "", self._summary_formats['border_cell_description'])
        self._sheet.write_rich_string(self._current_row, start_col+game_name_width+board_size_width,
                                      self._text_formats['text_regular'], 'S',
                                      self._text_formats['text_regular'], 'TD',
                                      self._summary_formats['border_cell_description'])

        self._current_row += 1

        self._sheet.merge_range(self._current_row, start_col, self._current_row + 1, start_col+game_name_width-1,
                                "", self._summary_formats['border_under_description'])
        self._sheet.write_rich_string(self._current_row, start_col,
                                      self._text_formats['text_biggest_bold'], self._game_full_name[:5],
                                      self._text_formats['text_biggest_bold'], self._game_full_name[5:] + ' (' + self._game_short_name + ')',
                                      self._summary_formats['border_under_description'])
        self._sheet.merge_range(self._current_row, start_col+game_name_width,
                                self._current_row + 1, start_col+game_name_width+board_size_width-1,
                                "", self._summary_formats['border_under_description'])
        self._sheet.write_rich_string(self._current_row, start_col+game_name_width,
                                      self._text_formats['text_regular'], str(self._board_shape[0]) + " rows\n",
                                      self._text_formats['text_regular'], str(self._board_shape[1]) + " reels",
                                      self._summary_formats['border_under_description'])
        self._sheet.merge_range(self._current_row, start_col+game_name_width+board_size_width,
                                self._current_row + 1, start_col+game_name_width+board_size_width+std_width-1,
                                "", self._summary_formats['border_under_description'])
        self._sheet.write_rich_string(self._current_row, start_col+game_name_width+board_size_width,
                                      self._text_formats['text_big'], str(round(self._std, 3)) + '\n',
                                      self._text_formats['text_small'], self._volatility,
                                      self._summary_formats['border_under_description'])

        self._current_row += 2

        self._sheet.merge_range(self._current_row, start_col, self._current_row, start_col+line_wins_width-1,
                                "", self._summary_formats['border_cell_description'])
        self._sheet.write_rich_string(self._current_row, start_col,
                                      self._text_formats['text_regular'], 'Line ',
                                      self._text_formats['text_regular'], 'Wins',
                                      self._summary_formats['border_cell_description'])
        self._sheet.merge_range(self._current_row, start_col+line_wins_width,
                                self._current_row + 5, start_col+line_wins_width+rtp_width-1,
                                "", self._summary_formats['border_rtp_value'])
        self._sheet.write_rich_string(self._current_row, start_col+line_wins_width,
                                      self._text_formats['text_biggest_bold'], "RTP\n",
                                      self._text_formats['text_biggest'], str(round(self._rtp * 100, 4)) + "%",
                                      self._summary_formats['border_rtp_value'])
        self._sheet.merge_range(self._current_row, start_col+line_wins_width+rtp_width,
                                self._current_row + 5, start_col+line_wins_width+rtp_width+frequency_width-1,
                                "", self._summary_formats['border_rtp_value'])
        self._sheet.write_rich_string(self._current_row, start_col+line_wins_width+rtp_width,
                                      self._text_formats['text_biggest_bold'], "Win Frequency\n",
                                      self._text_formats['text_biggest'], '1 in ' + str(round(100 / self._win_frequency, 4)),
                                      self._text_formats['text_regular'], '\n(',
                                      self._text_formats['text_regular'], str(round(self._win_frequency, 4)),
                                      self._text_formats['text_regular'], '%)',
                                      self._summary_formats['border_rtp_value'])

        self._current_row += 1

        self._sheet.merge_range(self._current_row, start_col,
                                self._current_row + 1, start_col+line_wins_width-1,
                                "", self._summary_formats['border_under_description'])
        self._sheet.write_rich_string(self._current_row, start_col,
                                      self._text_formats['text_regular'], self._line_wins_txt[:1],
                                      self._text_formats['text_regular'], self._line_wins_txt[1:],
                                      self._summary_formats['border_under_description'])

        self._current_row += 2

        self._sheet.merge_range(self._current_row, start_col,
                                self._current_row + 2, start_col+line_wins_width-1,
                                "", self._summary_formats['border_rtp_value'])
        self._sheet.write_rich_string(self._current_row, start_col,
                                      self._text_formats['text_big_bold'], str(len(self._winlines)) + '\n',
                                      self._text_formats['text_regular'], 'winlines',
                                      self._summary_formats['border_rtp_value'])

        self._current_row += 3

        self._sheet.merge_range(self._current_row, start_col,
                                self._current_row, start_col+game_features_width-1,
                                "", self._summary_formats['border_cell_description'])
        self._sheet.write_rich_string(self._current_row, 1,
                                      self._text_formats['text_regular'], 'Game ',
                                      self._text_formats['text_regular'], 'Features',
                                      self._summary_formats['border_cell_description'])
        self._sheet.merge_range(self._current_row, start_col+game_features_width,
                                self._current_row, start_col+game_features_width+rtp_distribution_width-1,
                                "", self._summary_formats['border_cell_description'])
        self._sheet.write_rich_string(self._current_row, start_col+game_features_width,
                                      self._text_formats['text_regular'], 'RTP ',
                                      self._text_formats['text_regular'], 'distribution by game sections',
                                      self._summary_formats['border_cell_description'])

        self._current_row += 1

        self._sheet.merge_range(self._current_row, start_col,
                                self._current_row + 11, start_col+game_features_width-1,
                                "", self._summary_formats['border_under_description_feature'])
        features_list = [str(i+1) + '. ' + feature_name + '\n' for i, feature_name in enumerate(self._features_list.split(', '))]
        formats_list = [self._text_formats['text_regular'] for _ in range(len(features_list))]
        tuple_segments = zip(formats_list, features_list)
        segments = []
        for form, feature in tuple_segments:
            segments.append(form)
            segments.append(feature)
        self._sheet.write_rich_string(self._current_row, start_col,
                                      *segments,
                                      self._summary_formats['border_under_description_feature'])
        self._sheet.insert_image(self._current_row, start_col+game_features_width,
                                 self._graphMaker.GetSummaryPiePlotPath(),
                                 {
                                      'x_offset': 1,
                                      'y_offset': 1,
                                      'x_scale': 0.562,
                                      'y_scale': 0.586,
                                      'object_position': 2
                                  })

        self._current_row += 12

        self._sheet.merge_range(self._current_row, start_col,
                                self._current_row, start_col+game_features_width-1,
                                "", self._summary_formats['border_cell_description'])
        self._sheet.write_rich_string(self._current_row, start_col,
                                      self._text_formats['text_regular'], 'Statistics calculated ',
                                      self._text_formats['text_regular'], 'based on',
                                      self._summary_formats['border_cell_description'])

        self._current_row += 1

        self._sheet.merge_range(self._current_row, start_col,
                                self._current_row, start_col + game_features_width-1,
                                "", self._summary_formats['border_under_description'])
        self._sheet.write_rich_string(self._current_row, start_col,
                                      self._text_formats['text_regular'], '{:,d}'.format(self._spin_count),
                                      self._text_formats['text_regular'], ' spins',
                                      self._summary_formats['border_under_description'])

        self._current_row += 1

        self._sheet.merge_range(self._current_row, start_col,
                                self._current_row, start_col+game_features_width-1,
                                "", self._summary_formats['border_cell_description'])
        self._sheet.write_rich_string(self._current_row, start_col,
                                      self._text_formats['text_regular'], 'Simulation ',
                                      self._text_formats['text_regular'], 'ID',
                                      self._summary_formats['border_cell_description'])

        self._current_row += 1

        self._sheet.merge_range(self._current_row, start_col,
                                self._current_row+1, start_col+game_features_width-1,
                                "", self._summary_formats['border_under_description'])
        self._sheet.write_rich_string(self._current_row, start_col,
                                      self._text_formats['text_regular'], self._simulation_id[:1],
                                      self._text_formats['text_regular'], self._simulation_id[1:],
                                      self._summary_formats['border_under_description'])

        self._current_row += 5

    def _AllocateSpaceGlobalContent(self):
        self.WriteHeader("Content", print_counter=False, include_in_content=False)
        self._global_content_row_col = [self._current_row, 1]
        self._current_row += self._allocated_rows_for_content + 3

    def _WriteContent(self):
        content_row = self._global_content_row_col[0]
        big_start = 1
        small_start = 2
        big_width = 12
        small_width = 11

        inner_break = False
        for big_index, (big_text, big_row) in enumerate(self._headers_content):
            self._sheet.set_row_pixels(content_row, 40)

            try:
                self._sheet.merge_range(content_row, big_start,
                                        content_row, big_start + big_width - 1,
                                        '', self._content_formats['big'])
            except xlsxwriter.exceptions.OverlappingRange:
                content_row -= 1
                self._sheet.merge_range(content_row, 15,
                                        content_row, 25,
                                        'WARNING !!! NOT ENOUGH SPACE ALLOCATED TO INSERT CONTENT !!!',
                                        self._text_formats['text_red_info'])
                break
            link_to = "internal:'" + self._sheet_name + "'!" + xl_rowcol_to_cell(big_row, 1)
            self._sheet.write_url(content_row, big_start, link_to, string=big_text,
                                  cell_format=self._content_formats['big'], tip='Click Here for Quick Jump')
            content_row += 1
            for small_index, (small_text, small_row) in enumerate(self._inner_headers_content[big_index]):
                self._sheet.set_row_pixels(content_row, 30)
                try:
                    self._sheet.merge_range(content_row, small_start,
                                            content_row, small_start + small_width - 1,
                                            '', self._content_formats['small'])
                except xlsxwriter.exceptions.OverlappingRange:
                    content_row -= 1
                    self._sheet.merge_range(content_row, 15,
                                            content_row, 25,
                                            'WARNING !!! NOT ENOUGH SPACE ALLOCATED TO INSERT CONTENT !!!',
                                            self._text_formats['text_red_info'])
                    inner_break = True
                    break
                link_to = "internal:'" + self._sheet_name + "'!" + xl_rowcol_to_cell(small_row, 1)
                self._sheet.write_url(content_row, small_start, link_to, string=small_text,
                                      cell_format=self._content_formats['small'], tip='Click Here for Quick Jump')
                content_row += 1
            if inner_break:
                break

    def _WritePayTable(self):
        self.WriteHeader("Paytable", print_counter=True, include_in_content=True)

        paying_combinations = self._paytable.GetPayingCombinations()
        start_col = 1
        symbol_id_width = 2
        symbol_name_width = 3
        one_symbol_pay_width = 2
        total_symbol_pay_width = one_symbol_pay_width * len(paying_combinations)

        self._sheet.merge_range(self._current_row, start_col,
                                self._current_row+1, start_col+symbol_id_width-1,
                                "", self._paytable_formats['border_description'])
        self._sheet.write_rich_string(self._current_row, start_col,
                                      self._text_formats['text_big_bold'], 'Symbol ',
                                      self._text_formats['text_big_bold'], 'ID',
                                      self._paytable_formats['border_description'])
        self._sheet.merge_range(self._current_row, start_col+symbol_id_width,
                                self._current_row+1, start_col+symbol_id_width+symbol_name_width-1,
                                "", self._paytable_formats['border_description'])
        self._sheet.write_rich_string(self._current_row, start_col+symbol_id_width,
                                      self._text_formats['text_big_bold'], 'Symbol ',
                                      self._text_formats['text_big_bold'], 'Name',
                                      self._paytable_formats['border_description'])
        self._sheet.merge_range(self._current_row, start_col+symbol_id_width+symbol_name_width,
                                self._current_row, start_col+symbol_id_width+symbol_name_width+total_symbol_pay_width-1,
                                "", self._paytable_formats['border_description'])
        self._sheet.write_rich_string(self._current_row, start_col+symbol_id_width+symbol_name_width,
                                      self._text_formats['text_regular_bold'], 'Winline ',
                                      self._text_formats['text_regular_bold'], 'Pay',
                                      self._paytable_formats['border_description'])

        self._current_row += 1

        for index, index_pay in enumerate(paying_combinations):
            self._sheet.merge_range(self._current_row, start_col+symbol_id_width+symbol_name_width+index*one_symbol_pay_width,
                                    self._current_row, start_col+symbol_id_width+symbol_name_width+((index+1)*one_symbol_pay_width-1),
                                    "", self._paytable_formats['border_description'])
            self._sheet.write_rich_string(self._current_row, start_col+symbol_id_width+symbol_name_width+index*one_symbol_pay_width,
                                          self._text_formats['text_regular_bold'], str(index_pay+1),
                                          self._text_formats['text_small'], ' symbols',
                                          self._paytable_formats['border_description'])

        self._current_row += 1

        for index in range(self._paytable.GetSymbolCount()):
            self._sheet.set_row_pixels(self._current_row, 30)
            current_border_format = self._paytable_formats['border_even'] if index % 2 == 0 else self._paytable_formats['border_odd']
            self._sheet.merge_range(self._current_row, start_col,
                                    self._current_row, start_col+symbol_id_width-1,
                                    "", current_border_format)
            self._sheet.write_rich_string(self._current_row, start_col,
                                          self._text_formats['text_regular'], str(index),
                                          self._text_formats['text_regular'], ' ',
                                          current_border_format)
            self._sheet.merge_range(self._current_row, start_col+symbol_id_width,
                                    self._current_row, start_col+symbol_id_width+symbol_name_width-1,
                                    "", current_border_format)
            self._sheet.write_rich_string(self._current_row, start_col+symbol_id_width,
                                          self._text_formats['text_regular'], ' '.join(self._paytable.GetSymbolName(index).split('_')),
                                          self._text_formats['text_regular'], ' ',
                                          current_border_format)
            for i, paying_index in enumerate(paying_combinations):
                self._sheet.merge_range(self._current_row, start_col+symbol_id_width+symbol_name_width+i*one_symbol_pay_width,
                                        self._current_row, start_col+symbol_id_width+symbol_name_width+((i+1)*one_symbol_pay_width-1),
                                        "", current_border_format)
                self._sheet.write_rich_string(self._current_row, start_col+symbol_id_width+symbol_name_width+i*one_symbol_pay_width,
                                              self._text_formats['text_regular'],
                                              '{:.1f}'.format(self._paytable.GetSymbolPay(index, paying_index+1)/self._base_bet),
                                              self._text_formats['text_regular'], 'x ',
                                              self._text_formats['text_smallest'], 'bets',
                                              current_border_format)
            self._current_row += 1
        self._current_row += 3

    def _WriteWinlines(self):
        self.WriteHeader("Winlines", print_counter=True, include_in_content=True)

        start_col = 1
        winline_width = self._board_shape[1]
        winline_height = self._board_shape[0]
        winline_pairs = []

        for i in range(0, len(self._winlines), 2):
            winline_pairs.append(self._winlines[i: i+2])

        winline_index = 0
        for winlines in winline_pairs:
            zero_col = start_col
            for winline in winlines:
                self._WriteWinline(winline, winline_index, winline_height, self._current_row, zero_col)
                winline_index += 1
                zero_col += winline_width+1

            self._current_row += winline_height + 2
        self._current_row += 2

    def _WriteWinline(self, winline: list, winline_index: int, winline_height: int, row: int, col: int):
        cur_row = row
        cur_col = col
        self._sheet.merge_range(row, col, row, col + len(winline) - 1,
                                "Index " + str(winline_index), self._text_formats['text_regular'])

        cur_row += 1

        for inner_row in range(cur_row, cur_row + winline_height):
            self._sheet.set_row_pixels(inner_row, 30)
            for inner_col in range(cur_col, cur_col + len(winline)):
                cur_format = self._par_formats.GetWinlineFormat(self._board_shape[0],
                                                                self._board_shape[1],
                                                                inner_row - cur_row,
                                                                inner_col - cur_col,
                                                                winline[inner_col-cur_col] == inner_row-cur_row)
                self._sheet.write_rich_string(inner_row, inner_col,
                                              self._text_formats['text_regular'],
                                              str(inner_row-cur_row) if winline[inner_col-cur_col] == inner_row-cur_row else " ",
                                              self._text_formats['text_regular'], " ",
                                              cur_format)

    def _WriteRTPDistributions(self):
        self.WriteHeader("RTP Distribution", include_in_content=True, print_counter=True)
        self.WriteInnerHeader("Total RTP Distribution", include_in_content=True, print_counter=True)
        self._WriteSpinWinDistributionByRanges(self._total_spin_win_df, "Total_rtp_distribution", 0)
        for i, section_name in enumerate(self._section_names):
            self._current_row += 2
            self.WriteInnerHeader(section_name + " RTP distribution", print_counter=True, include_in_content=True)
            self._WriteSpinWinDistributionByRanges(self._section_all_spin_wins[i], section_name + "_rtp_distribution", i+1)
        self._current_row += 2

    def _WriteSpinWinDistributionByRanges(self, spin_win_df: pd.DataFrame, graph_name: str, color_index: int):
        start_col = 1
        ranges_width = 2
        win_freq_width = 2
        avg_win_width = 2
        rtp_width = 1
        rtp_picture_width = 5

        border_description = 'border_description_' + str(color_index)
        border_even = 'border_even_' + str(color_index)
        border_odd = 'border_odd_' + str(color_index)

        self._sheet.set_row_pixels(self._current_row, 30)
        self._sheet.merge_range(self._current_row, start_col,
                                self._current_row, start_col+ranges_width-1,
                                "", self._rtp_distribution_formats[border_description])
        self._sheet.write_rich_string(self._current_row, start_col,
                                      self._text_formats['text_big_bold'], "Win ",
                                      self._text_formats['text_big_bold'], "Ranges",
                                      self._rtp_distribution_formats[border_description])
        self._sheet.merge_range(self._current_row, start_col+ranges_width,
                                self._current_row, start_col+ranges_width+win_freq_width-1,
                                "", self._rtp_distribution_formats[border_description])
        self._sheet.write_rich_string(self._current_row, start_col+ranges_width,
                                      self._text_formats['text_big_bold'], "Win ",
                                      self._text_formats['text_big_bold'], "in Range",
                                      self._rtp_distribution_formats[border_description])
        self._sheet.merge_range(self._current_row, start_col+ranges_width+win_freq_width,
                                self._current_row, start_col+ranges_width+win_freq_width+avg_win_width-1,
                                "", self._rtp_distribution_formats[border_description])
        self._sheet.write_rich_string(self._current_row, start_col+ranges_width+win_freq_width,
                                      self._text_formats["text_big_bold"], "Average ",
                                      self._text_formats["text_big_bold"], "Win",
                                      self._rtp_distribution_formats[border_description])
        self._sheet.merge_range(self._current_row, start_col+ranges_width+win_freq_width+avg_win_width,
                                self._current_row, start_col+ranges_width+win_freq_width+avg_win_width+rtp_width+rtp_picture_width-1,
                                "", self._rtp_distribution_formats[border_description])
        self._sheet.write_rich_string(self._current_row, start_col+ranges_width+win_freq_width+avg_win_width,
                                      self._text_formats['text_big_bold'], "RTP ",
                                      self._text_formats['text_big_bold'], "per Range",
                                      self._rtp_distribution_formats[border_description])

        self._current_row += 1

        self._graphMaker.MakeBarPlot_SimpleRTPDistribution(spin_win_df, graph_name, color_index)
        self._sheet.insert_image(self._current_row, start_col + ranges_width + win_freq_width + avg_win_width + rtp_width,
                                 self._graphMaker.GetSpinWinRTPPath(graph_name),
                                 {
                                     'x_offset': 1,
                                     'y_offset': 1,
                                     'x_scale': 0.553,
                                     'y_scale': 0.416,
                                     'object_position': 2
                                 })

        for i in range(len(spin_win_df.index)):
            self._sheet.set_row_pixels(self._current_row, 30)
            current_border_format = self._rtp_distribution_formats[border_even] if i % 2 == 0 else self._rtp_distribution_formats[border_odd]

            self._sheet.merge_range(self._current_row, start_col,
                                    self._current_row, start_col+ranges_width-1,
                                    "", current_border_format)
            self._sheet.write_rich_string(self._current_row, start_col,
                                          self._text_formats['text_regular_bold'], spin_win_df.index[i][:1],
                                          self._text_formats['text_regular_bold'], spin_win_df.index[i][1:],
                                          current_border_format)
            self._sheet.merge_range(self._current_row, start_col+ranges_width,
                                    self._current_row, start_col+ranges_width+win_freq_width-1,
                                    "", current_border_format)
            cur_win_1_in = np.round(spin_win_df['win_1_in'].iloc[i], 2)
            self._sheet.write_rich_string(self._current_row, start_col+ranges_width,
                                          self._text_formats['text_smallest'], "1 in ",
                                          self._text_formats['text_regular'], '{:,.2f}'.format(cur_win_1_in),
                                          current_border_format)
            self._sheet.merge_range(self._current_row, start_col+ranges_width+win_freq_width,
                                    self._current_row, start_col+ranges_width+win_freq_width+avg_win_width-1,
                                    "", current_border_format)
            cur_avg_win = np.round(spin_win_df['avg_win'].iloc[i] / self._base_bet, 2)
            self._sheet.write_rich_string(self._current_row, start_col+ranges_width+win_freq_width,
                                          self._text_formats['text_regular'], '{:.2f}'.format(cur_avg_win),
                                          self._text_formats['text_regular'], 'x' if not pd.isna(cur_avg_win) else " ",
                                          current_border_format)
            if rtp_width > 1:
                self._sheet.merge_range(self._current_row, start_col+ranges_width+win_freq_width+avg_win_width,
                                        self._current_row, start_col+ranges_width+win_freq_width+avg_win_width+rtp_width-1,
                                        "", current_border_format)
            cur_rtp_num = spin_win_df['rtp'].iloc[i]
            self._sheet.write_rich_string(self._current_row, start_col+ranges_width+win_freq_width+avg_win_width,
                                          self._text_formats['text_regular'], '{:.3f}%'.format(cur_rtp_num),
                                          self._text_formats['text_regular'], ' ',
                                          current_border_format)

            self._current_row += 1

        self._sheet.set_row_pixels(self._current_row, 30)
        self._sheet.merge_range(self._current_row, start_col,
                                self._current_row, start_col+ranges_width-1,
                                "", self._rtp_distribution_formats[border_description])
        self._sheet.write_rich_string(self._current_row, start_col,
                                      self._text_formats['text_big_bold'], "Tota",
                                      self._text_formats['text_big_bold'], "l",
                                      self._rtp_distribution_formats[border_description])
        self._sheet.merge_range(self._current_row, start_col+ranges_width,
                                self._current_row, start_col+ranges_width+win_freq_width-1,
                                "", self._rtp_distribution_formats[border_description])
        total_win_freq = np.round(np.sum(spin_win_df['total_counter']) / (np.sum(spin_win_df['total_counter']) - spin_win_df['total_counter'].iloc[0]), 2)
        self._sheet.write_rich_string(self._current_row, start_col+ranges_width,
                                      self._text_formats['text_smallest'], '1 in ',
                                      self._text_formats['text_regular_bold'], '{:.2f}'.format(total_win_freq),
                                      self._rtp_distribution_formats[border_description])
        self._sheet.merge_range(self._current_row, start_col+ranges_width+win_freq_width,
                                self._current_row, start_col+ranges_width+win_freq_width+avg_win_width-1,
                                "", self._rtp_distribution_formats[border_description])
        total_avg_win = np.sum(spin_win_df['total_win']) / np.sum(spin_win_df['total_counter'].iloc[1:]) / self._base_bet
        self._sheet.write_rich_string(self._current_row, start_col+ranges_width+win_freq_width,
                                      self._text_formats['text_regular_bold'], '{:.2f}'.format(total_avg_win),
                                      self._text_formats['text_regular_bold'], "x",
                                      self._rtp_distribution_formats[border_description])
        if rtp_width > 1:
            self._sheet.merge_range(self._current_row, start_col+ranges_width+win_freq_width+avg_win_width,
                                    self._current_row, start_col+ranges_width+win_freq_width+avg_win_width+rtp_width-1,
                                    "", self._rtp_distribution_formats["border_description"])
        total_rtp = np.sum(spin_win_df['total_win']) / (self._base_bet * self._spin_count) * 100
        self._sheet.write_rich_string(self._current_row, start_col+ranges_width+win_freq_width+avg_win_width,
                                      self._text_formats['text_regular_bold'], '{:.3f}%'.format(total_rtp),
                                      self._text_formats['text_regular_bold'], ' ',
                                      self._rtp_distribution_formats[border_description])
        self._current_row += 2

    def _WriteWinlinesWins(self):
        self.WriteHeader("Symbols wins sorted by payout", include_in_content=True, print_counter=True)
        for i, line_win_df in enumerate(self._line_wins_dfs):
            self._WriteRegularWinsSpreadsheet(line_win_df, self._section_names[i])

    def _WriteRegularWinsSpreadsheet(self, line_wins: pd.DataFrame, section_name: str):
        self.WriteInnerHeader(section_name + " Symbols Wins", print_counter=True, include_in_content=True)

        start_col = 1
        symbol_width = 3
        length_width = 2
        multiplier_width = 2
        probability_width = 2
        pulls_width = 2
        paytable_width = 2
        rtp_width = 2
        symbol_rtp_width = 3
        rtp_by_symbol_width = 10
        rtp_by_winline_len_width = 8

        section_index = self._section_names.index(section_name)

        self._sheet.set_row_pixels(self._current_row, 40)
        self._sheet.merge_range(self._current_row, start_col, self._current_row, start_col+symbol_width-1,
                                "", self._line_wins_formats['border_description_'+str(section_index)])
        self._sheet.write_rich_string(self._current_row, start_col,
                                      self._text_formats['text_big_bold'], "Symbol",
                                      self._text_formats['text_big_bold'], " ",
                                      self._line_wins_formats['border_description_'+str(section_index)])
        self._sheet.merge_range(self._current_row, start_col+symbol_width,
                                self._current_row, start_col + symbol_width + length_width - 1,
                                "", self._line_wins_formats['border_description_' + str(section_index)])
        self._sheet.write_rich_string(self._current_row, start_col+symbol_width,
                                      self._text_formats['text_big_bold'], "Counter",
                                      self._text_formats['text_big_bold'], " ",
                                      self._line_wins_formats['border_description_' + str(section_index)])
        self._sheet.merge_range(self._current_row, start_col + symbol_width + length_width,
                                self._current_row, start_col + symbol_width + length_width + multiplier_width - 1,
                                "", self._line_wins_formats['border_description_' + str(section_index)])
        self._sheet.write_rich_string(self._current_row, start_col + symbol_width + length_width,
                                      self._text_formats['text_big_bold'], "Multiplier",
                                      self._text_formats['text_big_bold'], " ",
                                      self._line_wins_formats['border_description_' + str(section_index)])
        self._sheet.merge_range(self._current_row, start_col + symbol_width + length_width + multiplier_width,
                                self._current_row, start_col + symbol_width + length_width + multiplier_width + probability_width - 1,
                                "", self._line_wins_formats['border_description_' + str(section_index)])
        self._sheet.write_rich_string(self._current_row, start_col + symbol_width + length_width + multiplier_width,
                                      self._text_formats['text_big_bold'], "Probability",
                                      self._text_formats['text_big_bold'], " ",
                                      self._line_wins_formats['border_description_' + str(section_index)])
        self._sheet.merge_range(self._current_row, start_col + symbol_width + length_width + multiplier_width + probability_width,
                                self._current_row,
                                start_col + symbol_width + length_width + multiplier_width + probability_width +pulls_width - 1,
                                "", self._line_wins_formats['border_description_' + str(section_index)])
        self._sheet.write_rich_string(self._current_row, start_col + symbol_width + length_width + multiplier_width + probability_width,
                                      self._text_formats['text_regular_bold'], "Pulls to hit\n",
                                      self._text_formats['text_regular_bold'], "(1 in)",
                                      self._line_wins_formats['border_description_' + str(section_index)])
        self._sheet.merge_range(self._current_row,
                                start_col + symbol_width + length_width + multiplier_width + probability_width + pulls_width,
                                self._current_row,
                                start_col + symbol_width + length_width + multiplier_width + probability_width + pulls_width + paytable_width - 1,
                                "", self._line_wins_formats['border_description_' + str(section_index)])
        self._sheet.write_rich_string(self._current_row,
                                      start_col + symbol_width + length_width + multiplier_width + probability_width + pulls_width,
                                      self._text_formats['text_big_bold'], "Paytable",
                                      self._text_formats['text_big_bold'], " ",
                                      self._line_wins_formats['border_description_' + str(section_index)])
        self._sheet.merge_range(self._current_row,
                                start_col + symbol_width + length_width + multiplier_width + probability_width + pulls_width + paytable_width,
                                self._current_row,
                                start_col + symbol_width + length_width + multiplier_width + probability_width + pulls_width + paytable_width + rtp_width - 1,
                                "", self._line_wins_formats['border_description_' + str(section_index)])
        self._sheet.write_rich_string(self._current_row,
                                      start_col + symbol_width + length_width + multiplier_width + probability_width + pulls_width + paytable_width,
                                      self._text_formats['text_big_bold'], "RTP",
                                      self._text_formats['text_big_bold'], " ",
                                      self._line_wins_formats['border_description_' + str(section_index)])
        self._sheet.merge_range(self._current_row,
                                start_col + symbol_width + length_width + multiplier_width + probability_width + pulls_width + paytable_width + rtp_width,
                                self._current_row,
                                start_col + symbol_width + length_width + multiplier_width + probability_width + pulls_width + paytable_width + rtp_width + symbol_rtp_width - 1,
                                "", self._line_wins_formats['border_description_' + str(section_index)])
        self._sheet.write_rich_string(self._current_row,
                                      start_col + symbol_width + length_width + multiplier_width + probability_width + pulls_width + paytable_width + rtp_width,
                                      self._text_formats['text_big_bold'], "Total Symbol ",
                                      self._text_formats['text_big_bold'], "RTP",
                                      self._line_wins_formats['border_description_' + str(section_index)])
        self._current_row += 1

        symbol_length_rows = dict()
        for i in range(len(line_wins.index)):
            cur_symbol = line_wins["symbol_id"].iloc[i]
            cur_length = line_wins["winline_len"].iloc[i]
            if cur_symbol not in symbol_length_rows.keys():
                symbol_length_rows[cur_symbol] = dict()
                symbol_length_rows[cur_symbol][cur_length] = [i]
            elif cur_length not in symbol_length_rows[cur_symbol]:
                symbol_length_rows[cur_symbol][cur_length] = [i]
            else:
                symbol_length_rows[cur_symbol][cur_length].append(i)

        start_table_row = self._current_row
        for symbol, counters in symbol_length_rows.items():
            tot_symbol_height = 0
            for length, rows in counters.items():
                tot_symbol_height += len(rows)
            self._sheet.merge_range(self._current_row, start_col,
                                    self._current_row+tot_symbol_height-1, start_col + symbol_width-1,
                                    "", self._line_wins_formats['border_symbol_left_'+str(section_index)])
            self._sheet.write_rich_string(self._current_row, start_col,
                                          self._text_formats['text_regular'], ' '.join(self._paytable.GetSymbolName(symbol).split("_")),
                                          self._text_formats['text_regular'], " ",
                                          self._line_wins_formats['border_symbol_left_'+str(section_index)])
            self._current_row += tot_symbol_height

        for row in range(start_table_row, self._current_row):
            self._sheet.set_row_pixels(row, 30)

        self._current_row = start_table_row
        for symbol, counter in symbol_length_rows.items():
            for i, (count, rows) in enumerate(counter.items()):
                cur_format = self._line_wins_formats['border_length_middle_' + str(section_index)]
                if len(counter) == 1:
                    cur_format = self._line_wins_formats['border_length_one_' + str(section_index)]
                elif i == 0:
                    cur_format = self._line_wins_formats['border_length_up_' + str(section_index)]
                elif i == len(counter)-1:
                    cur_format = self._line_wins_formats['border_length_down_' + str(section_index)]
                self._sheet.merge_range(self._current_row, start_col+symbol_width,
                                        self._current_row+len(rows)-1, start_col+symbol_width+length_width-1,
                                        "", cur_format)
                self._sheet.write_rich_string(self._current_row, start_col+symbol_width,
                                              self._text_formats['text_regular'], str(count),
                                              self._text_formats['text_regular'], " ",
                                              cur_format)
                self._current_row += len(rows)

        self._current_row = start_table_row
        for i in range(len(line_wins.index)):
            cur_symbol = line_wins['symbol_id'].iloc[i]
            cur_length = line_wins['winline_len'].iloc[i]
            parity = 'even' if i % 2 == 0 else 'odd'

            symbol_rows = []
            for leng, rows in symbol_length_rows[cur_symbol].items():
                symbol_rows += rows

            cur_format = self._line_wins_formats['border_'+parity+'_one_'+str(section_index)]
            if len(symbol_rows) > 1 and i == symbol_rows[0]:
                cur_format = self._line_wins_formats['border_'+parity+'_up_'+str(section_index)]
            elif len(symbol_rows) > 1 and i == symbol_rows[-1]:
                cur_format = self._line_wins_formats['border_'+parity+'_down_'+str(section_index)]
            elif (len(symbol_rows) > 1 and
                  i != symbol_rows[0] and
                  i != symbol_rows[-1]):
                cur_format = self._line_wins_formats['border_'+parity+'_middle_'+str(section_index)]

            self._sheet.merge_range(self._current_row, start_col+symbol_width+length_width,
                                    self._current_row, start_col+symbol_width+length_width+multiplier_width-1,
                                    "", cur_format)
            self._sheet.write_rich_string(self._current_row, start_col+symbol_width+length_width,
                                          self._text_formats['text_regular'], str(line_wins['multiplier'].iloc[i]),
                                          self._text_formats['text_regular'], " ",
                                          cur_format)
            self._sheet.merge_range(self._current_row, start_col + symbol_width + length_width + multiplier_width,
                                    self._current_row, start_col + symbol_width + length_width + multiplier_width + probability_width - 1,
                                    "", cur_format)
            cur_prob = line_wins['probability'].iloc[i]
            self._sheet.write_rich_string(self._current_row, start_col + symbol_width + length_width + multiplier_width,
                                          self._text_formats['text_regular'], '{:.8f}'.format(cur_prob),
                                          self._text_formats['text_regular'], " ",
                                          cur_format)
            self._sheet.merge_range(self._current_row, start_col + symbol_width + length_width + multiplier_width + probability_width,
                                    self._current_row,
                                    start_col + symbol_width + length_width + multiplier_width + probability_width +pulls_width- 1,
                                    "", cur_format)
            cur_pulls = line_wins['pulls_to_hit'].iloc[i]
            self._sheet.write_rich_string(self._current_row, start_col + symbol_width + length_width + multiplier_width + probability_width,
                                          self._text_formats['text_regular'], '{:,.2f}'.format(cur_pulls),
                                          self._text_formats['text_regular'], " ",
                                          cur_format)
            self._sheet.merge_range(self._current_row,
                                    start_col + symbol_width + length_width + multiplier_width + probability_width + pulls_width,
                                    self._current_row,
                                    start_col + symbol_width + length_width + multiplier_width + probability_width + pulls_width + paytable_width - 1,
                                    "", cur_format)
            cur_paytable = self._paytable.GetSymbolPay(cur_symbol, cur_length) * line_wins['multiplier'].iloc[i] / self._base_bet
            self._sheet.write_rich_string(self._current_row,
                                          start_col + symbol_width + length_width + multiplier_width + probability_width+pulls_width,
                                          self._text_formats['text_regular'], '{:,.2f}x'.format(cur_paytable),
                                          self._text_formats['text_smallest'], " bets",
                                          cur_format)
            self._sheet.merge_range(self._current_row,
                                    start_col + symbol_width + length_width + multiplier_width + probability_width + pulls_width + paytable_width,
                                    self._current_row,
                                    start_col + symbol_width + length_width + multiplier_width + probability_width + pulls_width + paytable_width + rtp_width - 1,
                                    "", cur_format)
            self._sheet.write_rich_string(self._current_row,
                                          start_col + symbol_width + length_width + multiplier_width + probability_width + pulls_width + paytable_width,
                                          self._text_formats['text_regular'], '{:.3f}'.format(line_wins['rtp_perc'].iloc[i]),
                                          self._text_formats['text_smallest'], "%",
                                          cur_format)
            self._current_row += 1

        self._current_row = start_table_row
        total_section_symbols_rtp = 0
        total_symbols_rtps = dict()
        rtp_by_winline_length = dict()
        for symbol, counters in symbol_length_rows.items():
            symbol_total_rtp = 0
            symbol_height = 0
            for winline_len, rows in counters.items():
                symbol_height += len(rows)
                for row in rows:
                    symbol_total_rtp += line_wins['rtp_perc'].iloc[row]
                    if winline_len not in rtp_by_winline_length:
                        rtp_by_winline_length[winline_len] = line_wins['rtp_perc'].iloc[row]
                    else:
                        rtp_by_winline_length[winline_len] += line_wins['rtp_perc'].iloc[row]
            total_symbols_rtps[symbol] = symbol_total_rtp
            total_section_symbols_rtp += symbol_total_rtp
            self._sheet.merge_range(self._current_row, start_col+symbol_width+length_width+multiplier_width+probability_width+pulls_width+paytable_width+rtp_width,
                                    self._current_row+symbol_height-1, start_col+symbol_width+length_width+multiplier_width+probability_width+pulls_width+paytable_width+rtp_width+symbol_rtp_width-1,
                                    "", self._line_wins_formats['border_symbol_right_' +str(section_index)])
            self._sheet.write_rich_string(self._current_row, start_col+symbol_width+length_width+multiplier_width+probability_width+pulls_width+paytable_width+rtp_width,
                                          self._text_formats['text_regular'], '{:.3f}'.format(symbol_total_rtp),
                                          self._text_formats['text_regular'], "%",
                                          self._line_wins_formats['border_symbol_right_'+str(section_index)])
            self._current_row += symbol_height

        self._sheet.set_row_pixels(self._current_row, 30)
        self._sheet.merge_range(self._current_row, start_col+symbol_width+length_width+multiplier_width+probability_width+pulls_width+paytable_width+rtp_width,
                                self._current_row, start_col+symbol_width+length_width+multiplier_width+probability_width+pulls_width+paytable_width+rtp_width+symbol_rtp_width-1,
                                "", self._line_wins_formats['border_description_'+str(section_index)])
        self._sheet.write_rich_string(self._current_row, start_col+symbol_width+length_width+multiplier_width+probability_width+pulls_width+paytable_width+rtp_width,
                                      self._text_formats['text_big_bold'], '{:.3f}'.format(total_section_symbols_rtp),
                                      self._text_formats['text_big_bold'], "%",
                                      self._line_wins_formats['border_description_' + str(section_index)])

        self._current_row += 2
        self._sheet.set_row_pixels(self._current_row, 30)
        self._sheet.merge_range(self._current_row, start_col,
                                self._current_row, start_col + rtp_by_symbol_width -1,
                                "", self._line_wins_formats['border_description_'+str(section_index)])
        self._sheet.write_rich_string(self._current_row, start_col,
                                      self._text_formats['text_big_bold'], self._section_names[section_index],
                                      self._text_formats['text_big_bold'], ' RTP distribution by Symbols',
                                      self._line_wins_formats['border_description_'+str(section_index)])
        self._sheet.merge_range(self._current_row, start_col + rtp_by_symbol_width,
                                self._current_row, start_col + rtp_by_symbol_width + rtp_by_winline_len_width - 1,
                                "", self._line_wins_formats['border_description_' + str(section_index)])
        self._sheet.write_rich_string(self._current_row, start_col + rtp_by_symbol_width,
                                      self._text_formats['text_big_bold'], self._section_names[section_index],
                                      self._text_formats['text_big_bold'], ' RTP distribution by Winline Length',
                                      self._line_wins_formats['border_description_' + str(section_index)])

        self._current_row += 1
        self._graphMaker.MakePiePlot_RTPbySymbol(total_symbols_rtps, section_index)
        self._graphMaker.MakePiePlot_RTPbyWinlineLen(rtp_by_winline_length, section_index)
        self._sheet.insert_image(self._current_row, start_col,
                                 self._graphMaker.GetPiePlotRTPbySymbols(section_index),
                                 {
                                     'x_offset': 1,
                                     'y_offset': 1,
                                     'x_scale': 0.52,
                                     'y_scale': 0.52,
                                     'object_position': 2
                                 })
        self._sheet.insert_image(self._current_row, start_col+10,
                                 self._graphMaker.GetPiePlotRTPbyWinlineLen(section_index),
                                 {
                                     'x_offset': 7,
                                     'y_offset': 1,
                                     'x_scale': 0.526,
                                     'y_scale': 0.52,
                                     'object_position': 2
                                 })
        self._current_row += 24


    def _WriteConfidenceIntervals(self):
        self.WriteHeader("RTP Confidence Intervals", print_counter=True, include_in_content=True)

        start_col = 1
        tables_step = 8

        for perc in self._confidence_percents:
            int_perc = int(perc * 100)
            self.WriteInnerHeader(str(int_perc) + "% Confidence Intervals")
            start_row = self._current_row
            self._WriteConfidenceInterval(self._total_confidence_dfs[perc],
                                          self._total_spin_win_struct,
                                          int_perc,
                                          start_col,
                                          -1)
            for section_index in range(len(self._section_names)):
                self._current_row = start_row
                self._WriteConfidenceInterval(self._sections_confidence_dfs[(perc, section_index)],
                                              self._sections_spin_win_structs[section_index],
                                              int_perc,
                                              start_col + (section_index+1) * tables_step,
                                              section_index)
            self._current_row += 3



    def _WriteConfidenceInterval(self,
                                 conf_df: pd.DataFrame,
                                 spin_win_struct: BasicSpinWin,
                                 confidence: int,
                                 start_col: int,
                                 section_index: int):
        section_index_str = 'total' if section_index == -1 else str(section_index)
        info_width = 4
        val_width = 3

        num_games_width = 3
        left_width = 2
        right_width = 2

        self._sheet.set_row_pixels(self._current_row, 30)
        self._sheet.merge_range(self._current_row, start_col,
                                self._current_row, start_col + info_width + val_width - 1,
                                "", self._confidence_formats['border_description_'+section_index_str])
        main_text = ('Total' if section_index == -1 else self._section_names[section_index]) + " RTP Confidence Intervals"
        self._sheet.write_rich_string(self._current_row, start_col,
                                      self._text_formats['text_big_bold'], main_text[:1],
                                      self._text_formats['text_big_bold'], main_text[1:],
                                      self._confidence_formats['border_description_'+section_index_str])

        self._current_row += 1

        self._sheet.set_row_pixels(self._current_row, 30)
        self._sheet.merge_range(self._current_row, start_col,
                                self._current_row, start_col + info_width - 1,
                                "", self._confidence_formats['border_even_'+section_index_str])
        self._sheet.write_rich_string(self._current_row, start_col,
                                      self._text_formats['text_regular'], 'Theoretical RTP (bet): ',
                                      self._text_formats['text_regular'], ' ',
                                      self._confidence_formats['border_even_'+section_index_str])
        self._sheet.merge_range(self._current_row, start_col + info_width,
                                self._current_row, start_col + info_width + val_width - 1,
                                "", self._confidence_formats['border_even_' + section_index_str])
        rtp = np.round(spin_win_struct.GetRTP(True), 4)
        self._sheet.write_rich_string(self._current_row, start_col + info_width,
                                      self._text_formats['text_regular'], str(rtp),
                                      self._text_formats['text_regular'], ' ',
                                      self._confidence_formats['border_even_' + section_index_str])

        self._current_row += 1

        self._sheet.set_row_pixels(self._current_row, 30)
        self._sheet.merge_range(self._current_row, start_col,
                                self._current_row, start_col + info_width - 1,
                                "", self._confidence_formats['border_odd_' + section_index_str])
        self._sheet.write_rich_string(self._current_row, start_col,
                                      self._text_formats['text_regular'], 'Standard deviation (bet): ',
                                      self._text_formats['text_regular'], ' ',
                                      self._confidence_formats['border_odd_' + section_index_str])
        self._sheet.merge_range(self._current_row, start_col + info_width,
                                self._current_row, start_col + info_width + val_width - 1,
                                "", self._confidence_formats['border_odd_' + section_index_str])
        std = np.round(spin_win_struct.GetSTD(), 4)
        self._sheet.write_rich_string(self._current_row, start_col + info_width,
                                      self._text_formats['text_regular'], str(std),
                                      self._text_formats['text_regular'], ' ',
                                      self._confidence_formats['border_odd_' + section_index_str])

        self._current_row += 1

        self._sheet.set_row_pixels(self._current_row, 30)
        self._sheet.merge_range(self._current_row, start_col,
                                self._current_row, start_col + info_width - 1,
                                "", self._confidence_formats['border_even_' + section_index_str])
        self._sheet.write_rich_string(self._current_row, start_col,
                                      self._text_formats['text_regular'], 'Confidence: ',
                                      self._text_formats['text_regular'], ' ',
                                      self._confidence_formats['border_even_' + section_index_str])
        self._sheet.merge_range(self._current_row, start_col + info_width,
                                self._current_row, start_col + info_width + val_width - 1,
                                "", self._confidence_formats['border_even_' + section_index_str])
        self._sheet.write_rich_string(self._current_row, start_col + info_width,
                                      self._text_formats['text_regular_bold'], str(confidence),
                                      self._text_formats['text_regular_bold'], '%',
                                      self._confidence_formats['border_even_' + section_index_str])

        self._current_row += 2

        self._sheet.set_row_pixels(self._current_row, 30)
        self._sheet.merge_range(self._current_row, start_col,
                                self._current_row, start_col+num_games_width-1,
                                '', self._confidence_formats['border_description_' + section_index_str])
        self._sheet.write_rich_string(self._current_row, start_col,
                                      self._text_formats['text_big_bold'], 'Number of ',
                                      self._text_formats['text_big_bold'], 'Games',
                                      self._confidence_formats['border_description_' + section_index_str])
        self._sheet.merge_range(self._current_row, start_col+num_games_width,
                                self._current_row, start_col + num_games_width + left_width - 1,
                                '', self._confidence_formats['border_description_' + section_index_str])
        self._sheet.write_rich_string(self._current_row, start_col+num_games_width,
                                      self._text_formats['text_big_bold'], 'RTP Min ',
                                      self._text_formats['text_big_bold'], ' ',
                                      self._confidence_formats['border_description_' + section_index_str])
        self._sheet.merge_range(self._current_row, start_col + num_games_width + left_width,
                                self._current_row, start_col + num_games_width + left_width + right_width - 1,
                                '', self._confidence_formats['border_description_' + section_index_str])
        self._sheet.write_rich_string(self._current_row, start_col + num_games_width + left_width,
                                      self._text_formats['text_big_bold'], 'RTP Max ',
                                      self._text_formats['text_big_bold'], ' ',
                                      self._confidence_formats['border_description_' + section_index_str])
        self._current_row += 1

        for i in range(len(conf_df.index)):
            self._sheet.set_row_pixels(self._current_row, 30)
            current_format = self._confidence_formats['border_even_'+section_index_str] if i % 2 == 0 else self._confidence_formats['border_odd_'+section_index_str]

            self._sheet.merge_range(self._current_row, start_col,
                                    self._current_row, start_col + num_games_width - 1,
                                    '', current_format)
            self._sheet.write_rich_string(self._current_row, start_col,
                                          self._text_formats['text_regular_bold'], conf_df['num_games'].iloc[i],
                                          self._text_formats['text_regular_bold'], ' ',
                                          current_format)
            self._sheet.merge_range(self._current_row, start_col+num_games_width,
                                    self._current_row, start_col + num_games_width + left_width - 1,
                                    '', current_format)
            self._sheet.write_rich_string(self._current_row, start_col+num_games_width,
                                          self._text_formats['text_regular'], '{:.4f}%'.format(conf_df['left_border'].iloc[i] * 100),
                                          self._text_formats['text_regular'], ' ',
                                          current_format)
            self._sheet.merge_range(self._current_row, start_col + num_games_width + left_width,
                                    self._current_row, start_col + num_games_width + left_width + right_width - 1,
                                    '', current_format)
            self._sheet.write_rich_string(self._current_row, start_col + num_games_width + left_width,
                                          self._text_formats['text_regular'],
                                          '{:.4f}%'.format(conf_df['right_border'].iloc[i] * 100),
                                          self._text_formats['text_regular'], ' ',
                                          current_format)
            self._current_row += 1


    def _WriteWinAmountStatistics(self):
        self.WriteHeader("Win Amount Statistics", print_counter=True, include_in_content=True)
        self._WriteLiabilityRanges()
        self._WriteTopAward()
        self._WriteHighestWin()

    def _WriteLiabilityRanges(self):
        pass

    def _WriteTopAward(self):
        self.WriteInnerHeader("Top Award Statistics", print_counter=True, include_in_content=True)

        start_col = 4
        number_of_games_width = 3
        top_award_width = 3

        self._sheet.set_row_pixels(self._current_row, 30)
        self._sheet.merge_range(self._current_row, start_col,
                                self._current_row, start_col + number_of_games_width - 1,
                                "", self._top_award_formats['border_description'])
        self._sheet.write_rich_string(self._current_row, start_col,
                                      self._text_formats['text_big_bold'], 'Number of ',
                                      self._text_formats['text_big_bold'], 'Games',
                                      self._top_award_formats['border_description'])
        self._sheet.merge_range(self._current_row, start_col + number_of_games_width,
                                self._current_row, start_col + number_of_games_width + top_award_width - 1,
                                "", self._top_award_formats['border_description'])
        self._sheet.write_rich_string(self._current_row, start_col + number_of_games_width,
                                      self._text_formats['text_big_bold'], 'Top ',
                                      self._text_formats['text_big_bold'], 'Award',
                                      self._top_award_formats['border_description'])


        self._current_row += 1

        for i in range(len(self._top_award_df.index)):
            cur_format = self._top_award_formats['border_even'] if i % 2 == 0 else self._top_award_formats['border_odd']

            self._sheet.set_row_pixels(self._current_row, 30)
            self._sheet.merge_range(self._current_row, start_col,
                                    self._current_row, start_col + number_of_games_width - 1,
                                    "", cur_format)
            self._sheet.write_rich_string(self._current_row, start_col,
                                          self._text_formats['text_regular'], self._top_award_df.index[i],
                                          self._text_formats['text_regular'], ' ',
                                          cur_format)
            self._sheet.merge_range(self._current_row, start_col + number_of_games_width,
                                    self._current_row, start_col + number_of_games_width + top_award_width - 1,
                                    "", cur_format)
            bet_win = self._top_award_df['top_award_cents'].iloc[i] / self._base_bet
            self._sheet.write_rich_string(self._current_row, start_col + number_of_games_width,
                                          self._text_formats['text_regular'], '{:.1f}x'.format(bet_win),
                                          self._text_formats['text_smallest'], ' bets ',
                                          cur_format)
            self._current_row += 1
        self._current_row += 3

    def _WriteHighestWin(self):
        self.WriteInnerHeader("Max Bet Multiplier", print_counter=True, include_in_content=True)

        start_col = 3
        info_width = 4
        val_width = 4

        self._sheet.set_row_pixels(self._current_row, 30)
        self._sheet.merge_range(self._current_row, start_col,
                                self._current_row, start_col + info_width - 1,
                                "", self._top_award_formats['border_even'])
        self._sheet.write_rich_string(self._current_row, start_col,
                                      self._text_formats['text_regular'], 'Max Bet ',
                                      self._text_formats['text_regular'], 'Multiplier:',
                                      self._top_award_formats['border_even'])
        self._sheet.merge_range(self._current_row, start_col + info_width,
                                self._current_row, start_col + info_width + val_width - 1,
                                "", self._top_award_formats['border_even'])
        max_win_in_bets = self._max_win_in_cents / self._base_bet
        self._sheet.write_rich_string(self._current_row, start_col + info_width,
                                      self._text_formats['text_regular'], '{:,.1f}x'.format(max_win_in_bets),
                                      self._text_formats['text_smallest'], ' bets',
                                      self._top_award_formats['border_even'])

        self._current_row += 1

        self._sheet.set_row_pixels(self._current_row, 30)
        self._sheet.merge_range(self._current_row, start_col,
                                self._current_row, start_col + info_width - 1,
                                "", self._top_award_formats['border_odd'])
        self._sheet.write_rich_string(self._current_row, start_col,
                                      self._text_formats['text_regular'], 'Max Bet ',
                                      self._text_formats['text_regular'], 'Multiplier Counter:',
                                      self._top_award_formats['border_odd'])
        self._sheet.merge_range(self._current_row, start_col + info_width,
                                self._current_row, start_col + info_width + val_width - 1,
                                "", self._top_award_formats['border_odd'])
        self._sheet.write_rich_string(self._current_row, start_col + info_width,
                                      self._text_formats['text_regular'], '{:,d}'.format(self._max_win_count),
                                      self._text_formats['text_regular'], ' ',
                                      self._top_award_formats['border_odd'])

        self._current_row += 1

        self._sheet.set_row_pixels(self._current_row, 30)
        self._sheet.merge_range(self._current_row, start_col,
                                self._current_row, start_col + info_width - 1,
                                "", self._top_award_formats['border_even'])
        self._sheet.write_rich_string(self._current_row, start_col,
                                      self._text_formats['text_regular'], 'Pulls to Hit',
                                      self._text_formats['text_regular'], ' Max Bet Multiplier:',
                                      self._top_award_formats['border_even'])
        self._sheet.merge_range(self._current_row, start_col + info_width,
                                self._current_row, start_col + info_width + val_width - 1,
                                "", self._top_award_formats['border_even'])
        pulls_to_hit = self._spin_count / self._max_win_count
        self._sheet.write_rich_string(self._current_row, start_col + info_width,
                                      self._text_formats['text_smallest'], '1 in ',
                                      self._text_formats['text_regular'], '{:,.2f}'.format(pulls_to_hit),
                                      self._top_award_formats['border_even'])

        self._current_row += 1

        info_txt_1 = '*calculated based on simulation of '
        info_txt_2 = ' spins'
        self._sheet.merge_range(self._current_row, start_col,
                                self._current_row, start_col + info_width + val_width - 1,
                                "", self._text_formats['text_red_info'])
        self._sheet.write_rich_string(self._current_row, start_col,
                                      self._text_formats['text_red_info'], info_txt_1,
                                      self._text_formats['text_red_info'], '{:,d}'.format(self._spin_count),
                                      self._text_formats['text_red_info'], info_txt_2,
                                      self._text_formats['text_red_info'])
        self._current_row += 3



    def WriteCustomGameStatistics(self):
        return

    def _WriteReelsetsPivotTables(self):
        self.WriteHeader("Reelsets Pivot Tables", include_in_content=True, print_counter=True)
        self._WriteAllReelsetsTable()
        self._WriteBaseReelsetsTable()

    def _WriteAllReelsetsTable(self):
        self.WriteInnerHeader("RTP distribution by All Reelsets")
        for section_index in range(len(self._section_names)):
            self._WriteSectionReelsetsTable(self._all_reelsets_section_dfs[section_index],
                                            self._section_names[section_index] + ' Reelsets', section_index)
        self._current_row += 1

    def _WriteSectionReelsetsTable(self, reelsets_df, name: str, color_index: int):
        start_col = 1
        index_width = 1
        name_width = 6
        rtp_width = 2
        rtp_common_width = 3
        prob_width = 3
        hits_width = 3
        avg_win_width = 2
        max_win_width = 2
        win_freq_width = 3
        std_width = 2

        cur_start_col = start_col
        self._sheet.set_row_pixels(self._current_row, 30)
        self._sheet.merge_range(self._current_row, cur_start_col,
                                self._current_row,
                                cur_start_col + index_width + name_width + rtp_width + rtp_common_width + prob_width + hits_width +
                                avg_win_width + max_win_width + win_freq_width + std_width - 1,
                                name, self._reelsets_table_formats['border_header_'+str(color_index)])

        self._current_row += 1

        self._sheet.set_row_pixels(self._current_row, 30)
        self._sheet.write_rich_string(self._current_row, cur_start_col,
                                      self._text_formats['text_big_bold'], 'Index',
                                      self._text_formats['text_big_bold'], ' ',
                                      self._reelsets_table_formats['border_description_'+str(color_index)])
        cur_start_col += index_width
        self._sheet.merge_range(self._current_row, cur_start_col,
                                self._current_row, cur_start_col + name_width - 1,
                                "", self._reelsets_table_formats['border_description_'+str(color_index)])
        self._sheet.write_rich_string(self._current_row, cur_start_col,
                                      self._text_formats['text_big_bold'], 'Reelset Name',
                                      self._text_formats['text_big_bold'], ' ',
                                      self._reelsets_table_formats['border_description_'+str(color_index)])
        cur_start_col += name_width
        self._sheet.merge_range(self._current_row, cur_start_col,
                                self._current_row, cur_start_col + rtp_width - 1,
                                "", self._reelsets_table_formats['border_description_' + str(color_index)])
        self._sheet.write_rich_string(self._current_row, cur_start_col,
                                      self._text_formats['text_big_bold'], 'RTP, %',
                                      self._text_formats['text_big_bold'], ' ',
                                      self._reelsets_table_formats['border_description_' + str(color_index)])
        cur_start_col += rtp_width
        self._sheet.merge_range(self._current_row, cur_start_col,
                                self._current_row, cur_start_col + rtp_common_width - 1,
                                "", self._reelsets_table_formats['border_description_' + str(color_index)])
        self._sheet.write_rich_string(self._current_row, cur_start_col,
                                      self._text_formats['text_big_bold'], 'Common RTP, %',
                                      self._text_formats['text_big_bold'], ' ',
                                      self._reelsets_table_formats['border_description_' + str(color_index)])
        cur_start_col += rtp_common_width
        self._sheet.merge_range(self._current_row, cur_start_col,
                                self._current_row, cur_start_col + prob_width - 1,
                                "", self._reelsets_table_formats['border_description_' + str(color_index)])
        self._sheet.write_rich_string(self._current_row, cur_start_col,
                                      self._text_formats['text_big_bold'], 'Probability',
                                      self._text_formats['text_big_bold'], ' ',
                                      self._reelsets_table_formats['border_description_' + str(color_index)])
        cur_start_col += prob_width
        self._sheet.merge_range(self._current_row, cur_start_col,
                                self._current_row, cur_start_col + hits_width - 1,
                                "", self._reelsets_table_formats['border_description_' + str(color_index)])
        self._sheet.write_rich_string(self._current_row, cur_start_col,
                                      self._text_formats['text_big_bold'], 'Pulls to Hit',
                                      self._text_formats['text_big_bold'], ' ',
                                      self._reelsets_table_formats['border_description_' + str(color_index)])
        cur_start_col += hits_width
        self._sheet.merge_range(self._current_row, cur_start_col,
                                self._current_row, cur_start_col + avg_win_width - 1,
                                "", self._reelsets_table_formats['border_description_' + str(color_index)])
        self._sheet.write_rich_string(self._current_row, cur_start_col,
                                      self._text_formats['text_big_bold'], 'Average Win',
                                      self._text_formats['text_big_bold'], ' ',
                                      self._reelsets_table_formats['border_description_' + str(color_index)])
        cur_start_col += avg_win_width
        self._sheet.merge_range(self._current_row, cur_start_col,
                                self._current_row, cur_start_col + max_win_width - 1,
                                "", self._reelsets_table_formats['border_description_' + str(color_index)])
        self._sheet.write_rich_string(self._current_row, cur_start_col,
                                      self._text_formats['text_big_bold'], 'Max Win',
                                      self._text_formats['text_big_bold'], ' ',
                                      self._reelsets_table_formats['border_description_' + str(color_index)])
        cur_start_col += max_win_width
        self._sheet.merge_range(self._current_row, cur_start_col,
                                self._current_row, cur_start_col + win_freq_width - 1,
                                "", self._reelsets_table_formats['border_description_' + str(color_index)])
        self._sheet.write_rich_string(self._current_row, cur_start_col,
                                      self._text_formats['text_big_bold'], 'Win Frequency',
                                      self._text_formats['text_big_bold'], ' ',
                                      self._reelsets_table_formats['border_description_' + str(color_index)])
        cur_start_col += win_freq_width
        self._sheet.merge_range(self._current_row, cur_start_col,
                                self._current_row, cur_start_col + std_width - 1,
                                "", self._reelsets_table_formats['border_description_' + str(color_index)])
        self._sheet.write_rich_string(self._current_row, cur_start_col,
                                      self._text_formats['text_big_bold'], 'STD',
                                      self._text_formats['text_big_bold'], ' ',
                                      self._reelsets_table_formats['border_description_' + str(color_index)])

        self._current_row += 1

        for i in range(len(reelsets_df.index)):
            cur_start_col = start_col
            cur_format = self._reelsets_table_formats['border_even_'+str(color_index)] if i % 2 == 0 else self._reelsets_table_formats['border_odd_'+str(color_index)]
            self._sheet.set_row_pixels(self._current_row, 30)
            self._sheet.write_rich_string(self._current_row, cur_start_col,
                                          self._text_formats['text_regular'], str(i),
                                          self._text_formats['text_regular'], ' ',
                                          cur_format)
            cur_start_col += index_width
            cur_reel_name_format = self._reelsets_table_formats['border_even_name_'+str(color_index)] if i%2==0 else self._reelsets_table_formats['border_odd_name_'+str(color_index)]
            self._sheet.merge_range(self._current_row, cur_start_col,
                                    self._current_row, cur_start_col + name_width - 1,
                                    "", cur_format)
            self._sheet.write_rich_string(self._current_row, cur_start_col,
                                          cur_reel_name_format, reelsets_df['reelset_name'].iloc[i],
                                          cur_reel_name_format, ' ',
                                          cur_reel_name_format)
            cur_start_col += name_width
            self._sheet.merge_range(self._current_row, cur_start_col,
                                    self._current_row, cur_start_col + rtp_width - 1,
                                    "", cur_format)
            rtp = reelsets_df['rtp'].iloc[i] * 100
            self._sheet.write_rich_string(self._current_row, cur_start_col,
                                          self._text_formats['text_regular'], '{:.4f}'.format(rtp),
                                          self._text_formats['text_regular'], ' ',
                                          cur_format)
            cur_start_col += rtp_width
            self._sheet.merge_range(self._current_row, cur_start_col,
                                    self._current_row, cur_start_col + rtp_common_width - 1,
                                    "", cur_format)
            common_rtp = reelsets_df['rtp_common'].iloc[i] * 100
            self._sheet.write_rich_string(self._current_row, cur_start_col,
                                          self._text_formats['text_regular'], '{:.4f}'.format(common_rtp),
                                          self._text_formats['text_regular'], ' ',
                                          cur_format)
            cur_start_col += rtp_common_width
            self._sheet.merge_range(self._current_row, cur_start_col,
                                    self._current_row, cur_start_col + prob_width - 1,
                                    "", cur_format)
            prob = reelsets_df['prob'].iloc[i]
            self._sheet.write_rich_string(self._current_row, cur_start_col,
                                          self._text_formats['text_regular'], '{:.5f}'.format(prob),
                                          self._text_formats['text_regular'], ' ',
                                          cur_format)
            cur_start_col += prob_width
            self._sheet.merge_range(self._current_row, cur_start_col,
                                    self._current_row, cur_start_col + hits_width - 1,
                                    "", cur_format)
            hits = reelsets_df['hits_1_in'].iloc[i]
            self._sheet.write_rich_string(self._current_row, cur_start_col,
                                          self._text_formats['text_smallest'], '1 in ',
                                          self._text_formats['text_regular'], '{:.3f}'.format(hits),
                                          cur_format)
            cur_start_col += hits_width
            self._sheet.merge_range(self._current_row, cur_start_col,
                                    self._current_row, cur_start_col + avg_win_width - 1,
                                    "", cur_format)
            avg_win = reelsets_df['avg_win_no_zero'].iloc[i] / self._base_bet
            self._sheet.write_rich_string(self._current_row, cur_start_col,
                                          self._text_formats['text_regular'], '{:.2f}x'.format(avg_win),
                                          self._text_formats['text_smallest'], ' bets',
                                          cur_format)
            cur_start_col += avg_win_width
            self._sheet.merge_range(self._current_row, cur_start_col,
                                    self._current_row, cur_start_col + max_win_width - 1,
                                    "", cur_format)
            max_win = reelsets_df['max_win'].iloc[i] / self._base_bet
            self._sheet.write_rich_string(self._current_row, cur_start_col,
                                          self._text_formats['text_regular'], '{:.2f}x'.format(max_win),
                                          self._text_formats['text_smallest'], ' bets',
                                          cur_format)
            cur_start_col += max_win_width
            self._sheet.merge_range(self._current_row, cur_start_col,
                                    self._current_row, cur_start_col + win_freq_width - 1,
                                    "", cur_format)
            win_fr = reelsets_df['win_freq_1_in'].iloc[i]
            self._sheet.write_rich_string(self._current_row, cur_start_col,
                                          self._text_formats['text_smallest'], '1 in ',
                                          self._text_formats['text_regular'], '{:.4f}'.format(win_fr),
                                          cur_format)
            cur_start_col += win_freq_width
            self._sheet.merge_range(self._current_row, cur_start_col,
                                    self._current_row, cur_start_col + std_width - 1,
                                    "", cur_format)
            std = reelsets_df['std'].iloc[i]
            self._sheet.write_rich_string(self._current_row, cur_start_col,
                                          self._text_formats['text_regular'], '{:.3f}x'.format(std),
                                          self._text_formats['text_smallest'], ' bets',
                                          cur_format)
            self._current_row += 1

        cur_start_col = start_col
        self._sheet.set_row_pixels(self._current_row, 30)
        border_description_format = self._reelsets_table_formats['border_description_'+str(color_index)]
        self._sheet.merge_range(self._current_row, cur_start_col,
                                self._current_row, cur_start_col + index_width + name_width - 1,
                                '', border_description_format)
        self._sheet.write_rich_string(self._current_row, cur_start_col,
                                      self._text_formats['text_big_bold'], 'Total:',
                                      self._text_formats['text_big_bold'], ' ',
                                      border_description_format)
        cur_start_col += index_width + name_width
        self._sheet.merge_range(self._current_row, cur_start_col,
                                self._current_row, cur_start_col + rtp_width - 1,
                                "", border_description_format)
        rtp_sum = reelsets_df['rtp'].sum() * 100
        self._sheet.write_rich_string(self._current_row, cur_start_col,
                                      self._text_formats['text_big_bold'], '{:.4f}'.format(rtp_sum),
                                      self._text_formats['text_big_bold'], ' ',
                                      border_description_format)
        cur_start_col += rtp_width
        self._sheet.merge_range(self._current_row, cur_start_col,
                                self._current_row, cur_start_col + rtp_common_width - 1,
                                "", border_description_format)
        common_rtp_sum = np.sum(reelsets_df['rtp_common'] * reelsets_df['prob']) * 100
        self._sheet.write_rich_string(self._current_row, cur_start_col,
                                      self._text_formats['text_big_bold'], '{:.4f}'.format(common_rtp_sum),
                                      self._text_formats['text_big_bold'], ' ',
                                      border_description_format)
        cur_start_col += rtp_common_width
        self._sheet.merge_range(self._current_row, cur_start_col,
                                self._current_row, cur_start_col + prob_width - 1,
                                "", border_description_format)
        prob_sum = reelsets_df['prob'].sum()
        self._sheet.write_rich_string(self._current_row, cur_start_col,
                                      self._text_formats['text_big_bold'], '{:.4f}'.format(prob_sum),
                                      self._text_formats['text_regular'], ' ',
                                      border_description_format)
        cur_start_col += prob_width
        self._sheet.merge_range(self._current_row, cur_start_col,
                                self._current_row, cur_start_col + hits_width - 1,
                                "", border_description_format)
        hits_sum = 1 / np.sum(1 / reelsets_df['hits_1_in'])
        self._sheet.write_rich_string(self._current_row, cur_start_col,
                                      self._text_formats['text_smallest'], '1 in ',
                                      self._text_formats['text_big_bold'], '{:.3f}'.format(hits_sum),
                                      border_description_format)
        cur_start_col += hits_width
        self._sheet.merge_range(self._current_row, cur_start_col,
                                self._current_row, cur_start_col + avg_win_width - 1,
                                "", border_description_format)
        avg_win_sum = reelsets_df['total_win'].sum() / reelsets_df['win_hits'].sum() / self._base_bet
        self._sheet.write_rich_string(self._current_row, cur_start_col,
                                      self._text_formats['text_big_bold'], '{:.2f}x'.format(avg_win_sum),
                                      self._text_formats['text_smallest'], ' bets',
                                      border_description_format)
        cur_start_col += avg_win_width
        self._sheet.merge_range(self._current_row, cur_start_col,
                                self._current_row, cur_start_col + max_win_width - 1,
                                "", border_description_format)
        max_win_glob = reelsets_df['max_win'].max() / self._base_bet
        self._sheet.write_rich_string(self._current_row, cur_start_col,
                                      self._text_formats['text_big_bold'], '{:.2f}x'.format(max_win_glob),
                                      self._text_formats['text_smallest'], ' bets',
                                      border_description_format)
        cur_start_col += max_win_width
        self._sheet.merge_range(self._current_row, cur_start_col,
                                self._current_row, cur_start_col + win_freq_width - 1,
                                "", border_description_format)
        win_fr_glob = reelsets_df['total_hits'].sum() / reelsets_df['win_hits'].sum()
        self._sheet.write_rich_string(self._current_row, cur_start_col,
                                      self._text_formats['text_smallest'], '1 in ',
                                      self._text_formats['text_big_bold'], '{:.4f}'.format(win_fr_glob),
                                      border_description_format)
        cur_start_col += win_freq_width
        self._sheet.merge_range(self._current_row, cur_start_col,
                                self._current_row, cur_start_col + std_width - 1,
                                "", border_description_format)
        std_global = reelsets_df['global_std'].iloc[0]
        self._sheet.write_rich_string(self._current_row, cur_start_col,
                                      self._text_formats['text_big_bold'], '{:.3f}x'.format(std_global),
                                      self._text_formats['text_big_bold'], ' bets',
                                      border_description_format)
        self._current_row += 3

    def _WriteBaseReelsetsTable(self):
        self.WriteInnerHeader("RTP distribution by Base Reelsets")
        self._WriteSectionReelsetsTable(self._base_reelsets_df,
                                        'All RTP Distribution by Base Game Reelsets', 0)

    def _WriteReelsetsSpinWinTables(self):
        start_col = 1
        first_info_width = 2
        first_val_width = 1
        second_info_width = 2
        second_val_width = 7
        button_width = 2

        self.WriteHeader("Reelsets Statistics", include_in_content=True, print_counter=True)
        self._AllocateSpaceReelsetsSpinWins()
        for section_index, reelsets in enumerate(self._all_reelsets_spin_win_dfs):
            text = self._section_names[section_index] + " Reelsets"
            self._reelsets_spin_win_section_content.append((str(section_index + 1) + '. ' + text, self._current_row))
            self.WriteInnerHeader(text, include_in_content=True, print_counter=True)
            for reelset_index, reelset in enumerate(reelsets):
                text = str(section_index+1)+'.'+str(reelset_index)+". Reelset: '"+self._reelsets.GetReelsetName(section_index, reelset_index) + "' spin wins"
                self._reelsets_spin_win_reelsets_content[section_index].append((text, self._current_row))
                self._sheet.merge_range(self._current_row, start_col,
                                        self._current_row, start_col + first_info_width - 1,
                                        "", self._reelsets_spin_win_formats['info_cell_'+str(section_index)])
                self._sheet.write_rich_string(self._current_row, start_col,
                                              self._reelsets_spin_win_formats['info_cell_'+str(section_index)],
                                              'Section Index:',
                                              self._reelsets_spin_win_formats['info_cell_' + str(section_index)],
                                              ' ',
                                              self._reelsets_spin_win_formats['info_cell_'+str(section_index)])
                self._sheet.write(self._current_row, start_col+first_info_width,
                                  str(section_index), self._reelsets_spin_win_formats['value_cell_'+str(section_index)])

                self._sheet.merge_range(self._current_row, start_col+first_info_width+first_val_width,
                                        self._current_row, start_col+first_info_width+first_val_width+second_info_width-1,
                                        "", self._reelsets_spin_win_formats['info_cell_' + str(section_index)])
                self._sheet.write_rich_string(self._current_row, start_col+first_info_width+first_val_width,
                                              self._reelsets_spin_win_formats['info_cell_' + str(section_index)],
                                              'Section Name:',
                                              self._reelsets_spin_win_formats['info_cell_' + str(section_index)],
                                              ' ',
                                              self._reelsets_spin_win_formats['info_cell_' + str(section_index)])

                self._sheet.merge_range(self._current_row, start_col + first_info_width + first_val_width + second_info_width,
                                        self._current_row,
                                        start_col + first_info_width + first_val_width + second_info_width + second_val_width - 1,
                                        "", self._reelsets_spin_win_formats['value_cell_' + str(section_index)])
                self._sheet.write_rich_string(self._current_row, start_col + first_info_width + first_val_width + second_info_width,
                                              self._reelsets_spin_win_formats['value_cell_' + str(section_index)],
                                              self._section_names[section_index],
                                              self._reelsets_spin_win_formats['value_cell_' + str(section_index)],
                                              ' ',
                                              self._reelsets_spin_win_formats['value_cell_' + str(section_index)])
                if reelset_index > 0:
                    self._sheet.merge_range(self._current_row,
                                            start_col + first_info_width + first_val_width + second_info_width + second_val_width,
                                            self._current_row + 1,
                                            start_col + first_info_width + first_val_width + second_info_width + second_val_width + button_width - 1,
                                            '', self._content_formats['content_button_'+str(section_index)])
                    link_to = "internal:'" + self._sheet_name + "'!" + xl_rowcol_to_cell(self._global_content_row_col[0], 1)
                    self._sheet.write_url(self._current_row,
                                          start_col + first_info_width + first_val_width + second_info_width + second_val_width,
                                          link_to,
                                          string='Go To\nContent',
                                          tip='Click Here for Quick Jump',
                                          cell_format=self._content_formats['content_button_'+str(section_index)])
                self._current_row += 1

                self._sheet.merge_range(self._current_row, start_col,
                                        self._current_row, start_col + first_info_width - 1,
                                        "", self._reelsets_spin_win_formats['info_cell_' + str(section_index)])
                self._sheet.write_rich_string(self._current_row, start_col,
                                              self._reelsets_spin_win_formats['info_cell_' + str(section_index)],
                                              'Reelset Index:',
                                              self._reelsets_spin_win_formats['info_cell_' + str(section_index)],
                                              ' ',
                                              self._reelsets_spin_win_formats['info_cell_' + str(section_index)])
                self._sheet.write(self._current_row, start_col + first_info_width,
                                  str(reelset_index),
                                  self._reelsets_spin_win_formats['value_cell_' + str(section_index)])

                self._sheet.merge_range(self._current_row, start_col + first_info_width + first_val_width,
                                        self._current_row,
                                        start_col + first_info_width + first_val_width + second_info_width - 1,
                                        "", self._reelsets_spin_win_formats['info_cell_' + str(section_index)])
                self._sheet.write_rich_string(self._current_row, start_col + first_info_width + first_val_width,
                                              self._reelsets_spin_win_formats['info_cell_' + str(section_index)],
                                              'Reelset Name:',
                                              self._reelsets_spin_win_formats['info_cell_' + str(section_index)],
                                              ' ',
                                              self._reelsets_spin_win_formats['info_cell_' + str(section_index)])

                self._sheet.merge_range(self._current_row,
                                        start_col + first_info_width + first_val_width + second_info_width,
                                        self._current_row,
                                        start_col + first_info_width + first_val_width + second_info_width + second_val_width - 1,
                                        "", self._reelsets_spin_win_formats['value_cell_' + str(section_index)])
                self._sheet.write_rich_string(self._current_row,
                                              start_col + first_info_width + first_val_width + second_info_width,
                                              self._reelsets_spin_win_formats['value_cell_' + str(section_index)],
                                              self._all_reelsets_section_dfs[section_index]['reelset_name'].iloc[reelset_index],
                                              self._reelsets_spin_win_formats['value_cell_' + str(section_index)],
                                              ' ',
                                              self._reelsets_spin_win_formats['value_cell_' + str(section_index)])
                self._current_row += 1

                reelset_name = 'reelset_' + str(section_index)+'_'+str(reelset_index)
                self._WriteSpinWinDistributionByRanges(reelset, reelset_name, section_index)
                self._current_row += 1
        self._current_row += 2
        self._WriteContentReelsetsSpinWins()

    def _AllocateSpaceReelsetsSpinWins(self):
        self.WriteInnerHeader('Reelsets Statistics Content', include_in_content=False, print_counter=False)
        self._reelsets_spin_win_content_row_col = [self._current_row, 1]
        sections_count = len(self._reelsets)
        reelsets_count = sum([len(reelsets) for reelsets in self._reelsets.GetReelsets()])

        self._current_row += sections_count + reelsets_count + 3

    def _WriteContentReelsetsSpinWins(self):
        content_row = self._reelsets_spin_win_content_row_col[0]

        big_start_col = 1
        small_start_col = 2
        big_width = 12
        small_width = 11

        for big_index, (big_text, big_row) in enumerate(self._reelsets_spin_win_section_content):
            self._sheet.set_row_pixels(content_row, 40)
            self._sheet.merge_range(content_row, big_start_col,
                                    content_row, big_start_col+big_width-1,
                                    '', self._content_formats['big'])
            link_to = "internal:'" + self._sheet_name + "'!" + xl_rowcol_to_cell(big_row, 1)
            self._sheet.write_url(content_row, big_start_col, link_to, string=big_text, cell_format=self._content_formats['big'], tip='Click Here for Quick Jump')
            content_row += 1
            for small_index, (small_text, small_row) in enumerate(self._reelsets_spin_win_reelsets_content[big_index]):
                self._sheet.set_row_pixels(content_row, 30)
                self._sheet.merge_range(content_row, small_start_col,
                                        content_row, small_start_col + small_width - 1,
                                        '', self._content_formats['small'])
                link_to = "internal:'" + self._sheet_name + "'!" + xl_rowcol_to_cell(small_row, 1)
                self._sheet.write_url(content_row, small_start_col, link_to, string=small_text, cell_format=self._content_formats['small'], tip='Click Here for Quick Jump')
                content_row += 1


    def _WriteReelsetsStructure(self):
        self.WriteHeader("Reelsets Structure", print_counter=True, include_in_content=True)
        self._AllocateSpaceReelsetsStructure()

        start_col = 1
        first_info_width = 2
        first_val_width = 1
        second_info_width = 2
        second_val_width = 7
        third_info_width = 3
        third_val_width = 3
        position_width = 2
        symbol_width = 3
        weight_width = 1
        reel_width = symbol_width + weight_width
        button_width = 2

        for section_index, reelsets in enumerate(self._reelsets.GetReelsets()):
            text = self._section_names[section_index] + " Reelsets"
            self._reelsets_structure_section_content.append((str(section_index + 1) + '. ' + text, self._current_row))
            self.WriteInnerHeader(text, print_counter=True, include_in_content=True)

            info_format = self._reelsets_structure_formats['top_info_'+str(section_index)]
            val_format = self._reelsets_structure_formats['top_value_'+str(section_index)]
            top_description_format = self._reelsets_structure_formats['top_description_'+str(section_index)]
            bottom_description_format = self._reelsets_structure_formats['bottom_description_'+str(section_index)]

            for reelset_index, reelset in enumerate(reelsets):
                text = str(section_index+1)+'.'+str(reelset_index)+". Reelset: '"+self._reelsets.GetReelsetName(section_index, reelset_index)+"' structure"
                self._reelsets_structure_reelsets_content[section_index].append((text, self._current_row))
                self._sheet.merge_range(self._current_row, start_col,
                                        self._current_row, start_col + first_info_width - 1,
                                        "", info_format)
                self._sheet.write_rich_string(self._current_row, start_col,
                                              self._text_formats['text_regular'],
                                              'Section Index:',
                                              self._text_formats['text_regular'],
                                              ' ',
                                              info_format)
                self._sheet.write_rich_string(self._current_row, start_col+first_info_width,
                                              self._text_formats['text_regular_bold'], str(section_index),
                                              self._text_formats['text_regular_bold'], ' ',
                                              val_format)

                self._sheet.merge_range(self._current_row, start_col + first_info_width + first_val_width,
                                        self._current_row,
                                        start_col + first_info_width + first_val_width + second_info_width - 1,
                                        "", info_format)
                self._sheet.write_rich_string(self._current_row, start_col + first_info_width + first_val_width,
                                              self._text_formats['text_regular'],
                                              'Section Name:',
                                              self._text_formats['text_regular'],
                                              ' ', info_format)

                self._sheet.merge_range(self._current_row,
                                        start_col + first_info_width + first_val_width + second_info_width,
                                        self._current_row,
                                        start_col + first_info_width + first_val_width + second_info_width + second_val_width - 1,
                                        "", val_format)
                self._sheet.write_rich_string(self._current_row,
                                              start_col + first_info_width + first_val_width + second_info_width,
                                              self._text_formats['text_regular_bold'],
                                              self._section_names[section_index],
                                              self._text_formats['text_regular_bold'],
                                              ' ', val_format)
                self._sheet.merge_range(self._current_row,
                                        start_col+first_info_width+first_val_width+second_info_width+second_val_width,
                                        self._current_row,
                                        start_col+first_info_width+first_val_width+second_info_width+second_val_width+third_info_width-1,
                                        '', info_format)
                self._sheet.write_rich_string(self._current_row,
                                              start_col+first_info_width+first_val_width+second_info_width+second_val_width,
                                              self._text_formats['text_regular'], 'Total Section Weight:',
                                              self._text_formats['text_regular'], ' ', info_format)

                self._sheet.merge_range(self._current_row,
                                        start_col + first_info_width + first_val_width + second_info_width + second_val_width + third_info_width,
                                        self._current_row,
                                        start_col + first_info_width + first_val_width + second_info_width + second_val_width + third_info_width + third_val_width - 1,
                                        '', val_format)
                self._sheet.write_rich_string(self._current_row,
                                              start_col + first_info_width + first_val_width + second_info_width + second_val_width + third_info_width,
                                              self._text_formats['text_regular_bold'], '{:,d}'.format(self._all_reelsets_section_dfs[section_index]['reelset_weight'].sum()),
                                              self._text_formats['text_regular_bold'], ' ', val_format)
                if reelset_index > 0:
                    self._sheet.merge_range(self._current_row,
                                            start_col + first_info_width + first_val_width + second_info_width + second_val_width + third_info_width + third_val_width,
                                            self._current_row + 1,
                                            start_col + first_info_width + first_val_width + second_info_width + second_val_width + third_info_width + third_val_width + button_width - 1,
                                            '', self._content_formats['content_button_' + str(section_index)])
                    link_to = "internal:'" + self._sheet_name + "'!" + xl_rowcol_to_cell(self._global_content_row_col[0], 1)
                    self._sheet.write_url(self._current_row,
                                          start_col + first_info_width + first_val_width + second_info_width + second_val_width + third_info_width + third_val_width,
                                          link_to,
                                          string='Go To\nContent',
                                          tip='Click Here for Quick Jump',
                                          cell_format=self._content_formats['content_button_' + str(section_index)])

                self._current_row += 1

                self._sheet.merge_range(self._current_row, start_col,
                                        self._current_row, start_col + first_info_width - 1,
                                        "", info_format)
                self._sheet.write_rich_string(self._current_row, start_col,
                                              self._text_formats['text_regular'],
                                              'Reelset Index:',
                                              self._text_formats['text_regular'],
                                              ' ', info_format)
                self._sheet.write(self._current_row, start_col + first_info_width,
                                  str(reelset_index), val_format)

                self._sheet.merge_range(self._current_row, start_col + first_info_width + first_val_width,
                                        self._current_row,
                                        start_col + first_info_width + first_val_width + second_info_width - 1,
                                        "", info_format)
                self._sheet.write_rich_string(self._current_row, start_col + first_info_width + first_val_width,
                                              self._text_formats['text_regular'],
                                              'Reelset Name:',
                                              self._text_formats['text_regular'],
                                              ' ', info_format)

                self._sheet.merge_range(self._current_row,
                                        start_col + first_info_width + first_val_width + second_info_width,
                                        self._current_row,
                                        start_col + first_info_width + first_val_width + second_info_width + second_val_width - 1,
                                        "", val_format)
                self._sheet.write_rich_string(self._current_row,
                                              start_col + first_info_width + first_val_width + second_info_width,
                                              self._text_formats['text_regular_bold'],
                                              self._all_reelsets_section_dfs[section_index]['reelset_name'].iloc[
                                                  reelset_index],
                                              self._text_formats['text_regular_bold'],
                                              ' ',
                                              val_format)
                self._sheet.merge_range(self._current_row,
                                        start_col + first_info_width + first_val_width + second_info_width+second_val_width,
                                        self._current_row,
                                        start_col + first_info_width + first_val_width + second_info_width + second_val_width + third_info_width - 1,
                                        "", info_format)
                self._sheet.write_rich_string(self._current_row,
                                              start_col + first_info_width + first_val_width + second_info_width+second_val_width,
                                              self._text_formats['text_regular'], 'Reelset Range:',
                                              self._text_formats['text_regular'],
                                              ' ',
                                              info_format)
                self._sheet.merge_range(self._current_row,
                                        start_col + first_info_width + first_val_width + second_info_width + second_val_width + third_info_width,
                                        self._current_row,
                                        start_col + first_info_width + first_val_width + second_info_width + second_val_width + third_info_width +third_val_width- 1,
                                        "", val_format)
                reelset_range = self._reelsets.GetReelsetRange(section_index, reelset_index)
                range_str = '[' + '{:,d}'.format(reelset_range[0])+' - '+'{:,d}'.format(reelset_range[1])+']'
                self._sheet.write_rich_string(self._current_row,
                                              start_col + first_info_width + first_val_width + second_info_width + second_val_width+third_info_width,
                                              self._text_formats['text_regular_bold'], range_str,
                                              self._text_formats['text_regular_bold'],
                                              ' ',
                                              val_format)
                self._current_row += 1

                cur_col = start_col
                self._sheet.merge_range(self._current_row, cur_col,
                                        self._current_row+1, cur_col + position_width - 1,
                                        '', top_description_format)
                self._sheet.write_rich_string(self._current_row, cur_col,
                                              self._text_formats['text_regular_bold'], 'Position Index',
                                              self._text_formats['text_regular_bold'], ' ',
                                              top_description_format)
                cur_col += position_width

                for reel_index in range(len(reelset)):
                    self._sheet.merge_range(self._current_row, cur_col,
                                            self._current_row, cur_col + reel_width - 1,
                                            '', top_description_format)
                    self._sheet.write_rich_string(self._current_row, cur_col,
                                                  self._text_formats['text_big_bold'], 'Reel ',
                                                  self._text_formats['text_big_bold'], str(reel_index),
                                                  top_description_format)
                    cur_col += reel_width

                self._current_row += 1
                cur_col = start_col+position_width

                for reel_index in range(len(reelset)):
                    self._sheet.merge_range(self._current_row, cur_col,
                                            self._current_row, cur_col + symbol_width-1,
                                            '', bottom_description_format)
                    self._sheet.write_rich_string(self._current_row, cur_col,
                                                  self._text_formats['text_regular_bold'], 'Symbols',
                                                  self._text_formats['text_regular_bold'], ' ',
                                                  bottom_description_format)
                    cur_col += symbol_width
                    if weight_width > 1:
                        self._sheet.merge_range(self._current_row, cur_col,
                                                self._current_row, cur_col+weight_width-1,
                                                '', bottom_description_format)
                    self._sheet.write_rich_string(self._current_row, cur_col,
                                                  self._text_formats['text_small_bold'], 'Weights',
                                                  self._text_formats['text_small_bold'], ' ',
                                                  bottom_description_format)
                    cur_col += weight_width

                self._current_row += 1
                cur_col = start_col

                max_reel_length = 0
                for reel in reelset.GetReels():
                    if reel.Length() > max_reel_length:
                        max_reel_length = reel.Length()

                for i in range(max_reel_length):
                    cur_format = self._reelsets_structure_formats['even_position_'+str(section_index)]
                    if i % 2 == 1 and i == max_reel_length-1:
                        cur_format = self._reelsets_structure_formats['odd_position_end_' + str(section_index)]
                    elif i % 2 == 0 and i == max_reel_length-1:
                        cur_format = self._reelsets_structure_formats['even_position_end_' + str(section_index)]
                    elif i % 2 == 1:
                        cur_format = self._reelsets_structure_formats['odd_position_' + str(section_index)]

                    self._sheet.merge_range(self._current_row, cur_col,
                                            self._current_row, cur_col + position_width-1,
                                            '', cur_format)
                    self._sheet.write_rich_string(self._current_row, cur_col,
                                                  self._text_formats['text_regular'], str(i),
                                                  self._text_formats['text_regular'], ' ',
                                                  cur_format)
                    cur_col += position_width

                    for reel_index in range(len(reelset)):
                        cur_format_name = self._par_formats.GetReelsetStructureFormatName(section_index, not i % 2, i == max_reel_length-1, True)
                        cur_format = self._reelsets_structure_formats[cur_format_name]
                        self._sheet.merge_range(self._current_row, cur_col,
                                                self._current_row, cur_col+symbol_width-1,
                                                '', cur_format)

                        symbol_id = -1
                        symbol_name = ' '
                        symbol_weight = ' '
                        if i < reelset.GetReels(reel_index).Length():
                            symbol_id = reelset.GetReels(reel_index).GetSymbols()[i]
                            symbol_name = self._paytable.GetSymbolName(symbol_id)
                            symbol_weight = str(reelset.GetReels(reel_index).GetWeights()[i])

                        self._sheet.write_rich_string(self._current_row, cur_col,
                                                      self._text_formats['text_regular'], symbol_name,
                                                      self._text_formats['text_regular'], ' ',
                                                      cur_format)
                        cur_col += symbol_width

                        cur_format_name = self._par_formats.GetReelsetStructureFormatName(section_index, not i % 2,
                                                                                          i == max_reel_length - 1,
                                                                                          False)
                        cur_format = self._reelsets_structure_formats[cur_format_name]
                        if weight_width > 1:
                            self._sheet.merge_range(self._current_row, cur_col,
                                                    self._current_row, cur_col+weight_width-1,
                                                    '', cur_format)
                        self._sheet.write_rich_string(self._current_row, cur_col,
                                                      self._text_formats['text_regular'], symbol_weight,
                                                      self._text_formats['text_regular'], ' ',
                                                      cur_format)
                        cur_col += weight_width
                    cur_col = start_col
                    self._current_row += 1
                self._current_row += 2
            self._current_row += 1
        self._WriteContentReelsetsStructure()

    def _AllocateSpaceReelsetsStructure(self):
        self.WriteInnerHeader("Reelsets Structure Content", include_in_content=False, print_counter=False)
        self._reelsets_structure_content_row_col = [self._current_row, 1]
        section_count = len(self._reelsets)
        reelsets_count = sum([len(reelsets) for reelsets in self._reelsets.GetReelsets()])
        self._current_row += section_count + reelsets_count + 3

    def _WriteContentReelsetsStructure(self):
        content_row = self._reelsets_structure_content_row_col[0]
        big_start = 1
        small_start = 2
        big_width = 12
        small_width = 11

        for big_index, (big_text, big_row) in enumerate(self._reelsets_structure_section_content):
            self._sheet.set_row_pixels(content_row, 40)
            self._sheet.merge_range(content_row, big_start,
                                    content_row, big_start+big_width-1,
                                    '', self._content_formats['big'])
            link_to = "internal:'" + self._sheet_name + "'!" + xl_rowcol_to_cell(big_row, 1)
            self._sheet.write_url(content_row, big_start, link_to, string=big_text, cell_format=self._content_formats['big'], tip='Click Here for Quick Jump')
            content_row += 1
            for small_index, (small_text, small_row) in enumerate(self._reelsets_structure_reelsets_content[big_index]):
                self._sheet.set_row_pixels(content_row, 30)
                self._sheet.merge_range(content_row, small_start,
                                        content_row, small_start + small_width - 1,
                                        '', self._content_formats['small'])
                link_to = "internal:'" + self._sheet_name + "'!" + xl_rowcol_to_cell(small_row, 1)
                self._sheet.write_url(content_row, small_start, link_to, string=small_text, cell_format=self._content_formats['small'], tip='Click Here for Quick Jump')
                content_row += 1

    def WriteBasicSpinWin(self, spin_win_df: pd.DataFrame,
                          graph_name: str,
                          color_index: int = 0,
                          print_avg_win_no_zero: bool = True,
                          print_avg_win_with_zero: bool = False,
                          print_win_1_in_feature: bool = True,
                          print_win_1_in_base: bool = False,
                          print_win_freq_base_col: bool = True,
                          print_max_range_win_col: bool = True):

        start_col = 1
        info_table_info_width = 4
        info_table_val_width = 3
        info_table_start_col = self.GetCentreStartCol(12, info_table_info_width+info_table_val_width, 1)

        ranges_width = 2
        win_freq_feature_width = 2
        win_freq_base_width = 2
        avg_win_width = 2
        max_win_width = 2
        rtp_small_width = 2
        rtp_big_width = 2
        rtp_picture_width = 5

        border_description_format = self._rtp_distribution_formats['border_description_' + str(color_index)]
        border_even_format = self._rtp_distribution_formats['border_even_' + str(color_index)]
        border_odd_format = self._rtp_distribution_formats['border_odd_' + str(color_index)]

        # WRITE INFO TABLE

        info_table_rows_counter = 0
        if print_avg_win_with_zero:
            cur_format = border_even_format if info_table_rows_counter % 2 == 0 else border_odd_format
            self._sheet.set_row_pixels(self._current_row, 45)
            self._sheet.merge_range(self._current_row, info_table_start_col,
                                    self._current_row, info_table_start_col+info_table_info_width-1,
                                    "", cur_format)
            self._sheet.write_rich_string(self._current_row, info_table_start_col,
                                          self._text_formats['text_regular_bold'], "Average Win:\n",
                                          self._text_formats['text_smallest'], "(with zero wins)",
                                          cur_format)
            self._sheet.merge_range(self._current_row, info_table_start_col+info_table_info_width,
                                    self._current_row, info_table_start_col+info_table_info_width+info_table_val_width-1,
                                    "", cur_format)
            avg_win_with_zero = spin_win_df['avg_win_with_zero'].iloc[0]
            self._sheet.write_rich_string(self._current_row,
                                          info_table_start_col+info_table_info_width,
                                          self._text_formats['text_big'], "{:,.2f}x".format(avg_win_with_zero/self._base_bet),
                                          self._text_formats['text_smallest'], ' bets',
                                          cur_format)
            info_table_rows_counter += 1
            self._current_row += 1
        if print_avg_win_no_zero:
            cur_format = border_even_format if info_table_rows_counter % 2 == 0 else border_odd_format
            self._sheet.set_row_pixels(self._current_row, 45)
            self._sheet.merge_range(self._current_row, info_table_start_col,
                                    self._current_row, info_table_start_col+info_table_info_width-1,
                                    "", cur_format)
            self._sheet.write_rich_string(self._current_row, info_table_start_col,
                                          self._text_formats['text_regular_bold'], "Average Win:\n",
                                          self._text_formats['text_smallest'], "(no zero wins)",
                                          cur_format)
            self._sheet.merge_range(self._current_row, info_table_start_col+info_table_info_width,
                                    self._current_row, info_table_start_col+info_table_info_width+info_table_val_width-1,
                                    "", cur_format)
            avg_win_no_zero = spin_win_df['avg_win_no_zero'].iloc[0]
            self._sheet.write_rich_string(self._current_row,
                                          info_table_start_col+info_table_info_width,
                                          self._text_formats['text_big'], "{:,.2f}x".format(avg_win_no_zero/self._base_bet),
                                          self._text_formats['text_smallest'], ' bets',
                                          cur_format)
            info_table_rows_counter += 1
            self._current_row += 1
        if print_win_1_in_feature:
            cur_format = border_even_format if info_table_rows_counter%2 == 0 else border_odd_format
            self._sheet.set_row_pixels(self._current_row, 45)
            self._sheet.merge_range(self._current_row, info_table_start_col,
                                    self._current_row, info_table_start_col+info_table_info_width-1,
                                    "", cur_format)
            self._sheet.write_rich_string(self._current_row, info_table_start_col,
                                          self._text_formats['text_regular_bold'], "Win Frequency:\n",
                                          self._text_formats['text_smallest'], "(1 in .. feature spins)",
                                          cur_format)
            self._sheet.merge_range(self._current_row, info_table_start_col+info_table_info_width,
                                    self._current_row, info_table_start_col+info_table_info_width+info_table_val_width-1,
                                    "", cur_format)
            win_1_in_feature = spin_win_df['win_1_in_total_feature'].iloc[0]
            self._sheet.write_rich_string(self._current_row,
                                          info_table_start_col+info_table_info_width,
                                          self._text_formats['text_smallest'], '1 in ',
                                          self._text_formats['text_big'], "{:,.2f}".format(win_1_in_feature),
                                          cur_format)
            info_table_rows_counter += 1
            self._current_row += 1
        if print_win_1_in_base:
            cur_format = border_even_format if info_table_rows_counter%2 == 0 else border_odd_format
            self._sheet.set_row_pixels(self._current_row, 45)
            self._sheet.merge_range(self._current_row, info_table_start_col,
                                    self._current_row, info_table_start_col+info_table_info_width-1,
                                    "", cur_format)
            self._sheet.write_rich_string(self._current_row, info_table_start_col,
                                          self._text_formats['text_regular_bold'], "Win Frequency:\n",
                                          self._text_formats['text_smallest'], "(1 in .. base spins)",
                                          cur_format)
            self._sheet.merge_range(self._current_row, info_table_start_col+info_table_info_width,
                                    self._current_row, info_table_start_col+info_table_info_width+info_table_val_width-1,
                                    "", cur_format)
            win_1_in_feature = spin_win_df['win_1_in_total_base'].iloc[0]
            self._sheet.write_rich_string(self._current_row,
                                          info_table_start_col+info_table_info_width,
                                          self._text_formats['text_smallest'], '1 in ',
                                          self._text_formats['text_big'], "{:,.2f}".format(win_1_in_feature),
                                          cur_format)
            info_table_rows_counter += 1
            self._current_row += 1
        self._current_row += 1

        # WRITE RANGES TABLE
        cur_col = start_col
        self._sheet.set_row_pixels(self._current_row, 45)
        self._sheet.merge_range(self._current_row, start_col,
                                self._current_row, start_col + ranges_width - 1,
                                "", border_description_format)
        self._sheet.write_rich_string(self._current_row, start_col,
                                      self._text_formats['text_big_bold'], "Win ",
                                      self._text_formats['text_big_bold'], "Ranges",
                                      border_description_format)

        cur_col += ranges_width

        self._sheet.merge_range(self._current_row, cur_col,
                                self._current_row, cur_col + win_freq_feature_width - 1,
                                "", border_description_format)
        self._sheet.write_rich_string(self._current_row, start_col + ranges_width,
                                      self._text_formats['text_big_bold'], "Win in Range\n",
                                      self._text_formats['text_smallest'], "(1 in ... feature spins)",
                                      border_description_format)

        cur_col += win_freq_feature_width

        if print_win_freq_base_col:
            self._sheet.merge_range(self._current_row, cur_col,
                                    self._current_row, cur_col + win_freq_base_width - 1,
                                    "", border_description_format)
            self._sheet.write_rich_string(self._current_row, cur_col,
                                          self._text_formats['text_big_bold'], "Win in Range\n",
                                          self._text_formats['text_smallest'], "(1 in ... base spins)",
                                          border_description_format)
            cur_col += win_freq_base_width

        self._sheet.merge_range(self._current_row, cur_col,
                                self._current_row, cur_col + avg_win_width - 1,
                                "", border_description_format)
        self._sheet.write_rich_string(self._current_row, cur_col,
                                      self._text_formats["text_big_bold"], "Average Win\n",
                                      self._text_formats["text_smallest"], "(in range)",
                                      border_description_format)

        cur_col += avg_win_width

        if print_max_range_win_col:
            self._sheet.merge_range(self._current_row, cur_col,
                                    self._current_row, cur_col + max_win_width - 1,
                                    "", border_description_format)
            self._sheet.write_rich_string(self._current_row, cur_col,
                                          self._text_formats["text_big_bold"], "Max Win\n",
                                          self._text_formats["text_smallest"], "(in range)",
                                          border_description_format)

            cur_col += max_win_width

        self._sheet.merge_range(self._current_row, cur_col,
                                self._current_row,
                                cur_col + rtp_small_width - 1,
                                "", border_description_format)
        self._sheet.write_rich_string(self._current_row, cur_col,
                                      self._text_formats['text_big_bold'], "RTP\n",
                                      self._text_formats['text_smallest'], "(from whole RTP)",
                                      border_description_format)

        cur_col += rtp_small_width

        self._sheet.merge_range(self._current_row, cur_col,
                                self._current_row,
                                cur_col + rtp_big_width + rtp_picture_width - 1,
                                "", border_description_format)
        self._sheet.write_rich_string(self._current_row, cur_col,
                                      self._text_formats['text_big_bold'], "RTP\n",
                                      self._text_formats['text_smallest'], "(from common RTP)",
                                      border_description_format)

        cur_col += rtp_big_width
        self._current_row += 1

        self._graphMaker.MakeBarPlot_SimpleRTPDistribution(spin_win_df, graph_name, color_index)
        self._sheet.insert_image(self._current_row,
                                 cur_col,
                                 self._graphMaker.GetSpinWinRTPPath(graph_name),
                                 {
                                     'x_offset': 1,
                                     'y_offset': 1,
                                     'x_scale': 0.553,
                                     'y_scale': 0.416,
                                     'object_position': 2
                                 })

        for i in range(len(spin_win_df.index)):
            cur_col = start_col
            self._sheet.set_row_pixels(self._current_row, 30)
            current_border_format = border_even_format if i % 2 == 0 else border_odd_format

            self._sheet.merge_range(self._current_row, cur_col,
                                    self._current_row, cur_col + ranges_width - 1,
                                    "", current_border_format)
            self._sheet.write_rich_string(self._current_row, cur_col,
                                          self._text_formats['text_regular_bold'], spin_win_df.index[i][:1],
                                          self._text_formats['text_regular_bold'], spin_win_df.index[i][1:],
                                          current_border_format)

            cur_col += ranges_width

            self._sheet.merge_range(self._current_row, cur_col,
                                    self._current_row, cur_col + win_freq_feature_width - 1,
                                    "", current_border_format)
            cur_win_1_in = spin_win_df['win_1_in'].iloc[i]
            self._sheet.write_rich_string(self._current_row, cur_col,
                                          self._text_formats['text_smallest'], "1 in ",
                                          self._text_formats['text_regular'], '{:,.2f}'.format(cur_win_1_in),
                                          current_border_format)

            cur_col += win_freq_feature_width

            if print_win_freq_base_col:
                self._sheet.merge_range(self._current_row, cur_col,
                                        self._current_row, cur_col + win_freq_base_width - 1,
                                        "", current_border_format)
                cur_win_1_in_base = spin_win_df['win_1_in_small'].iloc[i]
                self._sheet.write_rich_string(self._current_row, cur_col,
                                              self._text_formats['text_smallest'], "1 in ",
                                              self._text_formats['text_regular'], '{:,.2f}'.format(cur_win_1_in_base),
                                              current_border_format)
                cur_col += win_freq_base_width

            self._sheet.merge_range(self._current_row, cur_col,
                                    self._current_row, cur_col + avg_win_width - 1,
                                    "", current_border_format)
            cur_avg_win = spin_win_df['avg_win'].iloc[i] / self._base_bet
            self._sheet.write_rich_string(self._current_row, cur_col,
                                          self._text_formats['text_regular'], '{:.2f}'.format(cur_avg_win),
                                          self._text_formats['text_regular'], 'x' if not pd.isna(cur_avg_win) else " ",
                                          self._text_formats['text_smallest'], ' bets' if not pd.isna(cur_avg_win) else " ",
                                          current_border_format)
            cur_col += avg_win_width

            if print_max_range_win_col:
                self._sheet.merge_range(self._current_row, cur_col,
                                        self._current_row, cur_col + max_win_width - 1,
                                        "", current_border_format)
                cur_max_win = spin_win_df['max_win'].iloc[i] / self._base_bet
                self._sheet.write_rich_string(self._current_row, cur_col,
                                              self._text_formats['text_regular'], '{:,.2f}'.format(cur_max_win),
                                              self._text_formats['text_regular'],
                                              'x' if not pd.isna(cur_avg_win) else " ",
                                              self._text_formats['text_smallest'],
                                              ' bets' if not pd.isna(cur_avg_win) else " ",
                                              current_border_format)
                cur_col += avg_win_width

            self._sheet.merge_range(self._current_row, cur_col,
                                    self._current_row, cur_col + rtp_small_width - 1,
                                    "", current_border_format)
            cur_small_rtp = spin_win_df['rtp'].iloc[i]
            self._sheet.write_rich_string(self._current_row, cur_col,
                                          self._text_formats['text_regular'], '{:,.3f}%'.format(cur_small_rtp),
                                          self._text_formats['text_regular'], " ",
                                          current_border_format)
            cur_col += rtp_small_width

            self._sheet.merge_range(self._current_row, cur_col,
                                    self._current_row, cur_col + rtp_big_width - 1,
                                    "", current_border_format)
            cur_big_rtp = spin_win_df['rtp_big'].iloc[i]
            self._sheet.write_rich_string(self._current_row, cur_col,
                                          self._text_formats['text_regular'], '{:,.3f}%'.format(cur_big_rtp),
                                          self._text_formats['text_regular'], " ",
                                          current_border_format)
            cur_col += rtp_big_width
            self._current_row += 1

        cur_col = start_col
        self._sheet.set_row_pixels(self._current_row, 30)
        self._sheet.merge_range(self._current_row, cur_col,
                                self._current_row, cur_col + ranges_width - 1,
                                "", border_description_format)
        self._sheet.write_rich_string(self._current_row, cur_col,
                                      self._text_formats['text_big_bold'], "Tota",
                                      self._text_formats['text_big_bold'], "l",
                                      border_description_format)

        cur_col += ranges_width

        self._sheet.merge_range(self._current_row, cur_col,
                                self._current_row, cur_col + win_freq_feature_width - 1,
                                "", border_description_format)
        total_win_freq = np.sum(spin_win_df['total_counter']) / (
                    np.sum(spin_win_df['total_counter']) - spin_win_df['total_counter'].iloc[0])
        self._sheet.write_rich_string(self._current_row, cur_col,
                                      self._text_formats['text_smallest'], '1 in ',
                                      self._text_formats['text_regular_bold'], '{:.2f}'.format(total_win_freq),
                                      border_description_format)

        cur_col += win_freq_feature_width

        if print_win_freq_base_col:
            self._sheet.merge_range(self._current_row, cur_col,
                                    self._current_row, cur_col + win_freq_base_width - 1,
                                    "", border_description_format)
            total_win_freq = self._spin_count / (np.sum(spin_win_df['total_counter']) - spin_win_df['total_counter'].iloc[0])
            self._sheet.write_rich_string(self._current_row, cur_col,
                                          self._text_formats['text_smallest'], '1 in ',
                                          self._text_formats['text_regular_bold'], '{:,.2f}'.format(total_win_freq),
                                          border_description_format)
            cur_col += win_freq_base_width

        self._sheet.merge_range(self._current_row, cur_col,
                                self._current_row, cur_col + avg_win_width - 1,
                                "", border_description_format)
        total_avg_win = np.sum(spin_win_df['total_win']) / np.sum(
            spin_win_df['total_counter'].iloc[1:]) / self._base_bet
        self._sheet.write_rich_string(self._current_row, cur_col,
                                      self._text_formats['text_regular_bold'], '{:,.2f}'.format(total_avg_win),
                                      self._text_formats['text_regular_bold'], "x ",
                                      self._text_formats['text_smallest'], ' bets',
                                      border_description_format)
        cur_col += avg_win_width

        if print_max_range_win_col:
            self._sheet.merge_range(self._current_row, cur_col,
                                    self._current_row, cur_col + max_win_width - 1,
                                    "", border_description_format)
            total_max_win = spin_win_df['max_win'].max() / self._base_bet
            self._sheet.write_rich_string(self._current_row, cur_col,
                                          self._text_formats['text_regular_bold'], '{:,.2f}'.format(total_max_win),
                                          self._text_formats['text_regular_bold'], "x",
                                          self._text_formats['text_smallest'], " bets",
                                          border_description_format)
            cur_col += max_win_width

        self._sheet.merge_range(self._current_row, cur_col,
                                self._current_row, cur_col + rtp_small_width - 1,
                                "", border_description_format)
        total_feature_rtp = spin_win_df['rtp'].sum()
        self._sheet.write_rich_string(self._current_row, cur_col,
                                      self._text_formats['text_regular_bold'], '{:,.2f}'.format(total_feature_rtp),
                                      self._text_formats['text_regular_bold'], "%",
                                      border_description_format)
        cur_col += rtp_small_width

        self._sheet.merge_range(self._current_row, cur_col,
                                self._current_row, cur_col + rtp_big_width - 1,
                                "", border_description_format)
        common_feature_rtp = spin_win_df['rtp_big'].sum()
        self._sheet.write_rich_string(self._current_row, cur_col,
                                      self._text_formats['text_regular_bold'], '{:,.2f}'.format(common_feature_rtp),
                                      self._text_formats['text_regular_bold'], "%",
                                      border_description_format)
        cur_col += rtp_big_width
        self._current_row += 4

    def WriteBasicVariant(self,
                          main_df: pd.DataFrame,
                          merged_df: pd.DataFrame,
                          variant_df_col_name: str,
                          txt_labels: dict,
                          variant_is_mult: bool = False,
                          color_index: int = 0,
                          n: tuple = (-1),
                          include_zero_in_total: bool = False,
                          print_n_plus_1_in_feature: bool = False,
                          print_n_plus_1_in_base: bool = False,
                          print_avg_variant_with_zero: bool = False,
                          print_avg_variant_no_zero: bool = False,
                          print_any_variant_1_in_base_no_zero: bool = False,
                          print_any_variant_1_in_base_with_zero: bool = False,
                          print_any_variant_1_in_feature_no_zero: bool = False,
                          print_pulls_to_hit_big_col: bool = False,
                          print_pulls_to_hit_small_col: bool = False,
                          print_percent_big_col: bool = False,
                          print_percent_small_col: bool = False):
        start_col = 1
        info_table_info_width = 5
        info_table_val_width = 2
        info_table_start_col = self.GetCentreStartCol(12, info_table_val_width+info_table_info_width, default_start_col=start_col)

        picture_width = 8
        variant_width = 2

        pulls_to_hit_big_width = 2
        pulls_to_hit_big_merged_width = 2

        pulls_to_hit_small_width = 2
        pulls_to_hit_small_merged_width = 2

        percent_big_width = 2
        percent_big_merged_width = 2

        percent_small_width = 2
        percent_small_merged_width = 2

        # WRITE INFO TABLE
        info_format = lambda row_counter: self._variant_formats['info_table_even_'+str(color_index)] if info_rows_counter % 2 == 0 else \
                                          self._variant_formats['info_table_odd_'+str(color_index)]
        info_rows_counter = 0

        def printInfoRow(format, txt_label, little_txt, val_prefix, txt_val, val_postfix):
            self._sheet.set_row_pixels(self._current_row, 45)
            self._sheet.merge_range(self._current_row, info_table_start_col,
                                    self._current_row, info_table_start_col + info_table_info_width - 1,
                                    "", format)
            self._sheet.write_rich_string(self._current_row, info_table_start_col,
                                          self._text_formats['text_big_bold'],
                                          txt_label + '\n',
                                          self._text_formats['text_smallest'], little_txt,
                                          format)
            self._sheet.merge_range(self._current_row, info_table_start_col + info_table_info_width,
                                    self._current_row,
                                    info_table_start_col + info_table_info_width + info_table_val_width - 1,
                                    "", format)
            self._sheet.write_rich_string(self._current_row, info_table_start_col + info_table_info_width,
                                          self._text_formats['text_smallest'], val_prefix+' ',
                                          self._text_formats['text_big'], txt_val,
                                          self._text_formats['text_smallest'], ' ' + val_postfix,
                                          format)
            self._current_row += 1

        if print_avg_variant_with_zero:
            printInfoRow(info_format(info_rows_counter),
                         txt_labels['avg_variant_with_zero'],
                         '(zero case included)',
                         '',
                         '{:,.2f}'.format(main_df['avg_val_with_zero'].iloc[0]) + ('x' if variant_is_mult else ''),
                         '' if not variant_is_mult else 'bets')
            info_rows_counter += 1
        if print_avg_variant_no_zero:
            printInfoRow(info_format(info_rows_counter),
                         txt_labels['avg_variant_no_zero'],
                         '(without zero case)',
                         '',
                         '{:,.2f}'.format(main_df['avg_val_without_zero'].iloc[0]) + ('x' if variant_is_mult else ''),
                         '' if not variant_is_mult else 'bets')
            info_rows_counter += 1
        if print_any_variant_1_in_base_with_zero:
            printInfoRow(info_format(info_rows_counter),
                         txt_labels['any_variant_1_in_base_with_zero'],
                         '(zero case included, 1 in ... base spins)',
                         '1 in',
                         '{:,.2f}'.format(main_df['any_variant_1_in_base_with_zero'].iloc[0]) + ('x' if variant_is_mult else ''),
                         '')
            info_rows_counter += 1
        if print_any_variant_1_in_base_no_zero:
            printInfoRow(info_format(info_rows_counter),
                         txt_labels['any_variant_1_in_base_no_zero'],
                         '(without zero case, 1 in ... base spins)',
                         '1 in',
                         '{:,.2f}'.format(main_df['any_variant_1_in_base_no_zero'].iloc[0]) + ('x' if variant_is_mult else ''),
                         '')
            info_rows_counter += 1
        if print_any_variant_1_in_feature_no_zero:
            printInfoRow(info_format(info_rows_counter),
                         txt_labels['any_variant_1_in_feature_no_zero'],
                         '(without zero case, 1 in ... feature spins)',
                         '1 in',
                         '{:,.2f}'.format(main_df['any_variant_1_in_feature_no_zero'].iloc[0]) + ('x' if variant_is_mult else ''),
                         '')
            info_rows_counter += 1
        for i, n_threshold in enumerate(n):
            if n_threshold == -1:
                break

            if print_n_plus_1_in_base:
                val = self._spin_count / np.sum(main_df['count'].loc[n_threshold:])
                printInfoRow(info_format(info_rows_counter),
                             txt_labels['plus_variants_1_in_base'].format(n_threshold),
                             '(1 in ... base spins)',
                             '1 in',
                             '{:,.2f}'.format(val),
                             '')
                info_rows_counter += 1
            if print_n_plus_1_in_feature:
                val = np.sum(main_df['count']) / np.sum(main_df['count'].loc[n_threshold:])
                printInfoRow(info_format(info_rows_counter),
                             txt_labels['plus_variants_1_in_feature'].format(n_threshold),
                             '(1 in ... feature spins)',
                             '1 in',
                             '{:,.2f}'.format(val),
                             '')
                info_rows_counter += 1

        self._current_row += 1

        # WRITE VAL DISTRIBUTION
        def printCommonCell(start_row, start_col,
                            end_row, end_col,
                            bg_format,
                            val_prefix,
                            txt_val,
                            val_postfix,
                            little_txt='',
                            prefix_txt_format=self._text_formats['text_smallest'],
                            val_txt_format=self._text_formats['text_regular'],
                            postfix_txt_format=self._text_formats['text_smallest'],
                            little_txt_format=self._text_formats['text_smallest']):
            self._sheet.merge_range(start_row, start_col, end_row, end_col,
                                    '', bg_format)
            self._sheet.write_rich_string(start_row, start_col,
                                          prefix_txt_format, val_prefix + ' ',
                                          val_txt_format, txt_val,
                                          postfix_txt_format, ' ' + val_postfix,
                                          little_txt_format, ' ' if little_txt == '' else ' \n'+little_txt,
                                          bg_format)
            return end_col + 1

        common_bg_format = lambda row_index, is_merged_col, is_centre_col: \
            ((self._variant_formats['even_left_'+str(color_index)] if row_index % 2 == 0 else \
             self._variant_formats['odd_left_'+str(color_index)]) if not is_centre_col else \
                 (self._variant_formats['even_centre_'+str(color_index)] if row_index % 2 == 0 else \
                  self._variant_formats['odd_centre_'+str(color_index)])) if not is_merged_col else \
                ((self._variant_formats['even_right_'+str(color_index)] if row_index % 2 == 0 else \
                 self._variant_formats['odd_right_'+str(color_index)]) if not is_centre_col else \
                     (self._variant_formats['even_centre_'+str(color_index)] if row_index % 2 == 0 else \
                      self._variant_formats['odd_centre_'+str(color_index)]))

        cur_col = start_col
        self._sheet.set_row_pixels(self._current_row, 45)
        cur_col = printCommonCell(self._current_row, cur_col,
                                  self._current_row, cur_col+picture_width-1,
                                  self._variant_formats['description_'+str(color_index)],
                                  '',
                                  txt_labels['picture_col'],
                                  '',
                                  val_txt_format=self._text_formats['text_big_bold'])
        cur_col = printCommonCell(self._current_row, cur_col,
                                  self._current_row, cur_col + variant_width - 1,
                                  self._variant_formats['description_' + str(color_index)],
                                  '',
                                  txt_labels['variant_col'],
                                  '',
                                  val_txt_format=self._text_formats['text_big_bold'])
        if print_pulls_to_hit_big_col:
            cur_col = printCommonCell(self._current_row, cur_col,
                                      self._current_row, cur_col + pulls_to_hit_big_width - 1,
                                      self._variant_formats['description_' + str(color_index)],
                                      '',
                                      txt_labels['pulls_to_hit_big'],
                                      '',
                                      little_txt='(1 in ... feature spins)',
                                      val_txt_format=self._text_formats['text_big_bold'])
            cur_col = printCommonCell(self._current_row, cur_col,
                                      self._current_row, cur_col + pulls_to_hit_big_merged_width - 1,
                                      self._variant_formats['description_' + str(color_index)],
                                      '',
                                      txt_labels['pulls_to_hit_big_merged'],
                                      '',
                                      little_txt='(1 in ... feature spins)',
                                      val_txt_format=self._text_formats['text_regular_bold'])
        if print_pulls_to_hit_small_col:
            cur_col = printCommonCell(self._current_row, cur_col,
                                      self._current_row, cur_col + pulls_to_hit_small_width - 1,
                                      self._variant_formats['description_' + str(color_index)],
                                      '',
                                      txt_labels['pulls_to_hit_small'],
                                      '',
                                      little_txt='(1 in ... base spins)',
                                      val_txt_format=self._text_formats['text_big_bold'])
            cur_col = printCommonCell(self._current_row, cur_col,
                                      self._current_row, cur_col + pulls_to_hit_small_merged_width - 1,
                                      self._variant_formats['description_' + str(color_index)],
                                      '',
                                      txt_labels['pulls_to_hit_small_merged'],
                                      '',
                                      little_txt='(1 in ... base spins)',
                                      val_txt_format=self._text_formats['text_regular_bold'])
        if print_percent_big_col:
            cur_col = printCommonCell(self._current_row, cur_col,
                                      self._current_row, cur_col + percent_big_width - 1,
                                      self._variant_formats['description_' + str(color_index)],
                                      '',
                                      txt_labels['percent_big'],
                                      '',
                                      little_txt='(from feature spins)',
                                      val_txt_format=self._text_formats['text_big_bold'])
            cur_col = printCommonCell(self._current_row, cur_col,
                                      self._current_row, cur_col + percent_big_merged_width - 1,
                                      self._variant_formats['description_' + str(color_index)],
                                      '',
                                      txt_labels['percent_big_merged'],
                                      '',
                                      little_txt='(from feature spins)',
                                      val_txt_format=self._text_formats['text_regular_bold'])
        if print_percent_small_col:
            cur_col = printCommonCell(self._current_row, cur_col,
                                      self._current_row, cur_col + percent_small_width - 1,
                                      self._variant_formats['description_' + str(color_index)],
                                      '',
                                      txt_labels['percent_small'],
                                      '',
                                      little_txt='(from whole spins)',
                                      val_txt_format=self._text_formats['text_big_bold'])
            cur_col = printCommonCell(self._current_row, cur_col,
                                      self._current_row, cur_col + percent_small_merged_width - 1,
                                      self._variant_formats['description_' + str(color_index)],
                                      '',
                                      txt_labels['percent_small_merged'],
                                      '',
                                      little_txt='(from whole spins)',
                                      val_txt_format=self._text_formats['text_regular_bold'])

        self._current_row += 1
        cur_col = start_col

        ready_picture_width, ready_picture_height = self._graphMaker.MakeSimpleBarPlot(main_df,
                                                                                       variant_df_col_name,
                                                                                       txt_labels['graph_name'],
                                                                                       include_zero_case=include_zero_in_total,
                                                                                       color_index=color_index,
                                                                                       bar_width=0.8,
                                                                                       figsize=(10, 10),
                                                                                       dpi=300,
                                                                                       transpose=True)
        x_scale, y_scale = self._graphMaker.GetXYScale(ready_picture_width,
                                                       ready_picture_height,
                                                       picture_width,
                                                       len(main_df.index),
                                                       cell_width_pixels=64,
                                                       cell_height_pixels=30)
        self._sheet.insert_image(self._current_row,
                                 cur_col,
                                 self._graphMaker.GetSimpleBarPath(txt_labels['graph_name']),
                                 {
                                     'x_offset': 0,
                                     'y_offset': 1,
                                     'x_scale': x_scale,
                                     'y_scale': y_scale,
                                     'object_position': 2
                                 })
        for i, variant in enumerate(main_df.index):
            cur_col = start_col + picture_width
            self._sheet.set_row_pixels(self._current_row, 30)

            merge_rows = -1
            index_in_merged_df = -1
            if variant in merged_df['left_var'].values:
                index_in_merged_df = merged_df['left_var'][merged_df['left_var'] == variant].index[0]
                left_var = variant
                right_var = merged_df['right_var'].iloc[index_in_merged_df]
                if right_var not in main_df.index:
                    right_var = main_df.index[-1]
                main_df_left_val_index = main_df.index.get_loc(left_var)
                main_df_right_val_index = main_df.index.get_loc(right_var)
                merge_rows = main_df_right_val_index - main_df_left_val_index  # should be + 1, but for merging not

            cur_col = printCommonCell(self._current_row, cur_col,
                                      self._current_row, cur_col + variant_width - 1,
                                      common_bg_format(i, False, True),
                                      '',
                                      str(variant) + '' if not variant_is_mult else 'x',
                                      'bets' if variant_is_mult else '',
                                      val_txt_format=self._text_formats['text_biggest_bold'])
            if print_pulls_to_hit_big_col:
                txt_val = '{:,.2f}'.format(main_df['pulls_to_hit_big'].loc[variant])
                cur_col = printCommonCell(self._current_row, cur_col,
                                          self._current_row, cur_col + pulls_to_hit_big_width - 1,
                                          common_bg_format(i, False, False),
                                          '1 in',
                                          txt_val,
                                          '')
                if variant in merged_df['left_var'].values:
                    txt_val = '{:,.2f}'.format(merged_df['pulls_to_hit_big'].iloc[index_in_merged_df])
                    cur_col = printCommonCell(self._current_row, cur_col,
                                              self._current_row + merge_rows, cur_col + pulls_to_hit_big_merged_width - 1,
                                              common_bg_format(1, True, False),
                                              '1 in',
                                              txt_val,
                                              '')
                else:
                    cur_col += pulls_to_hit_big_merged_width
            if print_pulls_to_hit_small_col:
                txt_val = '{:,.2f}'.format(main_df['pulls_to_hit_small'].loc[variant])
                cur_col = printCommonCell(self._current_row, cur_col,
                                          self._current_row, cur_col + pulls_to_hit_small_width - 1,
                                          common_bg_format(i, False, False),
                                          '1 in',
                                          txt_val,
                                          '')
                if variant in merged_df['left_var'].values:
                    txt_val = '{:,.2f}'.format(merged_df['pulls_to_hit_small'].iloc[index_in_merged_df])
                    cur_col = printCommonCell(self._current_row, cur_col,
                                              self._current_row + merge_rows,
                                              cur_col + pulls_to_hit_small_merged_width - 1,
                                              common_bg_format(1, True, False),
                                              '1 in',
                                              txt_val,
                                              '')
                else:
                    cur_col += pulls_to_hit_small_merged_width
            if print_percent_big_col:
                txt_val = '{:.2f}'.format(main_df['percent_big'].loc[variant])
                cur_col = printCommonCell(self._current_row, cur_col,
                                          self._current_row, cur_col + percent_big_width - 1,
                                          common_bg_format(i, False, False),
                                          '',
                                          txt_val+'%',
                                          '')
                if variant in merged_df['left_var'].values:
                    txt_val = '{:.2f}'.format(merged_df['percent_big'].iloc[index_in_merged_df])
                    cur_col = printCommonCell(self._current_row, cur_col,
                                              self._current_row + merge_rows,
                                              cur_col + percent_big_merged_width - 1,
                                              common_bg_format(1, True, False),
                                              '',
                                              txt_val+'%',
                                              '')
                else:
                    cur_col += percent_big_merged_width
            if print_percent_small_col:
                txt_val = '{:.2f}'.format(main_df['percent_small'].loc[variant])
                cur_col = printCommonCell(self._current_row, cur_col,
                                          self._current_row, cur_col + percent_small_width - 1,
                                          common_bg_format(i, False, False),
                                          '',
                                          txt_val+'%',
                                          '')
                if variant in merged_df['left_var'].values:
                    txt_val = '{:.2f}'.format(merged_df['percent_small'].iloc[index_in_merged_df])
                    cur_col = printCommonCell(self._current_row, cur_col,
                                              self._current_row + merge_rows,
                                              cur_col + percent_small_merged_width - 1,
                                              common_bg_format(1, True, False),
                                              '',
                                              txt_val+'%',
                                              '')
                else:
                    cur_col += percent_small_merged_width
            self._current_row += 1

        self._sheet.set_row_pixels(self._current_row, 45)
        cur_col = start_col +picture_width
        cur_col = printCommonCell(self._current_row, cur_col,
                                  self._current_row, cur_col + variant_width-1,
                                  self._variant_formats['description_'+str(color_index)],
                                  '',
                                  'Total:',
                                  '',
                                  val_txt_format=self._text_formats['text_big_bold'])
        if print_pulls_to_hit_big_col:
            val = '{:.2f}'.format(np.sum(main_df['count']) / np.sum(main_df['count']))
            if not include_zero_in_total and 0 in main_df.index:
                val = '{:.2f}'.format(np.sum(main_df['count']) / (np.sum(main_df['count']) - main_df['count'].loc[0]))
            cur_col = printCommonCell(self._current_row, cur_col,
                                      self._current_row, cur_col + pulls_to_hit_big_width + pulls_to_hit_big_merged_width - 1,
                                      self._variant_formats['description_' + str(color_index)],
                                      '1 in',
                                      val,
                                      '',
                                      prefix_txt_format=self._text_formats['text_regular'],
                                      val_txt_format=self._text_formats['text_big_bold'])
        if print_pulls_to_hit_small_col:
            val = '{:.2f}'.format(self._spin_count / np.sum(main_df['count']))
            if not include_zero_in_total and 0 in main_df.index:
                val = '{:.2f}'.format(self._spin_count / (np.sum(main_df['count']) - main_df['count'].loc[0]))
            cur_col = printCommonCell(self._current_row, cur_col,
                                      self._current_row, cur_col + pulls_to_hit_small_width + pulls_to_hit_small_merged_width - 1,
                                      self._variant_formats['description_' + str(color_index)],
                                      '1 in',
                                      val,
                                      '',
                                      prefix_txt_format=self._text_formats['text_regular'],
                                      val_txt_format=self._text_formats['text_big_bold'])
        if print_percent_big_col:
            val = '{:.2f}'.format(np.sum(main_df['percent_big']))
            if not include_zero_in_total and 0 in main_df.index:
                val = '{:.2f}'.format(np.sum(main_df['percent_big']) - main_df['percent_big'].loc[0])
            cur_col = printCommonCell(self._current_row, cur_col,
                                      self._current_row, cur_col + percent_big_width + percent_big_merged_width - 1,
                                      self._variant_formats['description_' + str(color_index)],
                                      '',
                                      val+'%',
                                      '',
                                      val_txt_format=self._text_formats['text_big_bold'])
        if print_percent_small_col:
            val = '{:.2f}'.format(np.sum(main_df['percent_small']))
            if not include_zero_in_total and 0 in main_df.index:
                val = '{:.2f}'.format(np.sum(main_df['percent_small']) - main_df['percent_small'].loc[0])
            cur_col = printCommonCell(self._current_row, cur_col,
                                      self._current_row, cur_col + percent_small_width + percent_small_merged_width - 1,
                                      self._variant_formats['description_' + str(color_index)],
                                      '',
                                      val+'%',
                                      '',
                                      val_txt_format=self._text_formats['text_big_bold'])
        self._current_row += 4

    def WriteBasicBoardDistribution(self,
                                    positions_counter_df: pd.DataFrame,
                                    row_count: int,
                                    col_count: int,
                                    graph_name: str,
                                    col_graph_color_index: int,
                                    row_graph_color_index: int,
                                    cells_scale_start: str,
                                    cells_scale_end: str,
                                    bar_plot_col_colors: list,
                                    bar_plot_row_colors: list):

        column_graph_height = 5
        row_graph_width = 2
        board_cell_width = 2
        board_cell_height = 2
        start_col = 1
        one_cell_height = 30

        row_graph_name = graph_name + '_by_rows'
        col_graph_name = graph_name + '_by_cols'

        # ИНДЕКСАЦИЯ ПОЗИЦИЙ ПОСТРОЧНАЯ !!!

        def on_one_row(cell_index: int):
            for row in range(row_count):
                if row * col_count <= cell_index < (row+1) * col_count:
                    return row

        def on_one_col(cell_index: int):
            return cell_index % col_count

        row_pic_width, row_pic_height = self._graphMaker.MakeSimpleBarPlot(positions_counter_df.groupby(on_one_row).sum(),
                                                                           count_col_name='count',
                                                                           graph_name=row_graph_name,
                                                                           include_zero_case=True,
                                                                           color_index=row_graph_color_index,
                                                                           bar_width=0.8,
                                                                           figsize=(5, 5),
                                                                           dpi=300,
                                                                           transpose=True,
                                                                           str_colors=bar_plot_row_colors)
        col_pic_width, col_pic_height = self._graphMaker.MakeSimpleBarPlot(positions_counter_df.groupby(on_one_col).sum(),
                                                                           count_col_name='count',
                                                                           graph_name=col_graph_name,
                                                                           include_zero_case=True,
                                                                           color_index=col_graph_color_index,
                                                                           bar_width=0.8,
                                                                           figsize=(5, 5),
                                                                           dpi=300,
                                                                           transpose=False,
                                                                           str_colors=bar_plot_col_colors)
        row_pic_x_scale, row_pic_y_scale = self._graphMaker.GetXYScale(row_pic_width, row_pic_height,
                                                                       row_graph_width,
                                                                       row_count * board_cell_height,
                                                                       cell_height_pixels=one_cell_height)
        col_pic_x_scale, col_pic_y_scale = self._graphMaker.GetXYScale(col_pic_width, col_pic_height,
                                                                       col_count * board_cell_width,
                                                                       column_graph_height,
                                                                       cell_height_pixels=one_cell_height)
        self._sheet.insert_image(self._current_row, start_col + row_graph_width,
                                 self._graphMaker.GetSimpleBarPath(col_graph_name),
                                 {
                                     'x_offset': 0,
                                     'y_offset': -1,
                                     'x_scale': col_pic_x_scale,
                                     'y_scale': col_pic_y_scale,
                                     'object_position': 2
                                 })
        for row in range(self._current_row, self._current_row + column_graph_height + row_count * board_cell_height):
            self._sheet.set_row_pixels(row, one_cell_height)
        self._current_row += column_graph_height
        self._sheet.insert_image(self._current_row, start_col,
                                 self._graphMaker.GetSimpleBarPath(row_graph_name),
                                 {
                                     'x_offset': -1,
                                     'y_offset': 0,
                                     'x_scale': row_pic_x_scale,
                                     'y_scale': row_pic_y_scale,
                                     'object_position': 2
                                 })

        color_scales = Colors.get_scales_for_numbers(positions_counter_df['count'],
                                                     cells_scale_start,
                                                     cells_scale_end)
        formats = self._par_formats.GetBoardPositionsFormats(color_scales)

        for board_cell_index in positions_counter_df.index:
            row = on_one_row(board_cell_index)
            col = on_one_col(board_cell_index)

            self._sheet.merge_range(self._current_row + row * board_cell_height,
                                    start_col + row_graph_width + col * board_cell_width,
                                    self._current_row + row * board_cell_height + board_cell_height - 1,
                                    start_col + row_graph_width + col * board_cell_width + board_cell_width - 1,
                                    '', formats['index_'+str(board_cell_index)])
            percent = '{:.2f}%'.format(positions_counter_df['percent_big'].loc[board_cell_index])
            pulls = '(1 in {:.2f})'.format(positions_counter_df['pulls_to_hit_big'].loc[board_cell_index])
            self._sheet.write_rich_string(self._current_row + row * board_cell_height,
                                          start_col + row_graph_width + col * board_cell_width,
                                          self._text_formats['text_big_bold'], percent+'\n',
                                          self._text_formats['text_small'], pulls,
                                          formats['index_'+str(board_cell_index)])
        self._current_row += row_count * board_cell_height + 3

    def PrintInfoRow(self,
                     start_col: int,
                     info_width: int,
                     val_width: int,
                     inf_txt_segments: list,
                     val_txt_segments: list,
                     bg_format_info,
                     bg_format_val,
                     row_height_in_pixels: int = 30,
                     merged_rows: int = 1
                     ):
        for row in range(self._current_row, self._current_row+merged_rows):
            self._sheet.set_row_pixels(row, row_height_in_pixels)

        val_txt_segments.append(self._text_formats['text_regular'])
        val_txt_segments.append(' ')
        inf_txt_segments.append(self._text_formats['text_regular'])
        inf_txt_segments.append(' ')

        self._sheet.merge_range(self._current_row, start_col,
                                self._current_row+merged_rows-1, start_col+info_width-1,
                                '', bg_format_info)
        self._sheet.write_rich_string(self._current_row, start_col,
                                      *inf_txt_segments,
                                      bg_format_info)
        self._sheet.merge_range(self._current_row, start_col+info_width,
                                self._current_row+merged_rows-1, start_col+info_width+val_width-1,
                                '', bg_format_val)
        self._sheet.write_rich_string(self._current_row, start_col+info_width,
                                      *val_txt_segments,
                                      bg_format_val)
        self._current_row += 1

    def WriteFeatureRespinSettings(self, color_index: int = 0):
        self.WriteInnerHeader("Feature Re-Spin Settings", include_in_content=True, print_counter=True)
        start_col = 1
        info_table_info_width = 5
        info_table_val_width = 2
        info_table_start_col = self.GetCentreStartCol(12, info_table_val_width + info_table_info_width,
                                                      default_start_col=start_col)

        # WRITE INFO TABLE
        info_format = lambda row_counter: self._variant_formats[
            'info_table_even_' + str(color_index)] if info_rows_counter % 2 == 0 else \
            self._variant_formats['info_table_odd_' + str(color_index)]
        info_rows_counter = 0

        def printInfoRow(format, txt_label, little_txt, val_prefix, txt_val, val_postfix):
            self._sheet.set_row_pixels(self._current_row, 45)
            self._sheet.merge_range(self._current_row, info_table_start_col,
                                    self._current_row, info_table_start_col + info_table_info_width - 1,
                                    "", format)
            self._sheet.write_rich_string(self._current_row, info_table_start_col,
                                          self._text_formats['text_big_bold'],
                                          txt_label + '\n',
                                          self._text_formats['text_smallest'], little_txt,
                                          format)
            self._sheet.merge_range(self._current_row, info_table_start_col + info_table_info_width,
                                    self._current_row,
                                    info_table_start_col + info_table_info_width + info_table_val_width - 1,
                                    "", format)
            self._sheet.write_rich_string(self._current_row, info_table_start_col + info_table_info_width,
                                          self._text_formats['text_smallest'], val_prefix + ' ',
                                          self._text_formats['text_big'], txt_val,
                                          self._text_formats['text_smallest'], ' ' + val_postfix,
                                          format)
            self._current_row += 1

        printInfoRow(info_format(info_rows_counter),
                     'Strategy',
                     ' ',
                     ' ',
                     self._feature_respin_strategy,
                     ' ')
        info_rows_counter += 1

        for i, (sim_type, mults) in enumerate(self._feature_respin_mults.items()):
            for j, (fs_count, mult) in enumerate(mults.items()):
                printInfoRow(info_format(info_rows_counter),
                             sim_type + ' ' + str(fs_count) + ' free spins',
                             '(average win)',
                             ' ',
                             '{:,.2f}'.format(mult) + 'x',
                             'bets')
                info_rows_counter += 1

        self._current_row += 1

    def GetCentreStartCol(self, total_width: int, table_width: int, default_start_col: int = 1):
        if table_width >= total_width:
            return default_start_col
        diff = total_width - table_width
        return int(default_start_col + np.floor(diff / 2))

    def GetHeaderWidth(self):
        return self._header_width

    def GetInnerHeaderWidth(self):
        return self._inner_header_width

    def GetSmallHeaderWidth(self):
        return self._small_header_width

    def GetHeaderStartColumn(self):
        return self._header_start_col

    def GetInnerHeaderStartColumn(self):
        return self._inner_header_start_col

    def GetSmallHeaderStartColumn(self):
        return self._small_header_start_col

    def GetSmallCellPixelHeight(self):
        return self._small_cell_pixel_height

    def GetNormalCellPixelHeight(self):
        return self._normal_cell_pixel_height

    def GetBigCellPixelHeight(self):
        return self._big_cell_pixel_height


if __name__ == "__main__":
    file = open(
        r"C:\Users\VitalijSaiganov\PycharmProjects\SlotHandler\Game_Source_Data\NBN\86\NBN_usual_86_no_gamble_stats.json")
    data = j.load(file)

    slot = BasicSlot()
    slot.ReadSlot_Json(data)

    stats = BasicStatistics()
    stats.ReadStatistics(data)

    calculation = BasicStatsCalculator(slot, stats)
    calculation.CalcStats()

    writer = BasicPARSheet(slot, stats, calculation)
    writer.WritePARSheet()





