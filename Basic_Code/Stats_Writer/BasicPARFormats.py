from Basic_Code.Basic_Structures.BasicSlot import BasicSlot
from Basic_Code.Basic_Structures.BasicStatistics import BasicStatistics
from Basic_Code.Basic_Calculator.BasicStatsCalculator import BasicStatsCalculator
from Basic_Code.Utils.BasicPathHadler import BasicPathHandler
from Basic_Code.Stats_Writer.BasicGraphMaker import BasicGraphMaker
from Basic_Code.Utils.BasicTimer import timer
import xlsxwriter as xl
import json as j
from collections import defaultdict
from Basic_Code.Stats_Writer.BasicGraphMaker import Colors


class BasicPARFormats:
    def __init__(self, book: xl.Workbook):
        self._book = book

        self._font_name = 'Fira sans' #'Helvetica'

        self._header_formats = dict()
        self._summary_formats = dict()
        self._paytable_formats = dict()
        self._text_formats = dict()
        self._winline_formats = dict()
        self._rtp_distribution_formats = dict()
        self._line_wins_formats = dict()
        self._confidence_formats = dict()
        self._top_award_formats = dict()
        self._reelsets_table_formats = dict()
        self._reelsets_spin_win_formats = dict()
        self._reelset_structure_formats = dict()
        self._content_formats = dict()
        self._variant_formats = dict()
        self._info_row_formats = dict()


    def InitFormats(self):
        self._InitHeaderFormats()
        self._InitSummaryFormats()
        self._InitPaytableFormats()
        self._InitFontFormats()
        self._InitWinlineFormats()
        self._InitRTPDistributionFormats()
        self._InitLineWinsFormats()
        self._InitConfidenceFormats()
        self._InitTopAwardFormats()
        self._InitReelsetsTableFormats()
        self._InitReelsetsSpinWinFormats()
        self._InitReelsetsStructureFormats()
        self._InitContentFormats()
        self._InitVariantFormats()
        self._InitInfoRowFormats()

    def GetHeaderFormats(self):
        return self._header_formats

    def GetSummaryFormats(self):
        return self._summary_formats

    def GetPaytableFormats(self):
        return self._paytable_formats

    def GetFontFormats(self):
        return self._text_formats

    def GetWinlineFormats(self):
        return self._winline_formats

    def GetRTPDistributionFormats(self):
        return self._rtp_distribution_formats

    def GetLineWinsFormats(self):
        return self._line_wins_formats

    def GetCondideneceFormats(self):
        return self._confidence_formats

    def GetTopAwardFormats(self):
        return self._top_award_formats

    def GetReelsetsTableFormats(self):
        return self._reelsets_table_formats

    def GetReelsetsSpinWinFormats(self):
        return self._reelsets_spin_win_formats

    def GetReelsetsStructureFormats(self):
        return self._reelset_structure_formats

    def GetContentFormats(self):
        return self._content_formats

    def GetVariantFormats(self):
        return self._variant_formats

    def GetInfoRowFormats(self):
        return self._info_row_formats

    def GetBoardPositionsFormats(self, colors: list):
        return self._InitBoardPositionFormats(colors)

    def _InitHeaderFormats(self):
        self._InitHeaderFormat('header')
        self._InitInnerHeaderFormat('inner_header')
        self._InitSmallHeaderFormat('small_header')

    def _InitSummaryFormats(self):
        self._InitBorderCellDescription('border_cell_description')
        self._InitBorderUnderDescription('border_under_description')
        self._InitBorderUnderDescription_Features('border_under_description_feature')
        self._InitBorderBigValues('border_rtp_value')

    def _InitPaytableFormats(self):
        border_color = 3
        bg_description = 7
        bg_even = -3
        bg_odd = -1

        self._InitPaytableBorder_Common('border_even', bg_even, border_color)
        self._InitPaytableBorder_Common('border_odd', bg_odd, border_color)
        self._InitPytableBorder_Description('border_description', bg_description, border_color)

    def _InitFontFormats(self):
        self._InitText_Smallest('text_smallest')
        self._InitText_SmallestBold('text_smallest_bold')
        self._InitText_Small('text_small')
        self._InitText_SmallBold('text_small_bold')
        self._InitText_Regular('text_regular')
        self._InitText_RegularBold('text_regular_bold')
        self._InitText_Big('text_big')
        self._InitText_BigBold('text_big_bold')
        self._InitText_Biggest('text_biggest')
        self._InitText_BiggestBold('text_biggest_bold')
        self._InitText_RedInfo('text_red_info')

    def _InitWinlineFormats(self):
        winline_color = Colors.blue_scale[8]
        border_color = Colors.blue_scale[3]

        self._InitWinlineFormat('side_top_free', [2], '', border_color)
        self._InitWinlineFormat('side_bottom_free', [4], '', border_color)
        self._InitWinlineFormat('side_left_free', [1], '', border_color)
        self._InitWinlineFormat('side_right_free', [3], '', border_color)
        self._InitWinlineFormat('corner_left_top_free', [1, 2], '', border_color)
        self._InitWinlineFormat('corner_right_top_free', [2, 3], '', border_color)
        self._InitWinlineFormat('corner_left_bottom_free', [1, 4], '', border_color)
        self._InitWinlineFormat('corner_right_bottom_free', [3, 4], '', border_color)

        self._InitWinlineFormat('side_top_color', [2], winline_color, border_color)
        self._InitWinlineFormat('side_bottom_color', [4], winline_color, border_color)
        self._InitWinlineFormat('side_left_color', [1], winline_color, border_color)
        self._InitWinlineFormat('side_right_color', [3], winline_color, border_color)
        self._InitWinlineFormat('corner_left_top_color', [1, 2], winline_color, border_color)
        self._InitWinlineFormat('corner_right_top_color', [2, 3], winline_color, border_color)
        self._InitWinlineFormat('corner_left_bottom_color', [1, 4], winline_color, border_color)
        self._InitWinlineFormat('corner_right_bottom_color', [3, 4], winline_color, border_color)

        self._InitWinlineFormat('centre_color', [], winline_color, border_color)
        self._InitWinlineFormat('centre_free', [], "", border_color)

    def _InitRTPDistributionFormats(self):
        description_bg_color = 7
        border_color = 5
        even_color = -3
        odd_color = -1

        for color_index in range(len(Colors.scales)):
            self._InitRTPDistribution_BorderDescriptionFormat('border_description_' + str(color_index),
                                                              Colors.scales[color_index][description_bg_color],
                                                              Colors.scales[color_index][border_color])
            self._InitRTPDistribution_LightFormat('border_even_'+str(color_index),
                                                  Colors.scales[color_index][even_color],
                                                  Colors.scales[color_index][border_color])
            self._InitRTPDistribution_LightFormat('border_odd_'+str(color_index),
                                                  Colors.scales[color_index][odd_color],
                                                  Colors.scales[color_index][border_color])


    def _InitLineWinsFormats(self):
        bg_description = 6
        border = 2
        bg_symbol = 11
        bg_common_even = 13
        bg_common_odd = 15

        for i in range(len(Colors.scales)):
            self._InitLineWins_BorderDescriptionFormat('border_description_'+str(i),
                                                       Colors.scales[i][bg_description],
                                                       Colors.scales[i][border])
            self._InitLineWins_BorderSymbolFormat('border_symbol_left_'+str(i),
                                                  Colors.scales[i][bg_symbol],
                                                  Colors.scales[i][border],
                                                  [2, 2, 1, 2])
            self._InitLineWins_BorderSymbolFormat('border_symbol_right_'+str(i),
                                                  Colors.scales[i][bg_common_odd],
                                                  Colors.scales[i][border],
                                                  [1, 2, 2, 2])
            self._InitLineWins_BorderLengthFormat('border_length_up_'+str(i),
                                                  Colors.scales[i][bg_common_even],
                                                  Colors.scales[i][border],
                                                  [1, 2, 1, 1])
            self._InitLineWins_BorderLengthFormat('border_length_middle_' + str(i),
                                                  Colors.scales[i][bg_common_even],
                                                  Colors.scales[i][border],
                                                  [1, 1, 1, 1])
            self._InitLineWins_BorderLengthFormat('border_length_down_' + str(i),
                                                  Colors.scales[i][bg_common_even],
                                                  Colors.scales[i][border],
                                                  [1, 1, 1, 2])
            self._InitLineWins_BorderLengthFormat('border_length_one_' + str(i),
                                                  Colors.scales[i][bg_common_even],
                                                  Colors.scales[i][border],
                                                  [1, 2, 1, 2])
            self._InitLineWins_BorderCommonFormat('border_even_up_'+str(i),
                                                  Colors.scales[i][bg_common_even],
                                                  Colors.scales[i][border],
                                                  [1, 2, 1, 1])
            self._InitLineWins_BorderCommonFormat('border_even_middle_' + str(i),
                                                  Colors.scales[i][bg_common_even],
                                                  Colors.scales[i][border],
                                                  [1, 1, 1, 1])
            self._InitLineWins_BorderCommonFormat('border_even_down_' + str(i),
                                                  Colors.scales[i][bg_common_even],
                                                  Colors.scales[i][border],
                                                  [1, 1, 1, 2])
            self._InitLineWins_BorderCommonFormat('border_even_one_' + str(i),
                                                  Colors.scales[i][bg_common_even],
                                                  Colors.scales[i][border],
                                                  [1, 2, 1, 2])
            self._InitLineWins_BorderCommonFormat('border_odd_up_' + str(i),
                                                  Colors.scales[i][bg_common_odd],
                                                  Colors.scales[i][border],
                                                  [1, 2, 1, 1])
            self._InitLineWins_BorderCommonFormat('border_odd_middle_' + str(i),
                                                  Colors.scales[i][bg_common_odd],
                                                  Colors.scales[i][border],
                                                  [1, 1, 1, 1])
            self._InitLineWins_BorderCommonFormat('border_odd_down_' + str(i),
                                                  Colors.scales[i][bg_common_odd],
                                                  Colors.scales[i][border],
                                                  [1, 1, 1, 2])
            self._InitLineWins_BorderCommonFormat('border_odd_one_' + str(i),
                                                  Colors.scales[i][bg_common_odd],
                                                  Colors.scales[i][border],
                                                  [1, 2, 1, 2])

    def _InitConfidenceFormats(self):
        color_total = Colors.scales[0]
        section_colors = Colors.scales[1:]

        bg_index = 6
        border_index = 2
        even_index = -3
        odd_index = -1

        self._InitConfidence_DescriptionFormat('border_description_total',
                                               color_total[bg_index],
                                               color_total[border_index])
        self._InitConfidence_CommonFormat('border_even_total',
                                          color_total[even_index],
                                          color_total[border_index])
        self._InitConfidence_CommonFormat('border_odd_total',
                                          color_total[odd_index],
                                          color_total[border_index])

        for section_index in range(len(Colors.scales)-1):
            self._InitConfidence_DescriptionFormat('border_description_' + str(section_index),
                                                   section_colors[section_index][bg_index],
                                                   section_colors[section_index][border_index])
            self._InitConfidence_CommonFormat('border_even_'+str(section_index),
                                              section_colors[section_index][even_index],
                                              section_colors[section_index][border_index])
            self._InitConfidence_CommonFormat('border_odd_' + str(section_index),
                                              section_colors[section_index][odd_index],
                                              section_colors[section_index][border_index])

    def _InitTopAwardFormats(self):
        self._InitTopAward_DescriptionFormat('border_description')
        self._InitTopAward_EvenFormat('border_even')
        self._InitTopAward_OddFormat('border_odd')

    def _InitReelsetsTableFormats(self):
        for color_index in range(len(Colors.scales)):
            self._InitReelsetsTable_TopHeaderFormat('border_header_'+str(color_index), color_index)
            self._InitReelsetsTable_DescriptionFormat('border_description_' + str(color_index), color_index)
            self._InitReelsetsTable_EvenFormat('border_even_'+str(color_index), color_index, 'centre')
            self._InitReelsetsTable_OddFormat('border_odd_'+str(color_index), color_index, 'centre')
            self._InitReelsetsTable_EvenFormat('border_even_name_' + str(color_index), color_index, 'left')
            self._InitReelsetsTable_OddFormat('border_odd_name_' + str(color_index), color_index, 'left')

    def _InitReelsetsSpinWinFormats(self):
        for color_index in range(len(Colors.scales)):
            self._InitReelsetsSpinWin_InfoCell('info_cell_'+str(color_index), color_index)
            self._InitReelsetsSpinWin_ValueCell('value_cell_'+str(color_index), color_index)
            self._InitReelsetsSpinWin_BottomBorderCell('bottom_border_'+str(color_index), color_index)

    def _InitReelsetsStructureFormats(self):
        top_info_color = -6
        top_value_color = -3
        description_color = 5
        under_description_color = 9
        border_color = 0
        even_color = -2
        odd_color = -1

        for color_index in range(len(Colors.scales)):
            self._InitReelsetsStructure_TopInfoFormat('top_info_'+str(color_index), color_index, top_info_color, border_color)
            self._InitReelsetsStructure_TopValueFormat('top_value_'+str(color_index), color_index, top_value_color, border_color)
            self._InitReelsetsStructure_DescriptionFormat('top_description_'+str(color_index), color_index, description_color, border_color)
            self._InitReelsetsStructure_DescriptionFormat('bottom_description_' + str(color_index), color_index,
                                                          under_description_color, border_color)
            self._InitReelsetsStructure_CommonFormat('even_symbols_'+str(color_index), color_index, even_color, border_color, [2, 1, 1, 1])
            self._InitReelsetsStructure_CommonFormat('odd_symbols_' + str(color_index), color_index, odd_color, border_color, [2, 1, 1, 1])
            self._InitReelsetsStructure_CommonFormat('even_weights_' + str(color_index), color_index, even_color,
                                                     border_color, [1, 1, 2, 1])
            self._InitReelsetsStructure_CommonFormat('odd_weights_' + str(color_index), color_index, odd_color,
                                                     border_color, [1, 1, 2, 1])
            self._InitReelsetsStructure_CommonFormat('even_position_' + str(color_index), color_index, even_color,
                                                     border_color, [2, 1, 2, 1])
            self._InitReelsetsStructure_CommonFormat('odd_position_' + str(color_index), color_index, odd_color,
                                                     border_color, [2, 1, 2, 1])
            self._InitReelsetsStructure_CommonFormat('even_symbols_end_' + str(color_index), color_index, even_color,
                                                     border_color, [2, 1, 1, 2])
            self._InitReelsetsStructure_CommonFormat('odd_symbols_end_' + str(color_index), color_index, odd_color,
                                                     border_color, [2, 1, 1, 2])
            self._InitReelsetsStructure_CommonFormat('even_weights_end_' + str(color_index), color_index, even_color,
                                                     border_color, [1, 1, 2, 2])
            self._InitReelsetsStructure_CommonFormat('odd_weights_end_' + str(color_index), color_index, odd_color,
                                                     border_color, [1, 1, 2, 2])
            self._InitReelsetsStructure_CommonFormat('even_position_end_' + str(color_index), color_index, even_color,
                                                     border_color, [2, 1, 2, 2])
            self._InitReelsetsStructure_CommonFormat('odd_position_end_' + str(color_index), color_index, odd_color,
                                                     border_color, [2, 1, 2, 2])

    def _InitVariantFormats(self):
        description_bg_color = 7
        even_bg_color = -2
        odd_bg_color = -1
        border_color = 2

        for color_index in range(len(Colors.scales)):
            self._InitVariantFormat_InfoTable('info_table_even_'+str(color_index),
                                              color_index,
                                              even_bg_color,
                                              border_color)
            self._InitVariantFormat_InfoTable('info_table_odd_' + str(color_index),
                                              color_index,
                                              odd_bg_color,
                                              border_color)
            self._InitVariantFormat_Description('description_'+str(color_index),
                                                color_index,
                                                description_bg_color,
                                                border_color)
            self._InitVariantFormat_Common('even_left_'+str(color_index),
                                           color_index,
                                           even_bg_color,
                                           border_color,
                                           [2, 1, 1, 1],
                                           'left')
            self._InitVariantFormat_Common('even_right_' + str(color_index),
                                           color_index,
                                           even_bg_color,
                                           border_color,
                                           [1, 1, 2, 1],
                                           'left')
            self._InitVariantFormat_Common('even_centre_' + str(color_index),
                                           color_index,
                                           even_bg_color,
                                           border_color,
                                           [1, 1, 1, 1],
                                           'left')
            self._InitVariantFormat_Common('odd_left_' + str(color_index),
                                           color_index,
                                           odd_bg_color,
                                           border_color,
                                           [2, 1, 1, 1],
                                           'left')
            self._InitVariantFormat_Common('odd_right_' + str(color_index),
                                           color_index,
                                           odd_bg_color,
                                           border_color,
                                           [1, 1, 2, 1],
                                           'left')
            self._InitVariantFormat_Common('odd_centre_' + str(color_index),
                                           color_index,
                                           odd_bg_color,
                                           border_color,
                                           [1, 1, 1, 1],
                                           'left')

    def _InitVariantFormat_InfoTable(self, format_name: str, color_index: int, bg_color: int, border_color: int):
        info_table_format = self._book.add_format()
        info_table_format.set_border(1)
        info_table_format.set_bg_color(Colors.scales[color_index][bg_color])
        info_table_format.set_border_color(Colors.scales[color_index][border_color])
        info_table_format.set_align('left')
        info_table_format.set_align('vcentre')
        info_table_format.set_text_wrap(True)
        self._variant_formats[format_name] = info_table_format

    def _InitVariantFormat_Description(self, format_name: str, color_index: int, bg_color: int, border_color: int):
        description_format = self._book.add_format()
        description_format.set_border(2)
        description_format.set_bg_color(Colors.scales[color_index][bg_color])
        description_format.set_border_color(Colors.scales[color_index][border_color])
        description_format.set_align('centre')
        description_format.set_align('vcentre')
        description_format.set_text_wrap(True)
        self._variant_formats[format_name] = description_format

    def _InitVariantFormat_Common(self, format_name: str, color_index: int, bg_color: int, border_color: int,
                                    borders: list, align: str):
        even_left_format = self._book.add_format()
        even_left_format.set_left(borders[0])
        even_left_format.set_top(borders[1])
        even_left_format.set_right(borders[2])
        even_left_format.set_bottom(borders[3])
        even_left_format.set_bg_color(Colors.scales[color_index][bg_color])
        even_left_format.set_border_color(Colors.scales[color_index][border_color])
        even_left_format.set_align(align)
        even_left_format.set_align('vcentre')
        even_left_format.set_text_wrap(True)
        self._variant_formats[format_name] = even_left_format

    def _InitInfoRowFormats(self):
        border_color = 1
        even_color = -3
        odd_color = -1

        for color_index in range(len(Colors.scales)):
            self._InitInfoRowFormat_Common('info_even_info_'+str(color_index),
                                           color_index,
                                           even_color,
                                           border_color,
                                           [2, 2, 1, 2],
                                           'left')
            self._InitInfoRowFormat_Common('info_odd_info_' + str(color_index),
                                           color_index,
                                           odd_color,
                                           border_color,
                                           [2, 2, 1, 2],
                                           'left')
            self._InitInfoRowFormat_Common('info_even_val_' + str(color_index),
                                           color_index,
                                           even_color,
                                           border_color,
                                           [1, 2, 2, 2],
                                           'left')
            self._InitInfoRowFormat_Common('info_odd_val_' + str(color_index),
                                           color_index,
                                           odd_color,
                                           border_color,
                                           [1, 2, 2, 2],
                                           'left')

    def _InitInfoRowFormat_Common(self, format_name: str, color_index: int, bg_color: int, border_color: int, borders: list,
                                  align: str):
        common_format = self._book.add_format()
        common_format.set_left(borders[0])
        common_format.set_top(borders[1])
        common_format.set_right(borders[2])
        common_format.set_bottom(borders[3])
        common_format.set_bg_color(Colors.scales[color_index][bg_color])
        common_format.set_border_color(Colors.scales[color_index][border_color])
        common_format.set_align(align)
        common_format.set_align('vcentre')
        common_format.set_text_wrap(True)
        self._info_row_formats[format_name] = common_format

    def _InitBoardPositionFormats(self, colors):
        res_formats = dict()
        for i, color in enumerate(colors):
            cur_format = self._book.add_format()
            cur_format.set_bg_color(color)
            cur_format.set_border_color('black')
            cur_format.set_border(2)
            cur_format.set_font(self._font_name)
            cur_format.set_text_wrap(True)
            cur_format.set_align('centre')
            cur_format.set_align('vcentre')
            res_formats['index_'+str(i)] = cur_format
        return res_formats

    def GetReelsetStructureFormatName(self, color_index: int, even: bool, end: bool, symbols: bool):
        if even:
            if end:
                if symbols:
                    return 'even_symbols_end_'+str(color_index)
                else:
                    return 'even_weights_end_'+str(color_index)
            else:
                if symbols:
                    return 'even_symbols_'+str(color_index)
                else:
                    return 'even_weights_'+str(color_index)
        else:
            if end:
                if symbols:
                    return 'odd_symbols_end_'+str(color_index)
                else:
                    return 'odd_weights_end_'+str(color_index)
            else:
                if symbols:
                    return 'odd_symbols_'+str(color_index)
                else:
                    return 'odd_weights_'+str(color_index)

    def _InitContentFormats(self):
        bg_big = -4
        bg_small = -1
        border = 4
        self._InitContentFormat_Big('big', bg_big, border)
        self._InitContentFormat_Small('small', bg_small, border)
        self._InitGoContentFormat_ButtonHeaders('content_button_big', 0, -1, 1)
        self._InitGoContentFormat_ButtonHeaders('content_button_small', 0, -1, 3)

        bg_button_color = -3
        border_button_color = 5

        for color_index in range(len(Colors.scales)):
            self._InitGoContentFormat_ButtonColor('content_button_'+str(color_index),
                                                  color_index,
                                                  bg_button_color,
                                                  border_button_color)

    def _InitGoContentFormat_ButtonHeaders(self, format_name: str, color_index: int, bg_color: int, border_color: int):
        go_home_button = self._book.add_format()
        go_home_button.set_font_name(self._font_name)
        go_home_button.set_font_size(12)
        go_home_button.set_underline(True)
        go_home_button.set_font_color('blue')
        go_home_button.set_bg_color(Colors.scales[color_index][bg_color])
        go_home_button.set_border(2)
        go_home_button.set_border_color(Colors.blue_hexes[border_color])
        go_home_button.set_align('centre')
        go_home_button.set_align('vcentre')
        go_home_button.set_text_wrap(True)
        self._content_formats[format_name] = go_home_button

    def _InitGoContentFormat_ButtonColor(self, format_name: str, color_index: int, bg_color: int, border_color: int):
        go_home_button = self._book.add_format()
        go_home_button.set_font_name(self._font_name)
        go_home_button.set_font_size(12)
        go_home_button.set_underline(True)
        go_home_button.set_font_color('blue')
        go_home_button.set_bg_color(Colors.scales[color_index][bg_color])
        go_home_button.set_border(1)
        go_home_button.set_border_color(Colors.scales[color_index][border_color])
        go_home_button.set_align('centre')
        go_home_button.set_align('vcentre')
        go_home_button.set_text_wrap(True)
        self._content_formats[format_name] = go_home_button

    def _InitContentFormat_Big(self, format_name: str, bg_color: int, border_color: int):
        big_format = self._book.add_format()
        big_format.set_font_name(self._font_name)
        big_format.set_font_size(16)
        big_format.set_font_color('blue')
        big_format.set_underline(True)
        big_format.set_align('left')
        big_format.set_align('vcentre')
        big_format.set_border(2)
        big_format.set_bg_color(Colors.blue_scale[bg_color])
        big_format.set_border_color(Colors.blue_scale[border_color])
        self._content_formats[format_name] = big_format

    def _InitContentFormat_Small(self, format_name: str, bg_color: int, border_color: int):
        small_format = self._book.add_format()
        small_format.set_font_name(self._font_name)
        small_format.set_font_size(12)
        small_format.set_font_color('blue')
        small_format.set_underline(True)
        small_format.set_align('left')
        small_format.set_align('vcentre')
        small_format.set_border(1)
        small_format.set_bg_color(Colors.blue_scale[bg_color])
        small_format.set_border_color(Colors.blue_scale[border_color])
        self._content_formats[format_name] = small_format


    def _InitReelsetsStructure_TopInfoFormat(self, format_name: str, color_index: int, bg_color: int, border_color: int):
        info_cell = self._book.add_format()
        info_cell.set_bg_color(Colors.scales[color_index][bg_color])
        info_cell.set_align('left')
        info_cell.set_align('vcentre')
        info_cell.set_font_name(self._font_name)
        info_cell.set_font_size(13)
        info_cell.set_border(1)
        info_cell.set_left(2)
        info_cell.set_border_color(Colors.scales[color_index][border_color])
        info_cell.set_text_wrap(True)
        self._reelset_structure_formats[format_name] = info_cell

    def _InitReelsetsStructure_TopValueFormat(self, format_name: str, color_index: int, bg_color: int, border_color: int):
        value_cell = self._book.add_format()
        value_cell.set_bg_color(Colors.scales[color_index][bg_color])
        value_cell.set_align('left')
        value_cell.set_align('vcentre')
        value_cell.set_font_name(self._font_name)
        value_cell.set_font_size(13)
        value_cell.set_bold(True)
        value_cell.set_border(1)
        value_cell.set_left(2)
        value_cell.set_border_color(Colors.scales[color_index][border_color])
        value_cell.set_text_wrap(True)
        self._reelset_structure_formats[format_name] = value_cell

    def _InitReelsetsStructure_DescriptionFormat(self, format_name: str, color_index: int, bg_color: int, border_color: int):
        up_description = self._book.add_format()
        up_description.set_bg_color(Colors.scales[color_index][bg_color])
        up_description.set_align('centre')
        up_description.set_align('vcentre')
        up_description.set_font_name(self._font_name)
        up_description.set_border(2)
        up_description.set_border_color(Colors.scales[color_index][border_color])
        self._reelset_structure_formats[format_name] = up_description

    def _InitReelsetsStructure_CommonFormat(self, format_name: str, color_index: int, bg_color: int, border_color: int, borders: list):
        common_format = self._book.add_format()
        common_format.set_bg_color(Colors.scales[color_index][bg_color])
        common_format.set_font_name(self._font_name)
        common_format.set_left(borders[0])
        common_format.set_top(borders[1])
        common_format.set_right(borders[2])
        common_format.set_bottom(borders[3])
        if borders[0] == borders[2]:
            common_format.set_align('centre')
        else:
            common_format.set_align('left')
        common_format.set_align('vcentre')
        common_format.set_border_color(Colors.scales[color_index][border_color])
        self._reelset_structure_formats[format_name] = common_format

    def _InitReelsetsSpinWin_InfoCell(self, format_name: str, color_index: int):
        info_cell = self._book.add_format()
        info_cell.set_bg_color(Colors.scales[color_index][-6])
        info_cell.set_align('left')
        info_cell.set_align('vcentre')
        info_cell.set_font_name(self._font_name)
        info_cell.set_font_size(13)
        info_cell.set_border(1)
        info_cell.set_left(2)
        info_cell.set_border_color(Colors.scales[color_index][5])
        info_cell.set_text_wrap(True)
        self._reelsets_spin_win_formats[format_name] = info_cell


    def _InitReelsetsSpinWin_ValueCell(self, format_name: str, color_index: int):
        value_cell = self._book.add_format()
        value_cell.set_bg_color(Colors.scales[color_index][-3])
        value_cell.set_align('left')
        value_cell.set_align('vcentre')
        value_cell.set_font_name(self._font_name)
        value_cell.set_font_size(13)
        value_cell.set_bold(True)
        value_cell.set_border(1)
        value_cell.set_left(2)
        value_cell.set_border_color(Colors.scales[color_index][5])
        value_cell.set_text_wrap(True)
        self._reelsets_spin_win_formats[format_name] = value_cell

    def _InitReelsetsSpinWin_BottomBorderCell(self, font_name: str, color_index: int):
        border_cell = self._book.add_format()
        border_cell.set_bottom(2)
        border_cell.set_border_color(Colors.scales[color_index][3])
        self._reelsets_spin_win_formats[font_name] = border_cell

    def _InitReelsetsTable_TopHeaderFormat(self, format_name: str, color_index: int):
        border_header = self._book.add_format()
        border_header.set_bg_color(Colors.scales[color_index][3])
        border_header.set_border_color(Colors.scales[color_index][3])
        border_header.set_border(2)
        border_header.set_align('centre')
        border_header.set_align('vcentre')
        border_header.set_font_name(self._font_name)
        border_header.set_font_size(20)
        border_header.set_font_color('white')
        border_header.set_text_wrap(True)
        border_header.set_bold(True)
        self._reelsets_table_formats[format_name] = border_header

    def _InitReelsetsTable_DescriptionFormat(self, format_name: str, color_index):
        border_description = self._book.add_format()
        border_description.set_bottom(2)
        border_description.set_top(2)
        border_description.set_left(2)
        border_description.set_right(2)
        border_description.set_bg_color(Colors.scales[color_index][9])
        border_description.set_border_color(Colors.scales[color_index][3])
        border_description.set_align('centre')
        border_description.set_align('vcentre')
        border_description.set_text_wrap(True)
        self._reelsets_table_formats[format_name] = border_description

    def _InitReelsetsTable_EvenFormat(self, format_name: str, color_index, align: str):
        border_even = self._book.add_format()
        border_even.set_border(1)
        border_even.set_bg_color(Colors.scales[color_index][-2])
        border_even.set_border_color(Colors.scales[color_index][3])
        border_even.set_align(align)
        border_even.set_align('vcentre')
        border_even.set_text_wrap(True)
        border_even.set_font_name(self._font_name)
        border_even.set_font_size(10)
        self._reelsets_table_formats[format_name] = border_even

    def _InitReelsetsTable_OddFormat(self, format_name: str, color_index, align: str):
        border_odd = self._book.add_format()
        border_odd.set_border(1)
        border_odd.set_bg_color(Colors.scales[color_index][-1])
        border_odd.set_border_color(Colors.scales[color_index][3])
        border_odd.set_align(align)
        border_odd.set_align('vcentre')
        border_odd.set_text_wrap(True)
        border_odd.set_font_name(self._font_name)
        border_odd.set_font_size(10)
        self._reelsets_table_formats[format_name] = border_odd

    def _InitTopAward_DescriptionFormat(self, format_name: str):
        border_description = self._book.add_format()
        border_description.set_bottom(2)
        border_description.set_top(2)
        border_description.set_left(2)
        border_description.set_right(2)
        border_description.set_bg_color(Colors.blue_scale[6])
        border_description.set_border_color(Colors.blue_scale[3])
        border_description.set_align('centre')
        border_description.set_align('vcentre')
        border_description.set_text_wrap(True)
        self._top_award_formats[format_name] = border_description

    def _InitTopAward_EvenFormat(self, format_name: str):
        border_even = self._book.add_format()
        border_even.set_bottom(1)
        border_even.set_top(1)
        border_even.set_left(1)
        border_even.set_right(1)
        border_even.set_bg_color(Colors.blue_scale[-3])
        border_even.set_border_color(Colors.blue_scale[3])
        border_even.set_align('centre')
        border_even.set_align('vcentre')
        border_even.set_text_wrap(True)
        self._top_award_formats[format_name] = border_even

    def _InitTopAward_OddFormat(self, format_name: str):
        border_odd = self._book.add_format()
        border_odd.set_bottom(1)
        border_odd.set_top(1)
        border_odd.set_left(1)
        border_odd.set_right(1)
        border_odd.set_bg_color(Colors.blue_scale[-1])
        border_odd.set_border_color(Colors.blue_scale[3])
        border_odd.set_align('centre')
        border_odd.set_align('vcentre')
        border_odd.set_text_wrap(True)
        self._top_award_formats[format_name] = border_odd

    def _InitConfidence_DescriptionFormat(self, format_name: str, bg_color: str, border_color: str):
        border_description = self._book.add_format()
        border_description.set_bottom(2)
        border_description.set_top(2)
        border_description.set_left(2)
        border_description.set_right(2)
        border_description.set_bg_color(bg_color)
        border_description.set_border_color(border_color)
        border_description.set_align('centre')
        border_description.set_align('vcentre')
        border_description.set_text_wrap(True)
        self._confidence_formats[format_name] = border_description

    def _InitConfidence_CommonFormat(self, format_name: str, bg_color: str, border_color: str):
        border_common = self._book.add_format()
        border_common.set_bottom(1)
        border_common.set_top(1)
        border_common.set_left(1)
        border_common.set_right(1)
        border_common.set_bg_color(bg_color)
        border_common.set_border_color(border_color)
        border_common.set_align('centre')
        border_common.set_align('vcentre')
        border_common.set_text_wrap(True)
        self._confidence_formats[format_name] = border_common

    def _InitLineWins_BorderDescriptionFormat(self, format_name: str, bg_color: str, border_color: str):
        border_description = self._book.add_format()
        border_description.set_bottom(2)
        border_description.set_top(2)
        border_description.set_left(2)
        border_description.set_right(2)
        border_description.set_bg_color(bg_color)
        border_description.set_border_color(border_color)
        border_description.set_align('centre')
        border_description.set_align('vcentre')
        border_description.set_text_wrap(True)
        self._line_wins_formats[format_name] = border_description

    def _InitLineWins_BorderSymbolFormat(self, format_name: str, bg_color: str, border_color: str, borders: list):
        border_symbol_left = self._book.add_format()
        border_symbol_left.set_left(borders[0])
        border_symbol_left.set_top(borders[1])
        border_symbol_left.set_right(borders[2])
        border_symbol_left.set_bottom(borders[3])
        border_symbol_left.set_bg_color(bg_color)
        border_symbol_left.set_border_color(border_color)
        border_symbol_left.set_align('centre')
        border_symbol_left.set_align('vcentre')
        border_symbol_left.set_text_wrap(True)
        self._line_wins_formats[format_name] = border_symbol_left

    def _InitLineWins_BorderLengthFormat(self, format_name: str, bg_color: str, border_color: str, borders: list):
        border_length = self._book.add_format()
        border_length.set_left(borders[0])
        border_length.set_top(borders[1])
        border_length.set_right(borders[2])
        border_length.set_bottom(borders[3])
        border_length.set_bg_color(bg_color)
        border_length.set_border_color(border_color)
        border_length.set_align('centre')
        border_length.set_align('vcentre')
        border_length.set_text_wrap(True)
        self._line_wins_formats[format_name] = border_length

    def _InitLineWins_BorderCommonFormat(self, format_name: str, bg_color: str, border_color: str, borders: list):
        border_common = self._book.add_format()
        border_common.set_left(borders[0])
        border_common.set_top(borders[1])
        border_common.set_right(borders[2])
        border_common.set_bottom(borders[3])
        border_common.set_bg_color(bg_color)
        border_common.set_border_color(border_color)
        border_common.set_align('centre')
        border_common.set_align('vcentre')
        border_common.set_text_wrap(True)
        self._line_wins_formats[format_name] = border_common

    def _InitHeaderFormat(self, format_name: str):
        header_format = self._book.add_format()
        header_format.set_font_name(self._font_name)
        header_format.set_font_size(36)
        header_format.set_font_color("white")
        header_format.set_bold()
        header_format.set_align('centre')
        header_format.set_align('vcentre')
        header_format.set_bg_color(Colors.blue_hexes[1])
        header_format.set_border(2)
        header_format.set_border_color(Colors.blue_hexes[1])
        self._header_formats[format_name] = header_format

    def _InitInnerHeaderFormat(self, format_name: str):
        inner_header_format = self._book.add_format()
        inner_header_format.set_font_name(self._font_name)
        inner_header_format.set_font_size(24)
        inner_header_format.set_font_color("white")
        inner_header_format.set_align('centre')
        inner_header_format.set_align('vcentre')
        inner_header_format.set_bg_color(Colors.blue_hexes[3])
        inner_header_format.set_border(2)
        inner_header_format.set_border_color(Colors.blue_hexes[3])
        self._header_formats[format_name] = inner_header_format

    def _InitSmallHeaderFormat(self, format_name: str):
        small_header_format = self._book.add_format()
        small_header_format.set_font_name(self._font_name)
        small_header_format.set_font_size(20)
        small_header_format.set_font_color("white")
        small_header_format.set_align('centre')
        small_header_format.set_align('vcentre')
        small_header_format.set_bg_color(Colors.blue_hexes[3])
        small_header_format.set_border(2)
        small_header_format.set_border_color(Colors.blue_hexes[3])
        self._header_formats[format_name] = small_header_format

    def _InitBorderCellDescription(self, format_name: str):
        border_description = self._book.add_format()
        border_description.set_bottom(1)
        border_description.set_top(2)
        border_description.set_left(2)
        border_description.set_right(2)
        border_description.set_bg_color(Colors.blue_scale[-3])
        border_description.set_border_color(Colors.blue_scale[1])
        border_description.set_align('centre')
        border_description.set_align('vcentre')
        border_description.set_text_wrap(True)
        self._summary_formats[format_name] = border_description

    def _InitBorderUnderDescription(self, format_name: str):
        border_under_description = self._book.add_format()
        border_under_description.set_bottom(2)
        border_under_description.set_top(1)
        border_under_description.set_left(2)
        border_under_description.set_right(2)
        border_under_description.set_bg_color(Colors.blue_scale[-1])
        border_under_description.set_border_color(Colors.blue_scale[1])
        border_under_description.set_align('centre')
        border_under_description.set_align('vcentre')
        border_under_description.set_text_wrap(True)
        self._summary_formats[format_name] = border_under_description

    def _InitBorderUnderDescription_Features(self, format_name: str):
        border_under_description = self._book.add_format()
        border_under_description.set_bottom(2)
        border_under_description.set_top(1)
        border_under_description.set_left(2)
        border_under_description.set_right(2)
        border_under_description.set_bg_color(Colors.blue_scale[-1])
        border_under_description.set_border_color(Colors.blue_scale[1])
        border_under_description.set_indent(1)
        border_under_description.set_align('left')
        border_under_description.set_align('vcentre')
        border_under_description.set_text_wrap(True)
        self._summary_formats[format_name] = border_under_description

    def _InitBorderBigValues(self, format_name: str):
        border_big_format = self._book.add_format()
        border_big_format.set_bottom(2)
        border_big_format.set_top(2)
        border_big_format.set_left(2)
        border_big_format.set_right(2)
        border_big_format.set_bg_color(Colors.blue_scale[-6])
        border_big_format.set_border_color(Colors.blue_scale[1])
        border_big_format.set_align('centre')
        border_big_format.set_align('vcentre')
        border_big_format.set_text_wrap(True)
        self._summary_formats[format_name] = border_big_format

    def _InitPaytableBorder_Common(self, format_name: str, bg_color: int, border_color: int):
        common_border_format = self._book.add_format()
        common_border_format.set_border_color(Colors.blue_scale[border_color])
        common_border_format.set_bg_color(Colors.blue_scale[bg_color])
        common_border_format.set_bottom(1)
        common_border_format.set_top(1)
        common_border_format.set_left(1)
        common_border_format.set_right(1)
        common_border_format.set_align('centre')
        common_border_format.set_align('vcentre')
        common_border_format.set_text_wrap(True)
        self._paytable_formats[format_name] = common_border_format

    def _InitPytableBorder_Description(self, format_name: str, bg_color: int, border_color: int):
        description_border_format = self._book.add_format()
        description_border_format.set_border_color(Colors.blue_scale[border_color])
        description_border_format.set_bg_color(Colors.blue_scale[bg_color])
        description_border_format.set_bottom(2)
        description_border_format.set_top(2)
        description_border_format.set_left(2)
        description_border_format.set_right(2)
        description_border_format.set_align('centre')
        description_border_format.set_align('vcentre')
        description_border_format.set_text_wrap(True)
        self._paytable_formats[format_name] = description_border_format

    def _InitText_Smallest(self, format_name: str):
        text_smallest = self._book.add_format()
        text_smallest.set_font_name(self._font_name)
        text_smallest.set_font_size(8)
        self._text_formats[format_name] = text_smallest

    def _InitText_SmallestBold(self, format_name: str):
        text_smallest_bold = self._book.add_format()
        text_smallest_bold.set_font_name(self._font_name)
        text_smallest_bold.set_font_size(8)
        text_smallest_bold.set_bold(True)
        self._text_formats[format_name] = text_smallest_bold

    def _InitText_Small(self, format_name: str):
        text_small = self._book.add_format()
        text_small.set_font_name(self._font_name)
        text_small.set_font_size(10)
        self._text_formats[format_name] = text_small

    def _InitText_SmallBold(self, format_name: str):
        text_small_bold = self._book.add_format()
        text_small_bold.set_font_name(self._font_name)
        text_small_bold.set_font_size(10)
        text_small_bold.set_bold(True)
        self._text_formats[format_name] = text_small_bold

    def _InitText_Regular(self, format_name: str):
        text_regular = self._book.add_format()
        text_regular.set_font_name(self._font_name)
        text_regular.set_font_size(12)
        self._text_formats[format_name] = text_regular

    def _InitText_RegularBold(self, format_name: str):
        text_regular_bold = self._book.add_format()
        text_regular_bold.set_font_name(self._font_name)
        text_regular_bold.set_font_size(12)
        text_regular_bold.set_bold(True)
        self._text_formats[format_name] = text_regular_bold

    def _InitText_Big(self, format_name: str):
        text_big = self._book.add_format()
        text_big.set_font_name(self._font_name)
        text_big.set_font_size(14)
        self._text_formats[format_name] = text_big

    def _InitText_BigBold(self, format_name: str):
        text_big_bold = self._book.add_format()
        text_big_bold.set_font_name(self._font_name)
        text_big_bold.set_font_size(14)
        text_big_bold.set_bold(True)
        self._text_formats[format_name] = text_big_bold

    def _InitText_Biggest(self, format_name: str):
        text_biggest = self._book.add_format()
        text_biggest.set_font_name(self._font_name)
        text_biggest.set_font_size(20)
        self._text_formats[format_name] = text_biggest

    def _InitText_BiggestBold(self, format_name: str):
        text_biggest_bold = self._book.add_format()
        text_biggest_bold.set_font_name(self._font_name)
        text_biggest_bold.set_font_size(20)
        text_biggest_bold.set_bold(True)
        self._text_formats[format_name] = text_biggest_bold

    def _InitText_RedInfo(self, format_name: str):
        text_red_info = self._book.add_format()
        text_red_info.set_font_name(self._font_name)
        text_red_info.set_font_size(10)
        text_red_info.set_italic(True)
        text_red_info.set_color('red')
        text_red_info.set_align('centre')
        self._text_formats[format_name] = text_red_info

    def _InitWinlineFormat(self, format_name: str, active_borders: list, bg_color: str = "", border_color: str = ""):
        winline_border_style = 2

        winline_format = self._book.add_format()
        if bg_color != "":
            winline_format.set_bg_color(bg_color)
        winline_format.set_align('centre')
        winline_format.set_align('vcentre')
        # BORDER INDEX:
        # 1 - left
        # 2 - top
        # 3 - right
        # 4 - bottom
        for border_index in active_borders:
            winline_format.set_border_color(border_color)
            if border_index == 1:
                winline_format.set_left(winline_border_style)
            elif border_index == 2:
                winline_format.set_top(winline_border_style)
            elif border_index == 3:
                winline_format.set_right(winline_border_style)
            elif border_index == 4:
                winline_format.set_bottom(winline_border_style)
        self._winline_formats[format_name] = winline_format

    def GetWinlineFormat(self, board_height, board_width, row, col, color: bool = False):
        if color:
            if row == 0:
                if col == 0:
                    return self._winline_formats['corner_left_top_color']
                elif col == board_width - 1:
                    return self._winline_formats['corner_right_top_color']
                else:
                    return self._winline_formats['side_top_color']
            elif row == board_height - 1:
                if col == 0:
                    return self._winline_formats['corner_left_bottom_color']
                elif col == board_width - 1:
                    return self._winline_formats['corner_right_bottom_color']
                else:
                    return self._winline_formats['side_bottom_color']
            else:
                if col == 0:
                    return self._winline_formats['side_left_color']
                elif col == board_width - 1:
                    return self._winline_formats['side_right_color']
                else:
                    return self._winline_formats['centre_color']
        else:
            if row == 0:
                if col == 0:
                    return self._winline_formats['corner_left_top_free']
                elif col == board_width - 1:
                    return self._winline_formats['corner_right_top_free']
                else:
                    return self._winline_formats['side_top_free']
            elif row == board_height - 1:
                if col == 0:
                    return self._winline_formats['corner_left_bottom_free']
                elif col == board_width - 1:
                    return self._winline_formats['corner_right_bottom_free']
                else:
                    return self._winline_formats['side_bottom_free']
            else:
                if col == 0:
                    return self._winline_formats['side_left_free']
                elif col == board_width - 1:
                    return self._winline_formats['side_right_free']
                else:
                    return self._winline_formats['centre_free']

    def _InitRTPDistribution_BorderDescriptionFormat(self, name: str, bg_color: str, border_color: str):
        description_border_format = self._book.add_format()
        description_border_format.set_border_color(border_color)
        description_border_format.set_bg_color(bg_color)
        description_border_format.set_bottom(2)
        description_border_format.set_top(2)
        description_border_format.set_left(2)
        description_border_format.set_right(2)
        description_border_format.set_align('centre')
        description_border_format.set_align('vcentre')
        description_border_format.set_text_wrap(True)
        self._rtp_distribution_formats[name] = description_border_format

    def _InitRTPDistribution_LightFormat(self, name: str, bg_color: str, border_color: str):
        rtp_border_format = self._book.add_format()
        rtp_border_format.set_border_color(border_color)
        rtp_border_format.set_bg_color(bg_color)
        rtp_border_format.set_bottom(1)
        rtp_border_format.set_top(1)
        rtp_border_format.set_left(1)
        rtp_border_format.set_right(1)
        rtp_border_format.set_align('centre')
        rtp_border_format.set_align('vcentre')
        rtp_border_format.set_text_wrap(True)
        self._rtp_distribution_formats[name] = rtp_border_format



