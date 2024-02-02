import pandas as pd

from Basic_Code.Basic_Structures.BasicSlot import BasicSlot
from Basic_Code.Basic_Structures.BasicStatistics import BasicStatistics
from Basic_Code.Basic_Calculator.BasicStatsCalculator import BasicStatsCalculator
from Basic_Code.Utils.BasicPathHadler import BasicPathHandler
from PIL import Image

import os
import matplotlib.pyplot as plt
from matplotlib import font_manager as fm
import seaborn as sns
import json as j
import numpy as np

class Colors:
    blue_hexes = ["#190280", "#05099D", "#0833BA", "#0D68D5", "#13A6F0", "#2FD8F5", "#4DF9F3", "#6BFDDC", "#8AFFCF", "#AAFFCD", "#CCFFD7"]
    dark_hexes = ["#25182A", "#2A1F34", "#2D263E", "#2F2C47", "#333651", "#4E5669", "#697480", "#849198", "#9FACAF", "#BAC6C6", "#D6DDDB"]
    green_hexes = ["#002C72", "#00568D", "#008AA6", "#02BEB6", "#06D6A0", "#24DF8B", "#44E67F", "#66ED7C", "#88F288", "#B6F7AA", "#DCFBCC"]
    salmon_hexes = ["#80801F", "#9F8F28", "#BF8B33", "#DF7F3E", "#FF6B4A", "#FF6060", "#FF7691", "#FF8DBA", "#FFA5DB", "#FFBEF2", "#FFD7FF"]
    golden_hexes = ["#55802E", "#819F3B", "#B4BF49", "#DFD557", "#FFD166", "#FFBF78", "#FFB28B", "#FFAD9F", "#FFB3B7", "#FFC8D6", "#FFDDED"]
    crayola_hexes = ["#72337C", "#7B4199", "#7D4FB6", "#7A5ED2", "#716EEE", "#8092F2", "#92B5F6", "#A4D1F9", "#B7E8FB", "#CBF8FD", "#DFFEFB"]
    navy_hexes = ["#0F0033", "#05003F", "#00044B", "#001856", "#003160", "#225D77", "#44848D", "#66A39E", "#88B8AB", "#AACDBC", "#CCE1D3"]
    gray_hexes = ["#7D7B7E", "#9B9A9E", "#B9B8BD", "#D7D8DD", "#F6F8FC", "#F7FAFD", "#F8FCFD", "#F9FDFD", "#FBFEFE", "#FCFEFE", "#FDFFFE"]
    mandy_hexes = ["#786123", "#94622D", "#B05D38", "#CB5044", "#E55062", "#EB658E", "#F07BB5", "#F492D4", "#F8A9ED", "#F8C1FB", "#F4D9FD"]
    jade_hexes = ["#00265A", "#00486F", "#007383", "#009789", "#00A976", "#22B76D", "#44C46B", "#66D171", "#88DC88", "#B5E7AA", "#DAF1CC"]

    # blue_scale = ["#063248", "#07405D", "#094F72", "#0B5D87", "#0C6C9C", "#0E7AB1", "#1089C6", "#1198DB", "#13A6F0", "#2EB0F2", "#48BAF3", "#63C4F5", "#7DCEF7", "#98D8F8", "#B2E2FA", "#CDECFC", "#E7F6FE"]

    blue_scale = ["#0B6490", "#0D6FA0", "#0E7AB0", "#0F85C0", "#1190D0", "#129BE0", "#13A6F0", "#28AEF1", "#3DB6F3", "#52BEF4", "#67C6F5", "#7CCEF7", "#91D6F8", "#A6DDF9", "#BBE5FB", "#D0EDFC"]
    green_scale = ["#048060", "#048F6B", "#049D75", "#05AB80", "#05BA8B", "#06C895", "#06D6A0", "#1CDAA8", "#32DDB1", "#48E1B9", "#5FE5C2", "#75E8CA", "#8BECD3", "#A1F0DB", "#B7F3E4", "#CDF7EC"]
    salmon_scale = ["#99402C", "#AA4731", "#BB4F36", "#CC563B", "#DD5D40", "#EE6445", "#FF6B4A", "#FF785A", "#FF856A", "#FF937A", "#FFA08A", "#FFAD9A", "#FFBAAB", "#FFC7BB", "#FFD4CB", "#FFE1DB"]
    golden_scale = ["#997D3D", "#AA8B44", "#BB994B", "#CCA752", "#DDB558", "#EEC35F", "#FFD166", "#FFD574", "#FFD981", "#FFDD8F", "#FFE19C", "#FFE5AA", "#FFEAB8", "#FFEEC5", "#FFF2D3", "#FFF6E0"]
    crayola_scale = ["#44428F", "#4B499F", "#5351AF", "#5A58BE", "#625FCE", "#6A67DE", "#716EEE", "#7E7BF0", "#8A88F1", "#9795F3", "#A4A2F4", "#B0AEF6", "#BDBBF7", "#C9C8F9", "#D6D5FA", "#E3E2FC"]
    mandy_scale = ["#89303B", "#993541", "#A83B48", "#B7404E", "#C74555", "#D64B5C", "#E55062", "#E76070", "#EA6F7E", "#EC7F8C", "#EE8E9A", "#F19EA8", "#F3ADB6", "#F5BDC4", "#F8CCD2", "#FADCE0"]
    jade_scale = ["#006547", "#00714F", "#007C57", "#00875E", "#009366", "#009E6E", "#00A976", "#17B182", "#2DB88E", "#44C09B", "#5BC8A7", "#71CFB3", "#88D7BF", "#9FDFCB", "#B5E6D7", "#CCEEE4"]

    hexes = [blue_hexes, green_hexes, salmon_hexes, golden_hexes, crayola_hexes, mandy_hexes, jade_hexes]
    scales = [blue_scale, green_scale, salmon_scale, golden_scale, crayola_scale, mandy_scale, jade_scale]

    @staticmethod
    def str_hex_to_int_vec(hex: str):
        assert hex[0] == '#'
        r, g, b = int(hex[1:3], 16), int(hex[3:5], 16), int(hex[5:7], 16)
        return np.array([r, g, b], dtype='int')

    @staticmethod
    def int_vec_to_hex_str(rgb: np.array):
        assert len(rgb) == 3
        rgb[0] = [rgb[0], 255][int(rgb[0] > 255)]
        rgb[1] = [rgb[1], 255][int(rgb[1] > 255)]
        rgb[2] = [rgb[2], 255][int(rgb[2] > 255)]
        rgb[0] = [rgb[0], 0][int(rgb[0] < 0)]
        rgb[1] = [rgb[1], 0][int(rgb[1] < 0)]
        rgb[2] = [rgb[2], 0][int(rgb[2] < 0)]
        r_str = hex(rgb[0])[2:]
        r_str = '0'+r_str.upper() if len(r_str) == 1 else r_str.upper()
        g_str = hex(rgb[1])[2:]
        g_str = '0' + g_str.upper() if len(g_str) == 1 else g_str.upper()
        b_str = hex(rgb[2])[2:]
        b_str = '0' + b_str.upper() if len(b_str) == 1 else b_str.upper()
        return '#'+r_str+g_str+b_str

    @staticmethod
    def get_vector_of_scales(start_hex_str: str, end_hex_str: str, number_of_scales: int):
        start_rgb = Colors.str_hex_to_int_vec(start_hex_str)
        end_rgb = Colors.str_hex_to_int_vec(end_hex_str)
        scale_rgb = end_rgb - start_rgb
        assert number_of_scales >= 2
        number_of_parts = number_of_scales - 2 + 1

        one_step = np.floor(scale_rgb / number_of_parts)
        res = []
        for i in range(number_of_scales):
            cur_vec = np.array(start_rgb + i * one_step, dtype='int')
            res.append(Colors.int_vec_to_hex_str(cur_vec))
        return res

    @staticmethod
    def get_scales_for_numbers(numbers, start_hex_str: str, end_hex_str: str):
        max_num = max(numbers)
        min_num = min(numbers)

        start_rgb = Colors.str_hex_to_int_vec(start_hex_str)
        end_rgb = Colors.str_hex_to_int_vec(end_hex_str)
        scale_rgb = end_rgb - start_rgb
        assert len(numbers) >= 2

        res = []
        for i, number in enumerate(numbers):
            if number == max_num:
                res.append(end_hex_str)
                continue
            elif number == min_num:
                res.append(start_hex_str)
                continue
            else:
                scale = (number - min_num) / (max_num - min_num)
                cur_vec = np.array(start_rgb + scale * scale_rgb, dtype='int')
                res.append(Colors.int_vec_to_hex_str(cur_vec))
        return res




class BasicGraphMaker:
    def __init__(self, slot: BasicSlot, stats: BasicStatistics, calculation: BasicStatsCalculator):
        self._section_names = slot.GetSectionNames()
        self._section_spin_dfs = calculation.GetSectionAllSpinWinsDF()
        self._bet = slot.GetBet()
        self._total_spin_count = stats.GetTotalSpinCount()

        self._total_spin_win_df = calculation.GetTotalSpinWinDF()
        self._section_spin_win_dfs = calculation.GetSectionFeatureSpinWinsDF()
        self._section_names = slot.GetSectionNames()
        self._paytable = slot.GetPaytable()

        self._path_handler = BasicPathHandler()
        self._temp_graphs_folder_path = self._path_handler.GetTempGraphsFolderPath(slot.GetGameName(), slot.GetVersion())
        self._pie_plot_name = "summary_pie_plot.png"
        self._spin_win_rtp_distribution = "rtp_distribution"
        self._simple_bar_name = 'simple_bar'
        self._rtp_distribution_by_symbol = "rtp_by_symbol"
        self._rtp_distribution_by_winline_len = "rtp_by_winline_len"

        self.CleanDir()

    def CleanDir(self):
        for file_name in os.listdir(self._temp_graphs_folder_path):
            os.remove(os.path.join(self._temp_graphs_folder_path, file_name))

    def _MakePiePlot_RTPbySections(self):
        rtps = []
        section_names = self._section_names.copy()
        for i, section_df in enumerate(self._section_spin_dfs):
            section_df_sums = section_df.sum()
            current_section_rtp_perc = section_df_sums["total_win"] / self._total_spin_count / self._bet * 100
            rtps.append(current_section_rtp_perc)
        total_rtp = sum(rtps)
        rtps.append(100 - total_rtp)
        explode = [0.02 for _ in range(len(rtps))]
        explode[-1] = 0.1
        section_names.append("Game Hold")

        # define Seaborn color palette to use
        colors = sns.color_palette(Colors.blue_hexes[-2::-1])

        fig = plt.figure(figsize=(7.1, 6))
        ax = fig.add_axes([0.1, 0.1, 0.8, 0.8])

        patches, texts, autotexts = ax.pie(rtps, explode=explode, labels=section_names, colors=colors, autopct='%.2f%%',
                                           pctdistance=0.7, labeldistance=1.05, radius=1.1)

        fontprop = fm.FontProperties()
        fontprop.set_size(15)
        plt.setp(autotexts, fontproperties=fontprop)
        plt.setp(texts, fontproperties=fontprop)
        plt.savefig(os.path.join(self._temp_graphs_folder_path, self._pie_plot_name), dpi=300)
        plt.close(fig)

    def MakePiePlot_RTPbySymbol(self, values: dict, section_index: int):
        symbol_ids = [symbol_id for symbol_id in values.keys()]
        symbol_ids.sort()
        symbol_names = [' '.join(self._paytable.GetSymbolName(symbol_id).split('_'))+' (ID: '+str(symbol_id)+')' for symbol_id in symbol_ids]
        symbol_vals = [values[symbol_id] for symbol_id in symbol_ids]

        explode = [0.02 for _ in range(len(symbol_names))]

        # define Seaborn color palette to use
        colors = sns.color_palette(Colors.scales[section_index][6::])

        fig = plt.figure(figsize=(13, 8))
        ax = fig.add_axes([0.1, 0.1, 0.8, 0.8])

        patches, texts, autotexts = ax.pie(symbol_vals, explode=explode, labels=symbol_names, colors=colors, autopct='%.2f%%',
                                           pctdistance=0.8, labeldistance=1.1, radius=1.1)

        fontprop = fm.FontProperties()
        fontprop.set_size(18)
        plt.setp(autotexts, fontproperties=fontprop)
        plt.setp(texts, fontproperties=fontprop)
        plt.savefig(os.path.join(self._temp_graphs_folder_path, self._rtp_distribution_by_symbol + '_' + str(section_index) + '.png'), dpi=300)
        plt.close(fig)

    def MakePiePlot_RTPbyWinlineLen(self, values: dict, section_index: int):
        winline_lens = [win_len for win_len in values.keys()]
        winline_lens.sort()
        winline_len_names = ['Length ' + str(i) for i in winline_lens]
        vals = [values[ln] for ln in winline_lens]

        explode = [0.02 for _ in range(len(winline_lens))]

        # define Seaborn color palette to use
        colors = sns.color_palette(Colors.scales[section_index][5::4])

        fig = plt.figure(figsize=(10, 8))
        ax = fig.add_axes([0.1, 0.1, 0.8, 0.8])

        patches, texts, autotexts = ax.pie(vals, explode=explode, labels=winline_len_names, colors=colors, autopct='%.2f%%',
                                           pctdistance=0.7, labeldistance=1.1, radius=1.1)

        fontprop = fm.FontProperties()
        fontprop.set_size(22)
        plt.setp(autotexts, fontproperties=fontprop)
        plt.setp(texts, fontproperties=fontprop)
        plt.savefig(os.path.join(self._temp_graphs_folder_path, self._rtp_distribution_by_winline_len + '_' + str(section_index) + '.png'), dpi=300)
        plt.close(fig)

    def GetPiePlotRTPbySymbols(self, section_index: int):
        return os.path.join(self._temp_graphs_folder_path, self._rtp_distribution_by_symbol + '_' + str(section_index) + '.png')

    def GetPiePlotRTPbyWinlineLen(self, section_index: int):
        return os.path.join(self._temp_graphs_folder_path, self._rtp_distribution_by_winline_len + '_' + str(section_index) + '.png')

    def MakeBarPlot_SimpleRTPDistribution(self, spin_win_df, name: str, color_index: int):
        fig, ax = plt.subplots(figsize=(6, 12))

        colors = sns.color_palette(Colors.scales[color_index][::-1])

        ax.spines['right'].set_visible(False)
        ax.spines['top'].set_visible(False)
        ax.spines['left'].set_visible(False)
        ax.spines['bottom'].set_visible(False)

        ax.grid(False)

        sns.set(style="whitegrid")
        graph = sns.barplot(spin_win_df, x='rtp', y=spin_win_df.index, palette=colors, hue=spin_win_df.index, legend=False)
        graph.set(xticklabels=[], yticklabels=[])
        graph.tick_params(bottom=False, left=False)
        graph.set(xlabel=None)
        fig.subplots_adjust(top=1, bottom=0, left=0, right=1)
        plt.savefig(os.path.join(self._temp_graphs_folder_path, self._spin_win_rtp_distribution + "_" + name + ".png"), dpi=300)
        plt.close(fig)

    def MakeSimpleBarPlot(self,
                          data: pd.DataFrame,
                          count_col_name: str,
                          graph_name: str,
                          include_zero_case: bool = False,
                          color_index: int = 0,
                          bar_width: float = 0.8,
                          figsize: tuple = (5, 5),
                          dpi: int = 300,
                          transpose: bool = False,
                          str_colors: list = []):
        fig, ax = plt.subplots(figsize=figsize)

        colors = sns.color_palette(Colors.scales[color_index][::-1])
        if str_colors != []:
            colors = sns.color_palette(str_colors)
        if len(data[count_col_name]) > len(Colors.scales[color_index]) and str_colors == []:
            scales = Colors.get_vector_of_scales(Colors.hexes[color_index][-1], Colors.hexes[color_index][4], len(data[count_col_name]))
            colors = sns.color_palette(scales)

        ax.spines['right'].set_visible(False)
        ax.spines['top'].set_visible(False)
        ax.spines['left'].set_visible(False)
        ax.spines['bottom'].set_visible(False)

        ax.grid(False)

        my_data = pd.DataFrame(columns=[count_col_name], index=data.index)
        my_data[count_col_name] = data[count_col_name]
        if not include_zero_case and 0 in my_data.index:
            my_data[count_col_name].loc[0] = 0

        sns.set(style="whitegrid")
        graph = sns.barplot(my_data, x=my_data.index, y=count_col_name, palette=colors[:len(my_data[count_col_name])],
                            orient='v', width=bar_width, hue=count_col_name, legend=False)
        graph.set(xticklabels=[], yticklabels=[])
        graph.tick_params(bottom=False, left=False)
        graph.set(xlabel=None)
        fig.subplots_adjust(top=1, bottom=0, left=0, right=1)

        full_picture_path = self.GetSimpleBarPath(graph_name)
        plt.savefig(full_picture_path, dpi=dpi)
        plt.close(fig)

        picture = Image.open(full_picture_path)
        width, height = picture.size
        if transpose:
            picture = picture.transpose(Image.TRANSPOSE)
            width, height = picture.size
        picture.save(full_picture_path)

        return width, height

    def GetXYScale(self,
                   picture_width: str,
                   picture_height: str,
                   need_width_in_columns: int,
                   need_height_in_rows: int,
                   cell_width_pixels: int = 64,
                   cell_height_pixels: int = 30):
        available_width_pixels = cell_width_pixels * need_width_in_columns
        available_height_pixels = cell_height_pixels * need_height_in_rows
        return available_width_pixels / picture_width, available_height_pixels / picture_height

    def GetSimpleBarPath(self, name: str):
        return os.path.join(self._temp_graphs_folder_path, self._simple_bar_name + "_" + name + ".png")

    def GetSummaryPiePlotPath(self):
        return os.path.join(self._temp_graphs_folder_path, self._pie_plot_name)

    def GetSpinWinRTPPath(self, name: str):
        return os.path.join(self._temp_graphs_folder_path,
                            self._spin_win_rtp_distribution + "_" + name + ".png")

    def MakePARPlots(self):
        self._MakePiePlot_RTPbySections()


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

    graph = BasicGraphMaker(slot, stats, calculation)

    color = Colors()
    print("Orig hex: ", Colors.scales[0][0])
    int_vec = color.str_hex_to_int_vec(Colors.scales[0][0])
    print("Int base 10 vec: ", int_vec)
    print("Converted hex: ", color.int_vec_to_hex_str(int_vec))

    numbers = np.array([1, 7, 5, 2, 10])
    # КРАСИВО: Colors.hexes[3][5], Colors.hexes[1][5]
    # КРАСИВО: Colors.hexes[2][4], Colors.hexes[3][4]
    scales = Colors.get_scales_for_numbers(numbers, Colors.hexes[2][5], Colors.hexes[3][5])
    data = pd.DataFrame(data=numbers, columns=['count'])
    colors = sns.color_palette(scales)
    sns.barplot(data, x=data.index, y='count', palette=colors, hue='count', legend=False)
    plt.show()


