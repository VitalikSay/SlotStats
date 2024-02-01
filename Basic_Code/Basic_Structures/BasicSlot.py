from Basic_Code.Basic_Structures.BasicReelsets import BasicReelsets
from Basic_Code.Basic_Structures.BasicWinLines import BasicWinlines
from Basic_Code.Basic_Structures.BasicPayTable import BasicPayTable
from Basic_Code.Utils.BasicTimer import timer
import json


class BasicSlot:
    """
    Methods:

    """
    def __init__(self):
        self._base_bet = -1
        self._feature_bet_mult = -1
        self._total_spin_count = -1
        self._winlimit = -1
        self._game_name_short = ""
        self._game_name_full = ""
        self._game_version = ""
        self._line_wins = ""
        self._features_list = []
        self._game_simulationID = ""
        self._simulation_type = ""
        self._section_names = []  # Taken from Reelsets

        self._feature_free_spins_count = -1
        self._feature_respin_strategy = ""
        self._usual_game_feature_avg_mult = -1.0
        self._feature_avg_mult = -1.0

        self._board_height = -1
        self._board_width = -1

        self._winlines = self.MakeWinlinesObj()
        self._paytable = self.MakePaytableObj()
        self._reelsets = self.MakeReelsetsObj()

        self._feature_list = ""  # List of Names of feature in game, used for PAR sheet writer

        self._gamble_active = False
        self._gamble_max_count = -1
        self._gamble_variant = -1
        self._gamble_choice = -1

        self._gamble_variants = ["color", "suit", "random"]
        self._gamble_choices = ["one round", "until the end"]


    def MakeWinlinesObj(self):
        return BasicWinlines()

    def MakeReelsetsObj(self):
        return BasicReelsets()

    def MakePaytableObj(self):
        return BasicPayTable()

    def GetBet(self, MultByMult=True):
        if MultByMult:
            return self._base_bet * self._feature_bet_mult
        else:
            return self._base_bet

    def GetFeatureMult(self):
        return self._feature_bet_mult

    def GetTotalSpinCount(self):
        return self._total_spin_count

    def GetSimulationId(self):
        return self._game_simulationID

    def GetGameName(self, short=True):
        if short:
            return self._game_name_short
        else:
            return self._game_name_full

    def GetVersion(self):
        return self._game_version

    def GetLineWins(self):
        return self._line_wins

    def GetFeaturesList(self):
        return self._feature_list

    def GetFeatureList(self):
        return self._feature_list

    def GetSectionNames(self):
        return self._section_names

    def GetWinlimit(self):
        return self._winlimit

    def GetBoardWidth(self):
        return self._board_width

    def GetBoardHeight(self):
        return self._board_height

    def GetWinlines(self, winline_index: int = -1):
        return self._winlines.GetWinlines(winline_index)

    def GetPaytable(self):
        return self._paytable

    def GetReelsets(self):
        return self._reelsets

    def IsGambleActive(self):
        return self._gamble_active

    def GetGambleChoice(self, index=True):
        if index:
            return self._gamble_choice
        else:
            return self._gamble_choices[self._gamble_choice]

    def GetGambleMaxCount(self):
        return self._gamble_max_count

    def GetGambleVariant(self, index=True):
        if index:
            return self._gamble_variant
        else:
            return self._gamble_variants[self._gamble_variant]

    def GetSimulationTytpe(self):
        return self._simulation_type

    def GetFeatureFreeSpinCount(self):
        return self._feature_free_spins_count

    def GetUsualGameFeatureAvgMult(self):
        return self._usual_game_feature_avg_mult

    def GetFeatureAvgMult(self):
        return self._feature_avg_mult

    @timer
    def ReadSlot_Json(self, data: dict):
        self._base_bet = data["Current Bet"]
        self._feature_bet_mult = data["feature Bet Multiplier"]
        self._total_spin_count = data["Total Spin Count"]
        self._winlimit = data["Winlimit"]
        self._game_name_short = data["Game Name Short"]
        self._game_name_full = data["Game Name Full"]
        self._game_version = data["Game Version"]
        self._line_wins = data["Line Wins"]
        self._feature_list = data.get('Feature List', '')
        self._game_simulationID = data["Simulation Id"]
        self._simulation_type = data["Simulation Type"]
        self._board_height = data["Board Size"]["Height"]
        self._board_width = data["Board Size"]["Width"]
        self._feature_list = data["Feature List"]

        self._feature_free_spins_count = data["Feature Free Spin Count"]
        self._feature_respin_strategy = data["Feature Re-Spin Strategy"]
        self._usual_game_feature_avg_mult = data["Usual Avg Feature Mult"]
        self._feature_avg_mult = data["Feature Avg Feature Mult"]

        self._winlines.ReadWinlines_Json(data["Winlines"])
        self._paytable.ReadPaytable_Json(data["Symbols"])
        self._reelsets.ReadReelsets_Json(data["Reelsets"])
        self._section_names = self._reelsets.GetSectionNames()

        self._gamble_active = data["Gamble"]["Active"]
        self._gamble_choice = data["Gamble"]["Choice"]
        self._gamble_max_count = data["Gamble"]["Max Count"]
        self._gamble_variant = data["Gamble"]["Variant"]


if __name__ == "__main__":
    file = open(r"C:\Users\VitalijSaiganov\PycharmProjects\XLSX_Writer\Data\Json Stats\SSH\86\SSH_usual_86_no_gamble_stats.json")
    data = json.load(file)
    r = BasicSlot()
    r.ReadSlot_Json(data)
    print(r.GetBoardWidth())

