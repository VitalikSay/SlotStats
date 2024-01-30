from Basic_Code.Basic_Structures.BasicSlot import BasicSlot
from Basic_Code.Utils.BasicTimer import timer
import json as j

class HSDSlot(BasicSlot):
    @timer
    def ReadSlot_Json(self, data: dict):
        self.__base_bet = data["Current Bet"]
        self.__feature_bet_mult = data["feature Bet Multiplier"]
        self.__total_spin_count = data["Total Spin Count"]
        self.__winlimit = data["Winlimit"]
        self.__game_name_short = data["Game Name Short"]
        self.__game_name_full = data["Game Name Full"]
        #self.__game_version = data["Game Version"]
        #self.__line_wins = data["Line Wins"]
        self.__feature_list = data['Feature List']
        #self.__game_simulationID = data["Simulation Id"]
        #self.__simulation_type = data["Simulation Type"]
        self.__board_height = data["Board Size"]["Height"]
        self.__board_width = data["Board Size"]["Width"]
        self.__feature_list = data["Feature List"]

        #self.__feature_free_spins_count = data["Feature Free Spin Count"]
        #self.__feature_respin_strategy = data["Feature Re-Spin Strategy"]
        #self.__usual_game_feature_avg_mult = data["Usual Avg Feature Mult"]
        #self.__feature_avg_mult = data["Feature Avg Feature Mult"]

        # self.__winlines.ReadWinlines_Json(data["Winlines"])
        # self.__paytable.ReadPaytable_Json(data["Symbols"])
        # self.__reelsets.ReadReelsets_Json(data["Reelsets"])
        # self.__section_names = self.__reelsets.GetSectionNames()

        self.__gamble_active = data["Gamble"]["Active"]
        self.__gamble_choice = data["Gamble"]["Choice"]
        self.__gamble_max_count = data["Gamble"]["Max Count"]
        self.__gamble_variant = data["Gamble"]["Variant"]

if __name__ == "__main__":
    file = open(
        r"C:\Users\VitalijSaiganov\PycharmProjects\SlotHandler\Game_Source_Data\HSD\90\HSD_usual_90_no_gamble_stats.json")
    data = j.load(file)

    slot = HSDSlot()
    slot.ReadSlot_Json(data)