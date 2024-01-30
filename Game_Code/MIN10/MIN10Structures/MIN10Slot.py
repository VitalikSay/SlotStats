from Basic_Code.Basic_Structures.BasicSlot import BasicSlot

class MIN10Slot(BasicSlot):
    def ReadSlot_Json(self, data: dict):
        self._base_bet = data["Current Bet"]
        self._feature_bet_mult = data["feature Bet Multiplier"]
        self._total_spin_count = data["Total Spin Count"]
        self._winlimit = data["Winlimit"]
        self._game_name_short = data["Game Name Short"]
        self._game_name_full = data["Game Name Full"]
        self._game_version = data["Game Version"]
        self._line_wins = data["Line Wins"]
        self._feature_list = data['Feature List']
        self._game_simulationID = data["Simulation Id"]
        self._simulation_type = data["Simulation Type"]
        self._board_height = data["Board Size"]["Height"]
        self._board_width = data["Board Size"]["Width"]
        self._feature_list = data["Feature List"]

        # self._feature_free_spins_count = data["Feature Free Spin Count"]
        # self._feature_respin_strategy = data["Feature Re-Spin Strategy"]
        # self._usual_game_feature_avg_mult = data["Usual Avg Feature Mult"]
        # self._feature_avg_mult = data["Feature Avg Feature Mult"]

        self._winlines.ReadWinlines_Json(data["Winlines"])
        self._paytable.ReadPaytable_Json(data["Symbols"])
        self._reelsets.ReadReelsets_Json(data["Reelsets"])
        self._section_names = self._reelsets.GetSectionNames()

        self._gamble_active = data["Gamble"]["Active"]
        self._gamble_choice = data["Gamble"]["Choice"]
        self._gamble_max_count = data["Gamble"]["Max Count"]
        self._gamble_variant = data["Gamble"]["Variant"]