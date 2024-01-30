import numpy as np


class BasicSpinWin:
    def __init__(self, spin_win: dict, bet: int, total_game_spin_count: int):
        self._spin_win = spin_win
        self._bet = bet
        self._total_game_spin_count = total_game_spin_count

        self._total_feature_win = 0  # in cents
        self._total_count = 0
        self._zero_win_count = 0

        self._rtp_small_decimal = 0  # using total game spin count
        self._rtp_big_decimal = 0  # using feature spin count

        self._max_win_in_cents = 0
        self._max_win_counter = 0
        self._min_win_in_cents = np.inf
        self._min_win_counter = 0

        self._avg_win_in_cents_with_zero = 0
        self._avg_win_in_cents_without_zero = 0

        self._mode_no_zero = 0   # in cents without zero
        self._mode_no_zero_count = 0
        self._mode_with_zero = 0  # in cents
        self._mode_with_zero_count = 0

        self._win_freq_in_feature_decimal = 0
        self._win_freq_in_global_decimal = 0
        self._hit_freq_in_global_decimal = 0

        self._std_all_feature = 0  # in bets
        self._std_reduced_by_number_of_spins = 0  # in bets

    def HahdleSpinWin(self):
        self._max_win_in_cents = max(self._spin_win.keys())
        self._max_win_counter = self._spin_win[self._max_win_in_cents]
        self._min_win_in_cents = min(self._spin_win.keys())
        self._min_win_counter = self._spin_win[self._min_win_in_cents]
        self._total_count = sum(self._spin_win.values())
        self._total_feature_win = sum([win * count for win, count in self._spin_win.items()])
        self._zero_win_count = self._spin_win.get(0, 0)

        self._rtp_small_decimal = self._total_feature_win / self._bet / self._total_game_spin_count
        self._rtp_big_decimal = self._total_feature_win / self._bet / self._total_count

        self._avg_win_in_cents_without_zero = self._total_feature_win / (self._total_count - self._zero_win_count)
        self._avg_win_in_cents_with_zero = self._total_feature_win / self._total_count

        self._win_freq_in_feature_decimal = (self._total_count - self._zero_win_count) / self._total_count
        self._win_freq_in_global_decimal = (self._total_count - self._zero_win_count) / self._total_game_spin_count
        self._hit_freq_in_global_decimal = self._total_count / self._total_game_spin_count

        temp_win_count_lst = list(self._spin_win.items())
        temp_win_count_lst.sort(key=lambda val: val[1])

        self._mode_with_zero_count = temp_win_count_lst[0][1]
        self._mode_with_zero = temp_win_count_lst[0][0]
        if temp_win_count_lst[0][0] == 0:
            self._mode_no_zero = temp_win_count_lst[1][0]
            self._mode_no_zero_count = temp_win_count_lst[1][1]
        else:
            self._mode_no_zero = temp_win_count_lst[0][0]
            self._mode_no_zero_count = temp_win_count_lst[0][1]

        dispersion_numerator = sum([np.power(val - self._avg_win_in_cents_with_zero, 2) * count for val, count in self._spin_win.items()])
        self._std_all_feature = np.power(dispersion_numerator / (self._total_count - 1), .5) / self._bet
        self._std_reduced_by_number_of_spins = np.power(dispersion_numerator / (self._total_game_spin_count - 1), .5) / self._bet

    def GetRTP(self, small: bool = True, decimal: bool = True):
        if small:
            return self._rtp_small_decimal * (100 if not decimal else 1)
        return self._rtp_big_decimal * (100 if not decimal else 1)

    def GetMaxWinCents(self):
        return self._max_win_in_cents

    def GetMaxWinCount(self):
        return self._max_win_counter

    def GetMinWinCents(self):
        return self._min_win_in_cents

    def GetMinWinCount(self):
        return self._min_win_counter

    def GetAvgWinInCents(self, with_zero_win: bool = True):
        if with_zero_win:
            return self._avg_win_in_cents_with_zero
        return self._avg_win_in_cents_without_zero

    def GetWinFreq(self, globl: bool = False, decimal: bool = True):
        if globl:
            return self._win_freq_in_global_decimal * (100 if not decimal else 1)
        return self._win_freq_in_feature_decimal * (100 if not decimal else 1)

    def GetSTD(self, glob: bool = True):
        if glob:
            return self._std_reduced_by_number_of_spins
        return self._std_all_feature  # in bets

    def GetModeWinInCents(self, with_zero: bool = True):
        if with_zero:
            return self._mode_with_zero
        return self._mode_no_zero

    def GetModeCount(self, with_zero: bool = True):
        if with_zero:
            return self._mode_with_zero_count
        return self._mode_no_zero_count



