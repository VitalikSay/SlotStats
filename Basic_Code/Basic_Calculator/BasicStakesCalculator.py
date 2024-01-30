from numpy import divide

class Stake_Struct:
    def __init__(self, mult=0, stakes=[], delimiter=", ", divider=0):
        self.mult = mult
        self.stakes = stakes
        self.delimiter = delimiter
        self.divider = divider

    def __str__(self):
        self.stakes.sort()
        if self.divider == 0:
            return "x" + str(self.mult) + "  :  " + self.delimiter.join([str(num) for num in self.stakes])
        else:
            return "x" + str(self.mult) + "  :  " + self.delimiter.join(
                ["{:.2f}".format(num / self.divider) for num in self.stakes])


class Stakes:
    def __init__(self, base_stakes_str, buy_mults, feature_mults, delimiter, divider):
        self.stake_delimiter = delimiter
        self.base_stakes = Stake_Struct(1, [int(num) for num in base_stakes_str.split()], self.stake_delimiter,
                                        divider=divider)
        # self.stake_limit = stake_limit
        self.buy_mults = buy_mults
        self.feature_mults = feature_mults
        self.feature_stakes = []
        self.buy_stakes = []
        self.all_available_stakes = []
        self.divider = divider

    def calc_feature_spin_stakes(self):
        for mult in self.feature_mults:
            current_stake_struct = Stake_Struct(delimiter=self.stake_delimiter, divider=self.divider)
            current_stake_struct.mult = mult
            current_stake_struct.stakes = [num * mult for num in self.base_stakes.stakes]
            self.feature_stakes.append(current_stake_struct)

    def calc_buy_feature_stakes(self):
        for mult in self.buy_mults:
            current_stake_struct = Stake_Struct(delimiter=self.stake_delimiter, divider=self.divider)
            current_stake_struct.mult = mult
            current_stake_struct.stakes = current_stake_struct.stakes = [num * mult for num in self.base_stakes.stakes]
            self.buy_stakes.append(current_stake_struct)

    def calc_all_available_stakes(self):
        all_stakes_list = []
        all_stakes_list += self.base_stakes.stakes
        for struct in self.buy_stakes:
            all_stakes_list += struct.stakes
        for struct in self.feature_stakes:
            all_stakes_list += struct.stakes
        all_stakes = list(set(all_stakes_list))
        all_stakes.sort()
        self.all_available_stakes = all_stakes

    def print_base_stakes(self):
        print("\nbase stakes:" + self.base_stakes)

    def print_feature_spin_stakes(self):
        print("\nfeature spin stakes:")
        for struct in self.feature_stakes:
            print(struct)

    def print_buy_feature_stakes(self):
        print("\nbuy feature stakes:")
        for struct in self.buy_stakes:
            print(struct)

    def print_base_stakes(self):
        print("\nbase stakes:")
        print(self.base_stakes)

    def print_all_available_stakes(self):
        self.calc_all_available_stakes()
        print("\nall stakes:")
        if self.divider == 0:
            print(self.stake_delimiter.join([str(num) for num in self.all_available_stakes]))
        else:
            print(self.stake_delimiter.join(["{:.2f}".format(num / 100) for num in self.all_available_stakes]))


if __name__ == "__main__":
    base_stakes_str = "10 20 50 100 150 200 250 300 400 500 750 1000 1250 1500 2000 2500 5000 7500 10000 12000 14000 16000 18000 20000 22500 25000"
    base_stakes_int = [int(num) for num in base_stakes_str.split()]
    # stake_limit = 10000

    feature_mults = [2]
    buy_mults = [40]

    s = Stakes(base_stakes_str, buy_mults, feature_mults, ", ", 100)
    s.calc_feature_spin_stakes()
    s.calc_buy_feature_stakes()
    s.print_all_available_stakes()
    s.print_base_stakes()
    s.print_feature_spin_stakes()
    s.print_buy_feature_stakes()