from itertools import combinations
import numpy as np

# need probs:
scatter_count_5_prob = 1 / 400_000
scatter_count_4_prob = 1 / 15_000
scatter_count_3_prob = 1 / 307
scatter_count_2_prob = 1 / 27
scatter_count_1_prob = 1 / 9


class ReelsetScatterProbs():
    def __init__(self, board_width=5):
        self.reel_probs = [0 for _ in range(board_width)]

    def __fract(self, number):
        if number == 1:
            return 1
        else:
            return number * self.__fract(number-1)

    def __combination(self, amount, total):
        return int(self.__fract(total) / (self.__fract(total-amount) * self.__fract(amount)))
        # Сочетания

    def __generateRandomNumbersWithSum(self, number_of_numbers: int, tot_prob: float = 1, min_prob: float = 0.001, digits_after_dot: int = 3):
        randoms = np.zeros(number_of_numbers, dtype="double")
        while(any(randoms < min_prob)):
            randoms = np.array([np.random.rand() for _ in range(number_of_numbers)])
            randoms /= np.sum(randoms)
            randoms *= tot_prob
            randoms = np.round(randoms, digits_after_dot)
            diff = tot_prob - np.sum(randoms)
            rand_ind = np.random.randint(0, len(randoms)-1)
            randoms[rand_ind] += diff
        return randoms


    def setProbs(self, probs: list):
        assert len(probs) == len(self.reel_probs)
        self.reel_probs = probs

    def setRandomProbs(self, probs_sum=1, min_prob=0.001, digits_after_dot: int = 3):
        self.reel_probs = list(self.__generateRandomNumbersWithSum(len(self.reel_probs), probs_sum, min_prob, digits_after_dot))
        print(self.reel_probs)


    def calcProbOfScatters(self, scatter_count):
        combinations_reel_indexes = combinations([*range(len(self.reel_probs))], scatter_count)

        res = 0
        for comb in combinations_reel_indexes:
            true_indexes = comb
            false_indexes = list(set([*range(len(self.reel_probs))]) - set(true_indexes))
            cur_prob = 1
            for true_index in true_indexes:
                cur_prob *= self.reel_probs[true_index]
            for false_index in false_indexes:
                cur_prob *= 1 - self.reel_probs[false_index]
            res += cur_prob
        return res

    def calcProbsOfAllScatterCombinations(self):
        res = []
        for num in range(len(self.reel_probs)+1):
            res.append(self.calcProbOfScatters(num))
        return  res


if __name__ == "__main__":
    t_all = ReelsetScatterProbs(5)
    t_135 = ReelsetScatterProbs(5)
    t_24 = ReelsetScatterProbs(5)

    print("respin")
    t_respin = ReelsetScatterProbs(5)
    n_respin = 0.0693
    respin_prob = 1/9
    respin_free_trigger_reelset_prob = 1
    t_respin.setProbs([n_respin] * 5)
    probs_respin = respin_prob * respin_free_trigger_reelset_prob * np.array(t_respin.calcProbsOfAllScatterCombinations())
    for i in range(len(probs_respin)):
        print(i, "scatters 1 in ", 1 / probs_respin[i])

    print()
    print("fs triggered 1 in ", 1 / np.sum(probs_respin[3:]))

    print()

    n = 0.07234
    n_all = 0.1183
    n_135 = 0.319
    n_24 = 0.5


    weight_all = 0.1
    weight_135 = 0.05
    weight_24 = 0.05

    t_all.setProbs([n_all] * 5)
    t_135.setProbs([n_135, 0, n_135, 0, n_135])
    t_24.setProbs([0, n_24, 0, n_24, 0])

    probs_all = weight_all * np.array(t_all.calcProbsOfAllScatterCombinations())
    probs_135 = weight_135 * np.array(t_135.calcProbsOfAllScatterCombinations())
    probs_24 = weight_24 * np.array(t_24.calcProbsOfAllScatterCombinations())

    print("base")
    probs_base = probs_all + probs_135 + probs_24
    for i in range(len(probs_base)):
        print(i, "scatters 1 in ", 1 / probs_base[i])
    print()
    print("fs triggered 1 in ", 1 / np.sum(probs_base[3:]))

    print()
    print("base + respin")
    probs_sum = probs_all + probs_135 + probs_24 + probs_respin
    for i in range(len(probs_sum)):
        print(i, "scatters 1 in ", 1 / probs_sum[i])

    print()
    print("fs triggered 1 in ", 1 / np.sum(probs_sum[3:]))








