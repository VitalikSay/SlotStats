import numpy as np
import matplotlib.pyplot as plt


class BasicWeightCalculator:
    def __init__(self):
        pass


    def MakeExpWeights(self, possible_cases: list, power: float, exp_coef: float, total_weight:int = 1000, min_weight: int = 3):
        original_cases = np.array(possible_cases, dtype="double")
        working_cases = original_cases.copy()

        working_cases *= power
        y = (np.e * (1/exp_coef)) ** working_cases
        y /= np.sum(y)
        rounded = np.round(y * total_weight, 0).astype("int32")

        # Выравнивание до полной массы:
        diff = total_weight - np.sum(rounded)
        max_case_index = rounded.argmax()
        rounded[max_case_index] += diff

        # Выравнивание до минимальных значений:
        indexes_that_less_than_min_weight = rounded < min_weight
        less_than_min_weight_items = np.sum(indexes_that_less_than_min_weight)
        if less_than_min_weight_items:
            needed_weight = less_than_min_weight_items * min_weight
            exist_weight = np.sum(rounded[indexes_that_less_than_min_weight])
            needed_weight -= exist_weight
            rounded[indexes_that_less_than_min_weight] = min_weight
            max_case_index = rounded.argmax()
            rounded[max_case_index] -= needed_weight

        return rounded

    def PrintWeightsInfo(self, cases: list, weights: np.array):
        orig_cases = np.array(cases)
        print()
        print("Probabilities:   ", weights / np.sum(weights))
        print("Weights:         ", weights)
        print()
        print("Probabilities Sum: ", np.sum(weights / np.sum(weights)))
        print("Weight Sum:        ", np.sum(weights))
        print()
        print("Math expectation: ", np.sum(orig_cases * (weights / np.sum(weights))))
        plt.plot(orig_cases, weights / np.sum(weights))
        plt.show()


if __name__ == "__main__":
    cases = [*range(3, 16)]

    r = BasicWeightCalculator()


    weights = r.MakeExpWeights(cases, 1.2, 5.5, total_weight=12500, min_weight=1)
    r.PrintWeightsInfo(cases, weights)