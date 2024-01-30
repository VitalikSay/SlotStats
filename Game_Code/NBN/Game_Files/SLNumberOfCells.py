from Basic_Code.Basic_Calculator.BasicWeightCalculator import BasicWeightCalculator
import numpy as np

bet_mults_border_before_elimination = np.array([1, 2, 3, 4, 5, 10, 15, 20, 25])
bet_mults_centree_before_elimination = np.array([1, 2, 3, 4, 5, 10, 15, 20, 25])

bet_mults_border_after_elimination = np.array([2, 3, 4, 5, 10, 15, 20, 25])
bet_mults_centree_after_elimination = np.array([2, 3, 4, 5, 10, 15, 20, 25])

weights_border_before_elimination = BasicWeightCalculator()
weights = weights_border_before_elimination.MakeExpWeights(bet_mults_border_before_elimination, 2, 4, 10000, 1)
weights_border_before_elimination.PrintWeightsInfo(bet_mults_border_before_elimination, weights)
