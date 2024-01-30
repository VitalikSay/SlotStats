import numpy as np
from Basic_Code.Basic_Calculator.BasicWeightCalculator import BasicWeightCalculator

symbols_weight_calc = BasicWeightCalculator()

symbol_mults = [1,2,3,4,5,6]
weights = symbols_weight_calc.MakeExpWeights(symbol_mults, 2, 6, total_weight=500, min_weight=3)
symbols_weight_calc.PrintWeightsInfo(symbol_mults, weights)
