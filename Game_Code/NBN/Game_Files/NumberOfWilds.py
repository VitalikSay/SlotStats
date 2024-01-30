from Basic_Code.Basic_Calculator.BasicWeightCalculator import BasicWeightCalculator

wild_count_probs_base = BasicWeightCalculator()
wild_count_probs_free = BasicWeightCalculator()
wild_count_probs_special = BasicWeightCalculator()

base_ceses = [*range(3, 16)]
free_cases = [*range(3, 16)]
special_cases = [*range(7, 10)]

# Wild BASE:
#wild_base_probs = wild_count_probs_base.MakeExpWeights(base_ceses, 1.2, 5.5, total_weight=11112, min_weight=1)
#wild_count_probs_base.PrintWeightsInfo(base_ceses, wild_base_probs)

# Wild FREE:
wild_free_probs = wild_count_probs_free.MakeExpWeights(free_cases, 1.2, 6.9, total_weight=100_000, min_weight=1)
wild_count_probs_free.PrintWeightsInfo(free_cases, wild_free_probs)