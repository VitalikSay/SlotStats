import numpy as np

pattern_0_mult_win_fr = [4.55, 1/1.31]; w_0 = 0.005; w_sp_0 = 0.01
pattern_1_mult_win_fr = [7.08, 1/1.08]; w_1 = 0.005; w_sp_1 = 0.01
pattern_2_mult_win_fr = [7.19, 1/1.08]; w_2 = 0.005; w_sp_2 = 0.01
pattern_3_mult_win_fr = [3.10, 1/1.73]; w_3 = 0.189; w_sp_3 = 0.17
pattern_4_mult_win_fr = [1.16, 1/3.40]; w_4 = 0.535; w_sp_4 = 0.33
pattern_5_mult_win_fr = [3.13, 1/1.73]; w_5 = 0.117; w_sp_5 = 0.15
pattern_6_mult_win_fr = [5.04, 1/1.29]; w_6 = 0.09; w_sp_6 = 0.18
pattern_7_mult_win_fr = [8.14, 1/1.07]; w_7 = 0.002; w_sp_7 = 0.01
pattern_8_mult_win_fr = [8.14, 1/1.07]; w_8 = 0.002; w_sp_8 = 0.01
pattern_9_mult_win_fr = [4.11, 1/1.54]; w_9 = 0.025; w_sp_9 = 0.06
pattern_10_mult_win_fr = [4.12, 1/1.54]; w_10 = 0.025; w_sp_10 = 0.06

patterns_data = [
    pattern_0_mult_win_fr,
    pattern_1_mult_win_fr,
    pattern_2_mult_win_fr,
    pattern_3_mult_win_fr,
    pattern_4_mult_win_fr,
    pattern_5_mult_win_fr,
    pattern_6_mult_win_fr,
    pattern_7_mult_win_fr,
    pattern_8_mult_win_fr,
    pattern_9_mult_win_fr,
    pattern_10_mult_win_fr,
]

pattern_mults = np.array([val[0] for val in patterns_data], dtype="double")
pattern_win_probs = np.array([val[1] for val in patterns_data], dtype="double")

patterns_weights_base = np.array([w_0, w_1, w_2, w_3, w_4, w_5, w_6, w_7, w_8, w_9, w_10], dtype="double")
patterns_weights_special = np.array([w_sp_0, w_sp_1, w_sp_2, w_sp_3, w_sp_4, w_sp_5, w_sp_6, w_sp_7, w_sp_8, w_sp_9, w_sp_10], dtype="double")

print("Base")
print("Weight sum: ", np.sum(patterns_weights_base))
print("Avg mult:   ", np.sum(pattern_mults * (patterns_weights_base / np.sum(patterns_weights_base))))
print("Avg win fr: ", 1 / np.sum(pattern_win_probs * (patterns_weights_base / np.sum(patterns_weights_base))))

print("\nSpecial")
print("Weight sum: ", np.sum(patterns_weights_special))
print("Avg mult:   ", np.sum(pattern_mults * (patterns_weights_special / np.sum(patterns_weights_special))))
print("Avg win fr: ", 1 / np.sum(pattern_win_probs * (patterns_weights_special / np.sum(patterns_weights_special))))