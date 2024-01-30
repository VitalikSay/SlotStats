import numpy as np

base_patterns_weights = np.array([10,8,8,140,514,123,129,4,4,30,30])
base_avg_mults = np.array([4.374, 6.879, 7.044, 2.911, 1.07, 2.958, 4.866, 7.964, 7.976, 3.909, 3.92])

print(np.sum(base_patterns_weights * base_avg_mults / np.sum(base_patterns_weights)))