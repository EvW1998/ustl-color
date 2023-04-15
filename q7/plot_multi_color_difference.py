import pandas as pd
import numpy as np
import matplotlib.pyplot as plt

plt.rcParams['text.usetex'] = True

difference1_data = pd.read_excel("1_th polynomial model.xlsx", sheet_name="Sheet")
difference1 = difference1_data.iloc[:, 5:6].to_numpy().flatten()
avg_difference1 = np.mean(difference1)

difference2_data = pd.read_excel("2_th polynomial model.xlsx", sheet_name="Sheet")
difference2 = difference2_data.iloc[:, 5:6].to_numpy().flatten()
avg_difference2 = np.mean(difference2)

difference3_data = pd.read_excel("3_th polynomial model.xlsx", sheet_name="Sheet")
difference3 = difference3_data.iloc[:, 5:6].to_numpy().flatten()
avg_difference3 = np.mean(difference3)

difference4_data = pd.read_excel("bp_network_model.xlsx", sheet_name="Sheet")
difference4 = difference4_data.iloc[:, 4:5].to_numpy().flatten()
avg_difference4 = np.mean(difference4)

num = np.arange(1, 25)

font1 = {'family': 'serif', 'color': 'blue', 'size': 36}
font2 = {'family': 'serif', 'color': 'darkred', 'size': 36}
plt.figure(figsize=(19, 15))
plt.xlim([0.5, 24.5])
plt.xticks(num)
plt.xticks(fontsize=18)
plt.yticks(fontsize=18)
plt.ylabel(r"Colour Difference $\Delta E^{*}_{ab}$", fontdict=font2)
plt.xlabel("Colour Card Number", fontdict=font2)
plt.grid(True)

plt.plot(num, difference1, 'ro-', linewidth=3, markersize=18)
plt.axhline(y=avg_difference1, color='r', linestyle='--', linewidth=3, label="Polynomial Model n=1, Avg = 18.72")

plt.plot(num, difference2, 'go-', linewidth=3, markersize=18)
plt.axhline(y=avg_difference2, color='g', linestyle='--', linewidth=3, label="Polynomial Model n=2, Avg = 6.96")

plt.plot(num, difference3, 'bo-', linewidth=3, markersize=18)
plt.axhline(y=avg_difference3, color='b', linestyle='--', linewidth=3, label="Polynomial Model n=3, Avg = 3.74")

plt.plot(num, difference4, 'mo-', linewidth=3, markersize=18)
plt.axhline(y=avg_difference4, color='m', linestyle='--', linewidth=3, label="BP Network, Avg = 4.88")

# texts = []
# for i in [2, 4, 8, 10, 13, 16, 19, 20, 23]:
#     x = num[i] - 0.5
#     y = difference[i] + 0.3
#     text = str(round(difference[i], 2))
#     texts.append(plt.text(x, y, text, color="red", fontsize=24))
#
# for i in [0, 3, 6, 9, 12, 15, 18, 21]:
#     x = num[i] - 0.5
#     y = difference[i] - 0.8
#     text = str(round(difference[i], 2))
#     texts.append(plt.text(x, y, text, color="red", fontsize=24))
#
# for i in [1, 7, 11, 14, 17, 22]:
#     x = num[i] - 1.3
#     y = difference[i] - 0.3
#     text = str(round(difference[i], 2))
#     texts.append(plt.text(x, y, text, color="red", fontsize=24))

plt.legend(fontsize=28)
plt.savefig('multi_color_difference.svg', bbox_inches='tight', pad_inches=0.3)
plt.savefig('multi_color_difference.png', bbox_inches='tight', pad_inches=0.3)
