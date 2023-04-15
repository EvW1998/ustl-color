import pandas as pd
import numpy as np
import matplotlib.pyplot as plt

plt.rcParams['text.usetex'] = True

difference_data = pd.read_excel("delta_e.xlsx", sheet_name="Sheet")
difference = difference_data.iloc[:, 1:2].to_numpy().flatten()
avg_difference = np.mean(difference)
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

plt.plot(num, difference, 'bo-', linewidth=3, markersize=18)
plt.axhline(y=avg_difference, color='b', linestyle='--', linewidth=2, label="Average Difference = 13.51")

texts = []
for i in [2, 4, 8, 10, 13, 16, 19, 20, 23]:
    x = num[i] - 0.5
    y = difference[i] + 0.3
    text = str(round(difference[i], 2))
    texts.append(plt.text(x, y, text, color="red", fontsize=24))

for i in [0, 3, 6, 9, 12, 15, 18, 21]:
    x = num[i] - 0.5
    y = difference[i] - 0.8
    text = str(round(difference[i], 2))
    texts.append(plt.text(x, y, text, color="red", fontsize=24))

for i in [1, 7, 11, 14, 17, 22]:
    x = num[i] - 1.3
    y = difference[i] - 0.3
    text = str(round(difference[i], 2))
    texts.append(plt.text(x, y, text, color="red", fontsize=24))

plt.legend(fontsize=28)
plt.savefig('color_difference.svg', bbox_inches='tight', pad_inches=0.3)
plt.savefig('color_difference.png', bbox_inches='tight', pad_inches=0.3)
