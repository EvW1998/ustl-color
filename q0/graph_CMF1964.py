import pandas as pd
import numpy as np
import matplotlib.pyplot as plt

cmfs_data = pd.read_excel("data.xlsx", sheet_name="CMFs")
cmf1931 = cmfs_data.iloc[:, 5:9].to_numpy()

cmf1931_sum = cmf1931[:, 1] + cmf1931[:, 2] + cmf1931[:, 3]
cmf1931_x = cmf1931[:, 1] / cmf1931_sum
cmf1931_y = cmf1931[:, 2] / cmf1931_sum
co = np.arange(360, 831)

font1 = {'family': 'serif', 'color': 'blue', 'size': 36}
font2 = {'family': 'serif', 'color': 'darkred', 'size': 24}

plt.figure(figsize=(19, 15))
plt.xlim([0, 1])
plt.ylim([0, 1])
plt.xticks(fontsize=18)
plt.yticks(fontsize=18)
plt.title("CMF 1964", fontdict=font1)
plt.xlabel("x / (x + y + z)", fontdict=font2)
plt.ylabel("y / (x + y + z)", fontdict=font2)
plt.grid(True)

plt.scatter(cmf1931_x, cmf1931_y, marker='o', edgecolor='k', c=co, cmap="nipy_spectral")
plt.colorbar().set_label(label="Visible Spectrum Wavelength (nm)", size=15)

texts = []

x = cmf1931_x[0] * 1.001
y = cmf1931_y[0] * 1.001
text = str(co[0]) + "nm"
texts.append(plt.text(x, y, text, color="red", fontsize=14))

x = cmf1931_x[-1] * 1.001
y = cmf1931_y[-1] * 1.001
text = str(co[-1]) + "nm"
texts.append(plt.text(x, y, text, color="red", fontsize=14))

for i in range(len(cmf1931_x[100:290:15])):
    x = cmf1931_x[100:290:15][i] * 1.001
    y = cmf1931_y[100:290:15][i] * 1.001
    text = str(co[100:290:15][i]) + "nm"
    texts.append(plt.text(x, y, text, color="red", fontsize=14))

plt.savefig('cmf1964.svg', bbox_inches='tight', pad_inches=0.5)
plt.savefig('cmf1964.png', bbox_inches='tight', pad_inches=0.5)
# plt.show()
