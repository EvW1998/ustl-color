import pandas as pd
import numpy as np
import matplotlib.pyplot as plt


plt.rcParams['text.usetex'] = True

cmfs_data = pd.read_excel("../data.xlsx", sheet_name="CMFs")
cmf1931 = cmfs_data.iloc[:, :4].to_numpy()

cmf1931_sum = cmf1931[:, 1] + cmf1931[:, 2] + cmf1931[:, 3]
cmf1931_x = cmf1931[:, 1] / cmf1931_sum
cmf1931_y = cmf1931[:, 2] / cmf1931_sum
co = np.arange(360, 831)

font1 = {'family': 'serif', 'color': 'blue', 'size': 36}
font2 = {'family': 'serif', 'color': 'darkred', 'size': 36}

plt.figure(figsize=(19, 15))
plt.xlim([0, 1])
plt.ylim([0, 1])
plt.xticks(fontsize=18)
plt.yticks(fontsize=18)
plt.title("CMF 1931 (360nm - 830nm)", fontdict=font1)
plt.xlabel(r"$x = \frac{X}{X+Y+Z}$", fontdict=font2)
plt.ylabel(r"$y = \frac{Y}{X+Y+Z}$", fontdict=font2)
plt.grid(True)

plt.plot([cmf1931_x[0], cmf1931_x[-1]], [cmf1931_y[0], cmf1931_y[-1]], marker=',', c="purple", linewidth=4)
plt.scatter(cmf1931_x, cmf1931_y, marker='o', edgecolor='k', c=co, cmap="nipy_spectral")
plt.colorbar().set_label(label="Visible Spectrum Wavelength (nm)", size=15)

texts = []

x = cmf1931_x[0] * 1.001
y = cmf1931_y[0] * 1.001
text = str(co[0]) + "nm"
texts.append(plt.text(x, y, text, color="red", fontsize=22))

x = cmf1931_x[-1] * 1.001
y = cmf1931_y[-1] * 1.001
text = str(co[-1]) + "nm"
texts.append(plt.text(x, y, text, color="red", fontsize=22))

for i in range(len(cmf1931_x[100:290:15])):
    x = cmf1931_x[100:290:15][i] * 1.001
    y = cmf1931_y[100:290:15][i] * 1.001
    text = str(co[100:290:15][i]) + "nm"
    texts.append(plt.text(x, y, text, color="red", fontsize=22))

plt.savefig('./cmf/cmf1931.svg', bbox_inches='tight', pad_inches=0.3)
plt.savefig('./cmf/cmf1931.png', bbox_inches='tight', pad_inches=0.3)
# plt.show()
