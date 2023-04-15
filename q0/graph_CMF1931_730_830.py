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
plt.xticks(fontsize=24)
plt.yticks(fontsize=24)
plt.title("CMF 1931 (730nm - 830nm STEP 5)", fontdict=font1)
plt.xlabel(r"$x = \frac{X}{X+Y+Z}$", fontdict=font2)
plt.ylabel(r"$y = \frac{Y}{X+Y+Z}$", fontdict=font2)
plt.grid(True)

plt.scatter(cmf1931_x[-101::5], cmf1931_y[-101::5], marker='o', edgecolor='k', c=co[-101::5],
            vmin=360, vmax=830, cmap="nipy_spectral")

plt.grid(True)
plt.colorbar().set_label(label="Visible Spectrum Wavelength (nm)", size=15)

texts = []

for i in range(len(cmf1931_x[-101::20])):
    x = cmf1931_x[-101::20][i]
    y = cmf1931_y[-101::20][i]
    text = str(co[-101::20][i]) + "nm"
    texts.append(plt.text(x, y, text, color="red", fontsize=22))

plt.savefig('./cmf/cmf1931_730_830.svg', bbox_inches='tight', pad_inches=0.3)
plt.savefig('./cmf/cmf1931_730_830.png', bbox_inches='tight', pad_inches=0.3)
# plt.show()
