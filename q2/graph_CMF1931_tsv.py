import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import math

plt.rcParams['text.usetex'] = True

cmfs_data = pd.read_excel("../data.xlsx", sheet_name="CMFs")
color_data = pd.read_excel("../q1/tsv_d65.xlsx", sheet_name="Sheet")
cmf1931 = cmfs_data.iloc[:, :4].to_numpy()
white_dot = color_data.iloc[2:3, :2].to_numpy()[0]
color = color_data.iloc[:6, 9:11].to_numpy().T

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
plt.xlabel(r"$x = \frac{X}{X+Y+Z}$", fontdict=font2)
plt.ylabel(r"$y = \frac{Y}{X+Y+Z}$", fontdict=font2)
plt.grid(True)

plt.plot([cmf1931_x[0], cmf1931_x[-1]], [cmf1931_y[0], cmf1931_y[-1]], marker=',', c="purple", linewidth=4)
plt.scatter(cmf1931_x, cmf1931_y, marker='o', edgecolor='k', c=co, cmap="nipy_spectral")
plt.colorbar().set_label(label="Visible Spectrum Wavelength (nm)", size=15)

plt.scatter(white_dot[0], white_dot[1], marker='*', edgecolor='k', c='white', s=200)
plt.scatter(color[0], color[1], marker='o', edgecolor='k', c='white', s=300)

texts = []

x = white_dot[0] * 1.01
y = white_dot[1] * 1.01
text = "WP"
texts.append(plt.text(x, y, text, color="red", fontsize=24))

for i in range(len(color[0])):
    x = color[0][i] * 1.01
    y = color[1][i] * 1.01
    text = str(i + 1)
    texts.append(plt.text(x, y, text, color="red", fontsize=24))

for color_index in range(len(color[0])):
    P = [white_dot[0], white_dot[1]]
    Q = [color[0][color_index], color[1][color_index]]

    a = Q[1] - P[1]
    b = P[0] - Q[0]
    c = -(a * (P[0]) + b * (P[1]))

    min_index = 0
    min_distance = abs((a * cmf1931_x[0] + b * cmf1931_y[0] + c)) / (math.sqrt(a * a + b * b))
    for i in range(len(cmf1931_x)):
        distance = abs((a * cmf1931_x[i] + b * cmf1931_y[i] + c)) / (math.sqrt(a * a + b * b))
        if distance < min_distance:
            min_distance = distance
            min_index = i

        if color_index == 2 and i > 200:
            break

    x = cmf1931_x[min_index] * 1.001
    y = cmf1931_y[min_index] * 1.001
    text = str(co[min_index]) + "nm"
    texts.append(plt.text(x, y, text, color="red", fontsize=28))

    plt.plot([white_dot[0], cmf1931_x[min_index]], [white_dot[1], cmf1931_y[min_index]],
             label='Point ' + str(color_index + 1) + ' - ' + str(co[min_index]) + 'nm', linewidth=3)

plt.legend(fontsize=28)

plt.savefig('cmf1931_tsv.svg', bbox_inches='tight', pad_inches=0.3)
plt.savefig('cmf1931_tsv.png', bbox_inches='tight', pad_inches=0.3)
# plt.show()
