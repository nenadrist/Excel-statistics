# -*- coding: utf-8 -*-
"""
Created on Fri Mar 24 12:11:19 2017

@author: Nenad
"""

import openpyxl
import glob

# Pick up Excel reports from a folder
cell_values = []
#cell_range = ['C12', 'C32', 'C46', 'C59']
cell_range = ['C12', 'C31', 'C44', 'C56']      # APRIL
file_names = glob.glob("d:\Dropbox\PYTHON\Project - Reports\April 2017\*.xlsx")
#file_names = glob.glob("d:\TEMP\Dropbox\PYTHON\MIT Hacking Medicine Hackathon\Py\Morning Reports/April 2017/*.xlsx")

# Catch figures from particular cells
for i in range(len(file_names)):
    wb = openpyxl.load_workbook(file_names[i], data_only=True, keep_vba=True)
    #sheets = wb.get_sheet_names()
    ws = wb.get_sheet_by_name('MTD Rsd')
    
    for cell in cell_range:
        cell_values.append(round(ws[cell].value, 2))

# Preparing data sets for visualizing
num = 0
graph = []
graph_array = []
for figure in cell_values:
    if num > 3:
        graph_array.append(graph)
        graph = []
        num = 0
    else:
        graph.append(figure)
        num += 1

rooms_values = []
fb_values = []
ood_values = []
other_values = []
# Creating lists grouping the values according to category
rooms_values = cell_values[0::4]
fb_values = cell_values[1::4]
ood_values = cell_values[2::4]
other_values = cell_values[3::4]
        
#print ('CELL VALUES:     ', cell_values)
#print ('ARRAY:      ', graph_array)

# Visualizing...
import matplotlib.pyplot as plt; plt.rcdefaults()
import numpy as np

objects = ('Rooms', 'F&B', 'OOD', 'Other')
y_pos = np.arange(len(objects))
results = graph_array[0]
plt.bar(y_pos, results, color='m', align='center', alpha=0.5)
plt.xticks(y_pos, objects)
plt.xlabel('Revenue centers')
plt.ylabel('RSD')
plt.title('MTD')
plt.show()


x1 = rooms_values #[5, 3, 7, 2, 4, 1]
x2 = fb_values
x3 = ood_values
x4 = other_values

plt.plot(x1, '-b', label='Rooms');
plt.xticks(range(1, len(x1)), range(1, len(x1)));
plt.plot(x2, '-g', label='F&B');
plt.xticks(range(1, len(x2)), range(1, len(x2)));
plt.plot(x3, '-r', label='OOD');
plt.xticks(range(1, len(x3)), range(1, len(x3)));
plt.plot(x4, '-c', label='Other');
plt.xticks(range(1, len(x4)), range(1, len(x4)));

plt.xlabel('Day of Month')
plt.ylabel('RSD')
plt.legend()
plt.show()

