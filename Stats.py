# SARIMA analysis
# Vitor Péricles de Carvalho
# Laic Asset Management
# versão 0.1 de 2020

# Import section
from statsmodels.tsa.statespace.sarimax import SARIMAX
from random2 import random
import pandas as pd
import numpy as np
import math
import xlrd as x1
import matplotlib.pyplot as plt
import xlsxwriter

# Reads Excel with time series
# no dates column, top row with series name
file = "IPCAalone.xls"
wb = x1.open_workbook(file)
s1 = wb.sheet_by_index(0)
s1.cell_value(0,0)
r = s1.nrows
c = s1.ncols
a1 = s1.cell_value(0,0)
print("SARIMA model by Vitor version 0.1 2020")
print()
print("Current file has time series of")
groups = []
i = 0
while i < c:
    headline = str(s1.cell_value(0,i))
    groups.append(headline)
    i = i+1
print(groups)

# Prepare a loop to fit model for all time series
# on the spreadsheet and forecast next value
runs = 0
fitholder = []
while runs < c:
    data = []
    value = 0
    i = 1
    while i < r:
        value = s1.cell_value(i,runs)
        data.append(value)
        i = i+1
    # print("Last value for",s1.cell_value(0,runs),value)
    # Fit model
    model = SARIMAX(data, order=(1, 1, 1), seasonal_order=(1, 1, 1, 12))
    model_fit = model.fit(start_params=[0, 0, 0, 0, 1],disp=False)
    seer = model_fit.predict(len(data), len(data))
    print("Forecast for next value of",s1.cell_value(0,runs),"=",seer)
    fitholder.append(seer)
    runs = runs+1

workbook = xlsxwriter.Workbook("NextStepForecast.xlsx")
worksheet = workbook.add_worksheet()
runs = 0
while runs < c:
     worksheet.write(0,runs,groups[runs])
     worksheet.write(1, runs, fitholder[runs])
     runs = runs+1
workbook.close()