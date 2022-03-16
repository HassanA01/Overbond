import matplotlib.pyplot as plt
import pandas as pd
import xlrd

workbook = xlrd.open_workbook("Book1.xls")      # Importing worksheet
worksheet = workbook.sheet_by_name("Sheet1")
i = 7
c = 10
DI = worksheet.cell_value(7, 5)             # Setting starting point for parsing data.
BP = worksheet.cell_value(10, 7)
AP = worksheet.cell_value(10, 8)
PL = worksheet.cell_value(10, 9)

DIs = []
BPr = []
API = []
PI = []

while DI != '':                 # Parsing the four required data piece from the data sheet.
    DIs.append(DI[3:])
    BPr.append(BP[3:])
    API.append(AP[3:])
    PI.append(PL[2:])
    i += 10
    c += 10
    DI = worksheet.cell_value(i, 5)
    BP = worksheet.cell_value(c, 7)
    AP = worksheet.cell_value(c, 8)
    PL = worksheet.cell_value(c, 9)


years = []

for i in range(len(BPr)):               # Converting all the values to floats to plot
    if BPr[i] != '':
        BPr[i] = float(BPr[i])
    if API[i] != '':
        API[i] = float(API[i])
    if PI[i] != '':
        PI[i] = float(PI[i])

for s in DIs:                                           # Creating dates with format to later parse with numpy
    years.append(s[:4] + '-' + s[4:6] + '-' + s[6:])

for i in range(len(years)):             # Removing empty strings aka where there is no data and assigning 0.0
    if API[i] == '':
        API[i] = 0.0
    if BPr[i] == '':
        BPr[i] = 0.0
    if PI[i] == '':
        PI[i] = 0.0

dates = [pd.to_datetime(d) for d in years]
plt.show()
plt.scatter(dates, API, label="CleanAsk")           # Plotting all the data required
plt.scatter(dates, BPr, label="CleanBid")
plt.scatter(dates, PI, label="Last Price")
plt.xlabel('years')
plt.ylabel('prices')
plt.legend()
plt.show()


