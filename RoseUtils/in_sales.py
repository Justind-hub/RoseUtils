from rtf_reader import read_rtf
from os import listdir, remove
from openpyxl import Workbook
import pandas as pd
from time import strptime, strftime
from datetime import datetime

desktop = "c:/users/justi/onedrive/desktop"
download = "c:/zocdownload/"

filelist = [download+file for file in listdir(download) if "ORDDTL" in file]
times, ordertypes, dollars,dates,stores = [], [], [], [], []
#coclose = strptime("09:40 pm","%H:%M%p")
for file in filelist:
    report = read_rtf(file)

    store = report[0][83:92].strip()
    datelinelen = len(report[6])
    report = [line for line in report if "  **  " in line or len(line) == datelinelen]
    
    
    date = ""
    for line in report:
        if len(line) == 30:
            date = line.strip()
            continue
        find = line.find(":") 
        times.append(line[find-2:find+5])
        ordertypes.append(line[:6].strip())
        dollar = line[137:].strip()
        dollar = dollar[:dollar.find(".")+3]
        dollars.append(float(dollar))
        dates.append(date)
        stores.append(store)

df = pd.DataFrame()
df['date'] = dates
df['type'] = ordertypes
df['dollar'] = dollars
df['time'] = times
df['store'] = stores
#df = df[df.time > coclose]


df.to_excel("output.xlsx")