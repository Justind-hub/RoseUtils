from striprtf.striprtf import rtf_to_text
import os
from os.path import exists
import ctypes   
from tkinter import Tk, filedialog
import pandas as pd

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font
from datetime import datetime


def run(zocdownload, outputfolder):
    wb= Workbook()
    ws = wb.create_sheet('Sheet', 0)
    STORECOL = {"1740":1,"1743":2,"2172":3,"2174":4,"2236":5,"2272":6,"2457":7,"2549":8,"2603":9,"2953":10,"3498":11,"4778":12}
    VARIANCEAMOUNT = 10


    #openrtf file
    def openrtf(file): #Call this to open an rtf file with the filepath in the thing.
                        #Will return a list of strings for each line of the file
        with open(file, 'r') as file:
            text = file.read()
        text = rtf_to_text(text)
        textlist = text.splitlines()
        return textlist #returns a list containing entire RTF file

    def left(s, amount): #litterally just the "left" function from excel
        return s[:amount]

    filelist = [] #create list of files to loop through -> filelist
    for file in os.listdir(zocdownload):
        if file.startswith("INVTAR"):
            filelist.append(file)
    f = 0
    for each in filelist:
            
        #clean the file
        textlist = openrtf(zocdownload+each)
        headerstart = left(textlist[1],10)
        i = 10
        while i < len(textlist):
            if left(textlist[i],10) == headerstart:
                del textlist[i-2:i+4]
                i += 4
            else:
                i += 1
        while("" in textlist) :
            textlist.remove("")
        del textlist[-22:]

        i = 3
        while i < len(textlist):
            try:
                x = int(left(textlist[i], 7))
            except:
                del textlist[i]
                i-=1
            i +=1

        #Date and store number
        store = textlist[0]
        store = store[82:90].strip()
        store = store[-4:]
        #store = int(store)
        date = textlist[2]
        date = date.strip()
        del textlist[0:4]
        col = STORECOL[store]
        #create 3 lists of item#, item name and variance

        item = []
        dvariance = []
        for i, line in enumerate(textlist):
            item.append(line[11:34].strip())
            dvariance.append(float(line[105:115].strip()))
        df = pd.DataFrame(data={'Item':item,'$ Variance':dvariance})
        #df.sort_values('$ Variance', inplace=True)
        df.drop(df[(df['$ Variance'] > -VARIANCEAMOUNT) & (df['$ Variance'] < VARIANCEAMOUNT)].index, inplace = True)
        df['Variance'] = df['Item'] + " " +df['$ Variance'].map(str)
        df.sort_values(by='$ Variance', key=abs, inplace=True, ascending=False)
        df.reset_index(inplace=True)
        df.drop(['Item','$ Variance','index'], axis=1, inplace=True)
        ws.cell(row = 2, column = col, value = int(store))
        
        #Write variances to cells
        i = 0
        variances = df['Variance'].to_list()
        while i < len(variances):
            ws.cell(column = col, row = i+3, value = variances[i])
            i+=1
        

        f +=1





    for col in ws.columns:
        column = col[0].column_letter # Get the column name
        maxl = 8
        for cell in col:
            x = str(cell.internal_value)
            if len(x) > maxl:
                maxl = len(x)
            cell.alignment = Alignment(horizontal="center")
        ws.column_dimensions[column].width = maxl/1.09
        if maxl == 8:
            ws.cell(row = 3, column = col[0].column, value = "No Variances over $"+str(VARIANCEAMOUNT)).font = Font(bold=True)
            ws.column_dimensions[column].width = 20

    ws.cell(row = 1, column = 1, value = date).alignment = Alignment(horizontal="center")
    ws.merge_cells("A1:L1")

    wb.save(outputfolder+"\\Inventory Target "+datetime.now().strftime("%m.%d.%y")+".xlsx")

if __name__ == '__main__':
    run()