import os
import pandas as pd
from datetime import datetime
from datetime import date as Date
import traceback
from time import sleep


def read_rtf(file:str, headers=True)->list[str]:
    with open(file, 'r') as file:
        text = file.readlines()
    for i in range(0,5):
        text.pop(0)
    text2 = []
    for line in text:
        if len(line) != 5:
            if "\\page" in line:
                text2.append("")
            else:
                text2.append(line[66:-2])
    del(text)
    if not headers:
        store = text2[0][83:87]
        header1 = text2[0][:70]
        header2 = text2[1][:70]
        header3 = text2[2][:70]
        i = 0
        while i < len(text2):
            line = text2[i]
            if header1 in line or header2 in line or header3 in line:
                text2.pop(i)
                i-=1
            i+=1 
        return text2, store


    return text2

def run(self):
    try:
        def xlookup(lookup_value, lookup_array, return_array, inf:str = ''):
            match_value = return_array.loc[lookup_array == lookup_value]
            if match_value.empty: return f'"{lookup_value}" not found!' if inf == '' else inf
            else: return match_value.tolist()[0]

        def converttime(str):
            return str[0:5] +" "+ str[-2:]

        def timeobj(str):
            return datetime.strptime(str,"%I:%M %p")

        def formatasminutes(time):
            return int(time.total_seconds()/60)

        def weeknum(d):
            return Date(int(d[-4:]),int(d[:2]),int(d[3:5])).isocalendar()[1] + 1

        path = self.zocdownloadfolder
        odlist = []
        elist = []

        odf = pd.DataFrame(columns=['store','date','order','time_in'])
        edf = pd.DataFrame(columns=['store','date','order','time_dispatch'])
        start = datetime.now()
        for file in os.listdir(path):
            if "ORD" in file: odlist.append(os.path.join(path,file))
            elif "SEC" in file: elist.append(os.path.join(path,file))

        for file in odlist:
            detail = read_rtf(file)
            if self.check_delete: os.remove(file)
            date = ""
            store = detail[0][83:87]       
            i = 1
            for row in detail:

                if len(row) == 30: date = row.strip()
                elif row[12:15] == "DDD": 
                    odf.loc[i] = [store,date,row[7:11],row[87:94]]  # type: ignore
                    i+=1
            sleep(.01)
            self.ui.outputbox.append(f"Running store {store}")
            sleep(.01)

        for file in elist:
            i = 1
            ex = read_rtf(file)
            if self.check_delete: os.remove(file)
            store = ex[0][83:87]

            date = ""
            for row in ex:
                if len(row) == 10: date = row.strip()
                elif len(row) == 121: 
                    edf.loc[i] = [store,date,row[32:36],row[15:23]]  # type: ignore
                    i +=1
        edf['lookup'] = edf['store'] + edf['date'] + edf['order']
        odf['lookup'] = odf['store'] + odf['date'] + odf['order']
        odf['time_dispatch'] = odf['lookup'].apply(xlookup, args = (edf['lookup'],edf['time_dispatch']))
        odf['time_in'] = odf['time_in'].apply(converttime)
        odf['difference'] = odf['time_dispatch'].apply(timeobj) - odf['time_in'].apply(timeobj)
        odf['difference'] = odf['difference'].apply(formatasminutes)
        odf.drop(['lookup'],axis=1,inplace=True)
        odf['week'] = odf['date'].apply(weeknum)

        self.ui.outputbox.append(f"Done! finished in {datetime.now() - start}")
        self.ui.outputbox.append(f"Opening file {self.outputfolder}DDD_Dispatch_Times.csv ")
        odf.to_csv(os.path.join(self.outputfolder,"DDD_Dispatch_Times.csv"),index=False)
        os.startfile(os.path.join(self.outputfolder,"DDD_Dispatch_Times.csv"))
    except:
        self.ui.outputbox.append("ENCOUNTERED ERROR")
        self.ui.outputbox.append("Please send the contents of this box to Justin")
        self.ui.outputbox.append(traceback.format_exc())
'''


d = "08/18/2022"



'''