from os import remove, startfile, listdir
from os.path import join
import pandas as pd
from datetime import datetime
from datetime import date as Date
import traceback
from time import sleep
from RoseUtils.rtf_reader import read_rtf

'''def read_rtf(file:str)->list[str]:
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
    return text2
'''
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
        for file in listdir(path):
            if "ORD" in file: odlist.append(join(path,file))
            elif "SEC" in file: elist.append(join(path,file))

        i = 1
        for file in odlist:
            detail = read_rtf(file)
            if self.check_delete: remove(file)
            date = ""
            store = detail[0][83:87]       
            
            for row in detail:

                if len(row) == 30: date = row.strip()
                elif row[12:15] == "DDD": 
                    odf.loc[i] = [store,date,row[7:11],row[87:94]]  # type: ignore
                    i+=1
            sleep(.01)
            self.ui.outputbox.append(f"Running store {store}")
            sleep(.01)

        i = 1
        for file in elist:
            ex = read_rtf(file)
            if self.check_delete: remove(file)
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
        odf.to_csv(join(self.outputfolder,"DDD_Dispatch_Times.csv"),index=False)
        startfile(join(self.outputfolder,"DDD_Dispatch_Times.csv"))
        #pd.pivot_table(data=odf,index=['store'], columns=['week'], values='difference',
            #aggfunc=['mean','median','count']).stack(level=0).unstack(level=0).stack().to_csv(self.outputfolder+"Entire_Report.csv")
        #os.startfile(self.outputfolder+"Entire_Report.csv")

    except:
        self.ui.outputbox.append("ENCOUNTERED ERROR")
        self.ui.outputbox.append("Please send the contents of this box to Justin")
        self.ui.outputbox.append(traceback.format_exc())
