from RoseUtils.rtf_reader import read_rtf as openrtf
import traceback
from time import sleep
from os import listdir, remove
from openpyxl import Workbook
import pandas as pd
from datetime import datetime



def run(self):
    try:
        zocdownload = self.zocdownloadfolder
        outputfolder = self.outputfolder
        delete = self.check_delete


        self.append_text.emit("Running on hands report...")
        sleep(.01)
        filelist = [zocdownload + file for file in listdir(zocdownload) if "INVVAL" in file]
        itemnumlist = []
        itemnamelist = []
        dataframes = []
        
        for i, file in enumerate(filelist):
            report = openrtf(file)
            if delete: remove(file)
            store = str(int(report[0][83:88].strip()))
            self.append_text.emit(f"Running store {store}")
            sleep(.01)
            file = []
            for line in report:
                try:
                    int(line[2:6])
                    file.append(line)
                except: pass
            if i == 0:
                for line in file:
                    itemnumlist.append(line[2:6])
                    itemnamelist.append(line[9:46].strip())
                itemdict = {'0':itemnamelist}
                itemnamesdf = pd.DataFrame(itemdict,index=itemnumlist)
                dataframes.append(itemnamesdf)
            itemnums = []
            onhands = []
            for line in file:
                itemnums.append(line[2:6])
                onhands.append(float(line[48:54].strip()))
            dictionary = {f"{store}":onhands}
            df = pd.DataFrame(dictionary,index=itemnums)
            dataframes.append(df)
            

        merge = pd. concat(dataframes,axis=1)
        merge = merge.sort_index(axis=1)
        #df.sort_index(axis=1)
        merge.to_excel(outputfolder+"All On Hands "+datetime.now().strftime("%m.%d.%y")+".xlsx")
        self.append_text.emit(f"On Hands Report Complete")
    except:
        self.append_text.emit("ENCOUNTERED ERROR")
        self.append_text.emit("Please send the contents of this box to Justin")
        self.append_text.emit(traceback.format_exc())