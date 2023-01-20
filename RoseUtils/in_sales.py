from RoseUtils.rtf_reader import read_rtf
from os import listdir, remove, startfile
import pandas as pd
import traceback



def run(self):    
    try:
        self.append_text.emit(f"Running Incremental Evening Sales Report")
        filelist = [self.zocdownloadfolder+file for file in listdir(self.zocdownloadfolder) if "ORDDTL" in file]
        times, ordertypes, dollars,dates,stores = [], [], [], [], []

        for file in filelist:
            report = read_rtf(file)

            store = report[0][83:92].strip()
            self.append_text.emit(f"Opening Store {store}...")
            datelinelen = len(report[6])
            report = [line for line in report if "  **  " in line or len(line) == datelinelen]
            
            
            date = ""
            for line in report:
                if line[5] == "d" or line[5] == "D": 
                    continue
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
                stores.append(int(store))
            if self.check_delete: remove(self.zocdownloadfolder+file)
        times2 = []

        for time in times:
            h = int(time[:2])
            m = int(time[3:5])
            ap = time[-2]
            if h == 12 and ap == "a": h = 0
            elif h != 12 and ap == "p": h += 12
            h = h + (m/60)
            times2.append(h)


        df = pd.DataFrame()
        df['date'] = dates
        df['type'] = ordertypes
        df['dollar'] = dollars
        df['time'] = times2
        df['store'] = stores

        df = df[ (df['time'] >= 21.66) | (df['time'] <= 3) ]

        self.append_text.emit(f"Found {len(times)} orders, filtered to {len(df)}")
        df.to_excel(self.outputfolder+"Evening Sales.xlsx",index=False)
        startfile(self.outputfolder+"Evening Sales.xlsx")
        self.append_text.emit(f"Report complete, opening file now")
    except:
        self.append_text.emit("ENCOUNTERED ERROR")
        self.append_text.emit("Please send the contents of this box to Justin")
        self.append_text.emit(traceback.format_exc())
        pass