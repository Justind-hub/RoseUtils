
from os import listdir, remove, startfile
import pandas as pd
import traceback
from PyQt5.QtWidgets import  QMessageBox, QFileDialog




def run(self):    
    try:
        self.append_text.emit(f"Running PA Promo report")

        report = ""
        costlist, check = QFileDialog.getOpenFileNames(None, "QFileDialog.getOpenFileName()",
                        self.downloadfolder,"CSV Files (Courses*.csv)")
        if check:
            if len(costlist) == 1:
                report = costlist[0]
            else:
                self.popup("Select exactly 1 report1",
                              QMessageBox.Warning,"Error")  # type: ignore
                return
        else:
            return
        
        sa = []
        cm = []
        tg = []
        ca = []
        fp = [] 
        gr = []
        ah = []
        dv = []
        hb = []
        cp = []
        orc = []
        hv = []
        df = pd.read_csv(report,skiprows=7)
        fname = df.Firstname.to_list()
        lname = df.Lastname.to_list()
        completes = df['% Completed'].to_list()
        stores = df.Stores.to_list()
        names = [f + " " + l for f, l in zip(fname, lname)]
        for n, c, s in zip(names, completes, stores):
            if c < 100:
                if "1740" in s: sa.append(n)
                if "1743" in s: cm.append(n)
                if "2172" in s: tg.append(n)
                if "2174" in s: ca.append(n)
                if "2236" in s: fp.append(n)
                if "2272" in s: gr.append(n)
                if "2457" in s: ah.append(n)
                if "2549" in s: dv.append(n)
                if "2603" in s: hb.append(n)
                if "2953" in s: cp.append(n)
                if "3498" in s: orc.append(n)
                if "4778" in s: hv.append(n)
            else:
                continue
        df = pd.DataFrame()
        df['1740'] = sa
        df['1743'] = cm
        df.to_excel(self.outputfolder+"pa_report.xlsx")

            












        print(" ")
    except:
        self.append_text.emit("ENCOUNTERED ERROR")
        self.append_text.emit("Please send the contents of this box to Justin")
        self.append_text.emit(traceback.format_exc())



