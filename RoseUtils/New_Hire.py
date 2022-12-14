from os import remove, listdir
from striprtf.striprtf import rtf_to_text #RTF Reader
import sqlite3 #Save to database
from xlsxwriter.workbook import Workbook #to export the database into an excel file
import traceback
from datetime import datetime

def run(self, zocdownload, database, export):
    try:
        if self.rcp:
            database = self.rcpdatabase
        else:
            database = self.ccddatabase
        self.ui.outputbox.setText("Running New Hire Report")
        ####### Change the 3 variables below. Inlude double "\\"s, including 2 at the end of paths
        ZOCDOWNLOAD_FOLDER = zocdownload  
        DATABASE_FILE = database
        EXPORT_EXCEL_FILE = export + "New_Hires "+datetime.now().strftime("%m.%d.%y")+".xlsx"
        ###### End of user setup section
        FILENAME = "ODMNHR"
        

        con = sqlite3.connect(DATABASE_FILE)
        cursor=con.cursor()

        ###### If there isn't a 'breaks' table in the database, create one
        cursor.execute('''
            CREATE TABLE IF NOT EXISTS NewHires(
                id INTEGER PRIMARY KEY,
                date TEXT,
                store TEXT,
                TM TEXT,
                TMID TEXT
            )
        ''')

        def openrtf(file): #Call this to open an rtf file with the filepath in the thing.
                            #Will return a list of strings for each line of the file
            with open(file, 'r', errors="ignore") as f:
                text = f.read()
            text = rtf_to_text(text)
            textlist = text.splitlines()
            if self.check_delete:
                remove(file)
            return textlist #returns a list containing entire RTF file


        def findline(file, text,start): #Returns the number of the first line containing the text supplied
            i = start
            while i < len(file):
                if file[i].find(text) == -1:
                    i+=1
                else:
                    return i
            return 0

        fileList = [file for file in listdir(ZOCDOWNLOAD_FOLDER) if file.startswith(FILENAME)] #create list of files to loop through -> filelist
        for file in fileList:
            file = openrtf(ZOCDOWNLOAD_FOLDER + file)
            if file[6] == 'No records found': continue
            store = file[1][81:85]
            newtms = file[6:-1]
            for tm in newtms:
                tmID = tm[0:8]
                name = tm[10:36].strip()
                date = tm[37:].strip()
                cursor.execute("INSERT INTO NewHires(date,store,TM,TMID) VALUES(?,?,?,?)", (date,store,name,tmID))
            
                
        con.commit()
        workbook = Workbook(EXPORT_EXCEL_FILE)
        worksheet = workbook.add_worksheet()
        cursor.execute("SELECT DISTINCT store, tm, date FROM NewHires ORDER BY store")
        mysel=cursor.execute("SELECT DISTINCT store, tm, date FROM NewHires ORDER BY date")
        for i, row in enumerate(mysel):
            for j, value in enumerate(row):
                worksheet.write(i, j, row[j])
        workbook.close()        
        con.close()

    except:
        self.ui.outputbox.append("ENCOUNTERED ERROR")
        self.ui.outputbox.append("Please send the contents of this box to Justin")
        self.ui.outputbox.append(traceback.format_exc())
