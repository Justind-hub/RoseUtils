import os #To search the folder
from striprtf.striprtf import rtf_to_text #RTF Reader
import sqlite3 #Save to database
from xlsxwriter.workbook import Workbook #to export the database into an excel file

def run():

    ####### Change the 3 variables below. Inlude double "\\"s, including 2 at the end of paths
    ZOCDOWNLOAD_FOLDER = "C:\\ZocDownload\\"   
    DATABASE_FILE = "C:\\Users\\justi\\OneDrive - Rose City Pizza\\Shared Documents\\Store_Data_Database\\Store_Data.db"
    EXPORT_EXCEL_FILE = "C:\\Users\\justi\\OneDrive\\Desktop\\New_Hires.xlsx"
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
        with open(file, 'r', errors="ignore") as file:
            text = file.read()
        text = rtf_to_text(text)
        textlist = text.splitlines()
        return textlist #returns a list containing entire RTF file


    def findline(file, text,start): #Returns the number of the first line containing the text supplied
        i = start
        while i < len(file):
            if file[i].find(text) == -1:
                i+=1
            else:
                return i
        return 0

    fileList = [file for file in os.listdir(ZOCDOWNLOAD_FOLDER) if file.startswith(FILENAME)] #create list of files to loop through -> filelist
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


if __name__ == '__main__':
    run()