import sqlite3 #Save to database
from xlsxwriter.workbook import Workbook #to export the database into an excel file

def run(self):
    self.outputbox.setText("Databases Exported")
    RCP_BREAKS = self.rcpdatabase[:self.rcpdatabase.rfind("/")]+"/Breaks.xlsx"
    RCP_DRIVOSITY = self.rcpdatabase[:self.rcpdatabase.rfind("/")]+"/Drivosity.xlsx"
    CCD_BREAKS = self.ccddatabase[:self.ccddatabase.rfind("/")]+"/Breaks.xlsx"
    con = sqlite3.connect(self.rcpdatabase)
    cursor = con.cursor()




    workbook = Workbook(RCP_BREAKS)
    worksheet = workbook.add_worksheet()
    cursor.execute("select * from breaks")
    mysel=cursor.execute("select * from breaks")
    for i, row in enumerate(mysel):
        for j, value in enumerate(row):
            worksheet.write(i, j, row[j])
    workbook.close()

    


    workbook = Workbook(RCP_DRIVOSITY)
    worksheet = workbook.add_worksheet()
    cursor.execute("select * from drivosity")
    mysel=cursor.execute("select * from drivosity")
    for i, row in enumerate(mysel):
        for j, value in enumerate(row):
            worksheet.write(i, j, row[j])
    workbook.close()
    con.close()



    con = sqlite3.connect(self.ccddatabase)
    cursor = con.cursor()
    workbook = Workbook(CCD_BREAKS)
    worksheet = workbook.add_worksheet()
    cursor.execute("select * from breaks")
    mysel=cursor.execute("select * from breaks")
    for i, row in enumerate(mysel):
        for j, value in enumerate(row):
            worksheet.write(i, j, row[j])
    workbook.close()









    con.close()