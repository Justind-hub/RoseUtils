import pandas as pd
import sqlite3
import openpyxl
from xlsxwriter.workbook import Workbook


def run():
    ######### USER CONFIG SECTION ############
    DOWNLOADS_FOLDER = "C:\\Users\\justi\\Downloads\\"
    DATABASE_FILE = "C:\\Users\\justi\\OneDrive - Rose City Pizza\\Shared Documents\\Store_Data_Database\\Store_Data.db"
    OUTPUT_FILE = "C:\\Users\\justi\\OneDrive - Rose City Pizza\\Shared Documents\\Store_Data_Database\\Drivosity.xlsx"
    ######### END USER CONFIG   ##############




    ######Setup for pandas, SQLite and openpyxl
    conn = sqlite3.connect(DATABASE_FILE)
    cursor = conn.cursor()
    df = pd.read_excel(DOWNLOADS_FOLDER+"report.xlsx",usecols="D,F,G,U",skiprows = 1,header=0)
    wb = openpyxl.load_workbook(DOWNLOADS_FOLDER+"report.xlsx")
    ws = wb.active


    ##### Create pandas DataFrame
    df["Name"] = df["First Name"] +" " + df["Last Name"]
    df = df.drop(["First Name","Last Name"], axis=1 )
    df = df.rename(columns={"Store #":"store","DriveScore":"score","Name":"name"})


    ####Read the date from drivosity report and create list for dataframe, 
                                                    #then add to dataframe
    date = ws['a1'].value
    date = date[date.find("-")+1:-1]
    date = date.replace("/",".")
    date = [date] * len(df)
    df["date"] = date
    df = df[['date','store','name','score']]
    df.to_sql('drivosity',conn,if_exists = 'append',index=False)
    conn.commit()

    #####Export database to excel file
    workbook = Workbook(OUTPUT_FILE)
    worksheet = workbook.add_worksheet()
    cursor.execute("select * from drivosity")
    mysel=cursor.execute("select * from drivosity")
    for i, row in enumerate(mysel):
        for j, value in enumerate(row):
            worksheet.write(i, j, row[j])
    workbook.close()

    conn.close()


if __name__ == '__main__':
    run()