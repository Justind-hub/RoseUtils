from striprtf.striprtf import rtf_to_text
import pandas as pd
from openpyxl import Workbook
import traceback










def run(self, filelist: list) -> None:
    try:


        def openrtf(file): #Call this to open an rtf file with the filepath in the thing.
                            #Will return a list of strings for each line of the file
            with open(file, 'r') as file:
                text = file.read()
            text = rtf_to_text(text)
            textlist = text.splitlines()
            return textlist #returns a list containing entire RTF file

        def getday(text, a): #send text file and starting line for the day of the week, returns dataframe
            drivers = []
            starthour = []
            instores = []
            deliveries = []
            products = []
            for i,line in enumerate(text):
                products.append(int(line[a+22:a+25].strip()))
                deliveries.append(int(line[a:a+2].strip()))
                starthour.append(int(line[:2]))
            return(pd.DataFrame({"Time":starthour,'Deliveries':deliveries,'Products':products}))




        wb= Workbook()
        mon = wb.create_sheet('Monday', index = 0)
        tue = wb.create_sheet('Tuesday', index = 0)
        wed = wb.create_sheet('Wednesday', index = 0)
        thu = wb.create_sheet('Thursday', index = 0)
        fri = wb.create_sheet('Friday', index = 0)
        sat = wb.create_sheet('Saturday', index = 0)
        sun = wb.create_sheet('Sunday', index = 0)
        f = 1
        for file in filelist:
            
            
            #*********************Start loop here**************
            textlist = openrtf(file)


            #set store number and date range, and save  
            store = textlist[1]
            store = store[82:90].strip()
            store = store[-4:]
            date = textlist[3]
            date = date.strip()
            col = f+1

            #clean up the file
            del textlist[50:]
            x = []
            for i, line in enumerate(textlist):
                if line[:5] == "     ":
                    x.append(i)
            x.reverse()

                #Delete useless lines, leaving data starting in row 5
            for y in x:
                del textlist[y]
            del textlist[:4]

            #create text1 and text2 for weekdays and weekend
            for i, x in enumerate(textlist):
                if len(x) == 0:
                    stoppoint = i
                    break

            text1 = textlist[:stoppoint]
            del textlist[:stoppoint+1]
            for i, x in enumerate(textlist):
                if len(x) == 0:
                    stoppoint = i
                    break
            text2 = textlist[:stoppoint]
            del textlist[:]
            #get each day as a df
            monday = getday(text1, 22)
            tuesday = getday(text1, 22+29)
            wednesday = getday(text1, 22+29+29)
            thursday = getday(text1, 22+29+29+29)
            friday = getday(text2, 22)
            saturday = getday(text2, 22+29)
            sunday = getday(text2, 22+29+29)


            for row in range(0,13):
                mon.cell(row = row+2, column = col, value = monday.loc[row]['Deliveries'])
                mon.cell(row = row+2+14, column = col, value = monday.loc[row]['Products'])
                mon.cell(row = 1, column = col, value = f"History{f}")

                tue.cell(row = row+2, column = col, value = tuesday.loc[row]['Deliveries'])
                tue.cell(row = row+2+14, column = col, value = tuesday.loc[row]['Products'])
                tue.cell(row = 1, column = col, value = f"History{f}")
                
                wed.cell(row = row+2, column = col, value = wednesday.loc[row]['Deliveries'])
                wed.cell(row = row+2+14, column = col, value = wednesday.loc[row]['Products'])
                wed.cell(row = 1, column = col, value = f"History{f}")

                thu.cell(row = row+2, column = col, value = thursday.loc[row]['Deliveries'])
                thu.cell(row = row+2+14, column = col, value = thursday.loc[row]['Products'])
                thu.cell(row = 1, column = col, value = f"History{f}")

                fri.cell(row = row+2, column = col, value = friday.loc[row]['Deliveries'])
                fri.cell(row = row+2+14, column = col, value = friday.loc[row]['Products'])
                fri.cell(row = 1, column = col, value = f"History{f}")

                sat.cell(row = row+2, column = col, value = saturday.loc[row]['Deliveries'])
                sat.cell(row = row+2+14, column = col, value = saturday.loc[row]['Products'])
                sat.cell(row = 1, column = col, value = f"History{f}")

                sun.cell(row = row+2, column = col, value = sunday.loc[row]['Deliveries'])
                sun.cell(row = row+2+14, column = col, value = sunday.loc[row]['Products'])
                sun.cell(row = 1, column = col, value = f"History{f}")



            #at end of loop:
            worksheets = [mon,tue,wed,thu,fri,sat,sun]
            for sheet in worksheets:
                sheet.cell(row = 2, column = 1, value = "10-11")
                sheet.cell(row = 3, column = 1, value = "11-12")
                sheet.cell(row = 4, column = 1, value = "12-1")
                sheet.cell(row = 5, column = 1, value = "1-2")
                sheet.cell(row = 6, column = 1, value = "2-3")
                sheet.cell(row = 7, column = 1, value = "3-4")
                sheet.cell(row = 8, column = 1, value = "4-5")
                sheet.cell(row = 9, column = 1, value = "5-6")
                sheet.cell(row = 10, column = 1, value = "6-7")
                sheet.cell(row = 11, column = 1, value = "7-8")
                sheet.cell(row = 12, column = 1, value = "8-9")
                sheet.cell(row = 13, column = 1, value = "9-10")
                sheet.cell(row = 2+14, column = 1, value = "10-11")
                sheet.cell(row = 3+14, column = 1, value = "11-12")
                sheet.cell(row = 4+14, column = 1, value = "12-1")
                sheet.cell(row = 5+14, column = 1, value = "1-2")
                sheet.cell(row = 6+14, column = 1, value = "2-3")
                sheet.cell(row = 7+14, column = 1, value = "3-4")
                sheet.cell(row = 8+14, column = 1, value = "4-5")
                sheet.cell(row = 9+14, column = 1, value = "5-6")
                sheet.cell(row = 10+14, column = 1, value = "6-7")
                sheet.cell(row = 11+14, column = 1, value = "7-8")
                sheet.cell(row = 12+14, column = 1, value = "8-9")
                sheet.cell(row = 13+14, column = 1, value = "9-10")



            f +=1



 

        wb.save(self.outputfolder+"Export.xlsx")
        self.outputbox.setText(f"Weekly Comp export successful, \nfile saved to {self.outputfolder}Export.xlsx")
    except:
        self.outputbox.append("ENCOUNTERED ERROR")
        self.outputbox.append("Please send the contents of this box to Justin")
        self.outputbox.append(traceback.format_exc())