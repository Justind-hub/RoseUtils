import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from matplotlib.backends.backend_pdf import PdfPages
import os
import traceback

def run(self):
    try:
        report = self.weeklycompslist[0]
        def save_pdf(pdf, hourstext, days, array,name):                                                         # Saves the plot to a PDF.
            for i in range(6,-1,-1):
                plt.rc('font',size=4)
                plt.bar(hourstext,array[i])
                plt.xticks(rotation=45)
                plt.title(days[i]+" "+name)
                pdf.savefig(bbox_inches='tight')
                plt.close()
            pdf.close()

        def create_days_list(drivers):                                                                          # create the two days lists
            returnlist = [[],[],[],[],[],[],[]]
            for col in drivers.iterrows():
                returnlist[0].append(col[1][2])
                returnlist[1].append(col[1][3])
                returnlist[2].append(col[1][4])
                returnlist[3].append(col[1][5])
                returnlist[4].append(col[1][6])
                returnlist[5].append(col[1][7])
                returnlist[6].append(col[1][8])
            for i in range(7):
                returnlist[i] = [x for x in returnlist[i] if type(x) != float and x != returnlist[i][0]]
            return returnlist

        def extract_shifts(instoredays):                                                                        # change each string shift into a list with in and out times as floats
            for d, day in enumerate(instoredays):
                for s, shift in enumerate(day):
                    if len(shift)>30:shift = shift[:8]+shift[26:]
                    if "Manager" in shift or "Store" in shift or "Driver" in shift:
                        instoredays[d][s]="Delete Me"
                        print(shift+" saved a crash")
                        continue
                    shift = shift.split(" - ")
                    instoredays[d][s] = shift
                    for i, time in enumerate(shift):
                        if "12:00 AM" in time: h = 24.00
                        elif "12:15 AM" in time: h = 24.25
                        elif "12:30 AM" in time: h = 24.50
                        elif "12:45 AM" in time: h = 24.75
                        elif "1:00 AM" in time and "11" not in time: h = 25.0
                        elif "1:15 AM" in time and "11" not in time: h = 25.25
                        elif "1:30 AM" in time and "11" not in time: h = 25.5
                        else:
                            if time[1] == ":": time = "0" + time
                            h = float(time[0:2])
                            if "P" in time: h+=12
                            if time[3:5] == "15": h+=0.25
                            elif time[3:5] == "30": h+=0.5
                            elif time[3:5] == "45": h+=0.75
                        instoredays[d][s][i] = h
            for i in range(7):
                for _ in range(instoredays[i].count("Delete Me")):
                    instoredays[i].remove("Delete Me")
                    
            return instoredays

        def create_np_array_from_lists(times, arr, driverdays):                                                 # counts number of shifts for every 15 minute
            for d, day in enumerate(driverdays):                                                                # increment of time
                for row, time in enumerate(times):
                    for shift in day:
                        if time >= shift[0]:
                            arr[row][d+1]+=1
                        if time >= shift[1]:
                            arr[row][d+1] -=1
            return arr

        times = [h/4 for h in range(9*4,24*4,1)]                                                                # creates a list of times from 9:00am - 12:45am

        hourstext = []
        drivernp = np.zeros((60,8),np.float16)                                                                  # Create numpy arrays to be used later
        instorenp = np.zeros((60,8),np.float16)

        for i, time in enumerate(times):                                                                        # Adds the time column to each NP array. Useless for the charts,
            drivernp[i][0] = time                                                                               # but useful for debugging and only adds milliseconds
            instorenp[i][0] = time                                                                              # Also creates the hourstext array for the X axis labels of each chart
            if time > 12.8: time -= 12
            text = str(int(time))
            text = text + ":"
            text = text + str((time-int(time))*60)
            hourstext.append(text)

        tables = pd.read_html(report)                                                                           # use pandas to return a list of all the tables on the schedule
        if self.check_delete:
            os.remove(report)
        tables.remove(tables[0])                                                                                # cleaning up the list
        tables.pop()

        hourstext = np.array(hourstext, str)                                                                    # convert hours text list to numpy array

        drivers = tables[1]                                                                                     #Seperate out the drivers and instore schedules
        instores = tables[2]

        if len(tables) == 4:                                                                                    # If mgrs are on their own schedule, copy the, to the instore schedule
            mgrs = tables[3]
            instores = pd.concat((instores,mgrs))                               
            del(mgrs)
        del(tables)
        for df in [drivers,instores]:
            df.drop([0,1], axis=1,inplace=True)

        driverdays = create_days_list(drivers)                                                                  # create the two days lists
        instoredays = create_days_list(instores)
        del(drivers, instores)

        instoredays = extract_shifts(instoredays)                                                               # change each string shift into a list with in and out times as floats
        driverdays = extract_shifts(driverdays)              

        drivernp = create_np_array_from_lists(times, drivernp, driverdays)                                      # counts number of shifts for every 15 minute
        instorenp = create_np_array_from_lists(times, instorenp, instoredays)                                   # increment of time

        days = {6: "Monday", 5: "Tuesday", 4:"Wednesday",3:"Thursday",2:"Friday",1:"Saturday",0:"Sunday"}

        inpdf = PdfPages(self.outputfolder+"Instores.pdf")
        drpdf = PdfPages(self.outputfolder+"Drivers.pdf")

        save_pdf(inpdf, hourstext, days, np.rot90(instorenp),"Instores")                                        # Create plots and save to PDFs
        os.startfile(self.outputfolder+"Instores.pdf")
        save_pdf(drpdf, hourstext, days, np.rot90(drivernp),"Drivers")
        os.startfile(self.outputfolder+"Drivers.pdf")

        writer = pd.ExcelWriter(self.outputfolder+'Schedule_By_Hour.xlsx')                                      # Export Excel file
        pd.DataFrame(drivernp).to_excel(writer,sheet_name="Drivers", engine='xlsxwriter',index=False)
        pd.DataFrame(instorenp).to_excel(writer,sheet_name="Instores", engine='xlsxwriter',index=False)
        writer.save()
        
        self.ui.outputbox.append(f"Completed")
    except:
        self.ui.outputbox.append("ENCOUNTERED ERROR")
        self.ui.outputbox.append("Please send the contents of this box to Justin")
        self.ui.outputbox.append(traceback.format_exc())