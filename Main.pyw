from datetime import datetime
start = datetime.now()

# Globals
import os
from os.path import exists
from subprocess import Popen, PIPE
import logging
import winsound
import time
import threading

#PyQt
from PyQt5 import QtWidgets as qtw
from PyQt5 import QtCore as qtc
from PyQt5 import QtGui
from PyQt5.QtWidgets import  QMessageBox, QFileDialog # Change to * if you get an error



# RoseUtils package functions
from RoseUtils.Main_UI import Ui_RoseUtils
from RoseUtils import Daily_DOR_Breaks, New_Hire, Target_Inventory
from RoseUtils import Weekly_DOR_CSC, Weeklycompfull, Daily_Drivosity
from RoseUtils import Epp, Comments, Release, gm_Target_inv, gm_weeklycomp
from RoseUtils import export_SQL,gm_on_hands, updater, costreport
from RoseUtils import Schedule_reviewer, DDD_Dispatch_Times
from RoseUtils.Buttons import Buttons #Buttons class init as "b"

#Downloaded
import PyPDF2
from PyPDF2 import PdfFileReader, PdfFileWriter

log = logging.getLogger("Justin")
logging.basicConfig(level=logging.INFO, 
                    format='%(asctime)s %(levelname)s:%(name)s:%(message)s',
                    filename='RoseUtils.log')
log.debug("Finished imports")


class RoseUtils(qtw.QMainWindow):
    set_text = qtc.pyqtSignal(str)
    append_text = qtc.pyqtSignal(str)
    def __init__(self, *args, **kwargs):                                       # __init__ - calls all other methods
        super().__init__(*args, **kwargs)
        self.b = Buttons()
        self.ui = Ui_RoseUtils()
        self.ui.setupUi(self)      
        self.timer = True
        self.initui()
        self.initvars()
        self.refreshfolders()
        self.show()
        self.buttons()
        self.menubars()
        self.signals()
        try:
            if not updater.run():
                self.update()
        except Exception:
            log.error("UNABLE TO RUN UPDATER")
        log.debug("Finished __init__")
        self.third = datetime.now() - start
        self.ui.outputbox.append(f"Finished start in {datetime.now() - start}")
        
    def menubars(self):                                                        # links the menubar buttons
        # Actions
        self.ui.actionBETA_V1_2.triggered.connect(lambda: Release.r12(self, True)) 
        self.ui.actionBETA_V1_1.triggered.connect(lambda: Release.r11(self, True)) 
        self.ui.actionBETA_V1_25.triggered.connect(lambda: Release.r125(self, True)) 
        self.ui.actionBETA_V1_3.triggered.connect(lambda: Release.r13(self, True)) 
        self.ui.actionBETA_V1_3_5.triggered.connect(lambda: Release.r135(self, True)) 
        self.ui.actionVersion_1_4.triggered.connect(lambda: Release.r14(self, True)) 
        self.ui.actionVersion_1_4_1.triggered.connect(lambda: Release.r14_1(self, True)) 
        self.ui.actionVersion_1_4_2.triggered.connect(lambda: Release.r14_2(self, True)) 
        self.ui.actionVersion_1_5.triggered.connect(lambda: Release.r15(self, True)) 
        self.ui.actionVersion_1_6.triggered.connect(lambda: Release.r16(self, True)) 
        self.ui.actionVersion_1_7.triggered.connect(lambda: Release.r17(self, True)) 
        self.ui.actionVersion_1_9.triggered.connect(lambda: Release.r19(self, True)) 
        self.ui.actionVersion_1_10.triggered.connect(lambda: Release.r110(self, True)) 
        self.ui.actionVersion_1_13.triggered.connect(lambda: Release.r113(self, True)) 
        self.ui.actionVersion_2_0.triggered.connect(lambda: Release.r20(self, True)) 
        
    def update(self):                                                          # Checks to see if update is available and downloads it
        log.debug("Update function called")
        process = Popen(['git', 'pull', str('https://github.com/Justind-hub/RoseUtils')],stdout=PIPE, stderr=PIPE)
        stdout, stderr = process.communicate()
        log.debug(f"stdout = {stdout}")
        log.debug(f"stderr = {stderr}")
        if "Already up to date" not in str(stdout):
            self.popup("Update Downloaded!\nPlease re-open the program.",QMessageBox.Information,"New Update Downloaded!")# type: ignore
            self.close()
        log.debug("Update function ran")

    @qtc.pyqtSlot(str)
    def set_Text(self, value):
        self.ui.outputbox.setText(value)

    @qtc.pyqtSlot(str)
    def append_Text(self, value):
        self.ui.outputbox.append(value)

    def signals(self):
        self.set_text.connect(self.set_Text) # type: ignore
        self.append_text.connect(self.append_Text) # type: ignore
        

    def initvars(self):                                                        # Class Attributes and program settings initialized
        log.debug("initvars function called")
        
        self.check_delete = False
        self.rcp = True
        filelist = []
        pdflist = []
        cost_report_list = []
        self.gm_filename_checked = False
        if not exists("Settings\\"):
            os.mkdir("Settings")
            self.popup("Please set your folders and files on the Configuration tab",
                        QMessageBox.Information,"First time set-up")# type: ignore
            self.ui.tabWidget.setCurrentIndex(1)
        if exists("Settings\\GM"): 
            self.ui.tabWidget.setTabVisible(0, False)
            self.ui.hider.show()
            self.ui.pwd_submit.show()
            self.ui.pwd_box.show()
            self.ui.pwd_lbl.show()
            self.ui.checkbox_GM.setChecked(True)
            self.ui.checkbox_GM.setEnabled(False)
            self.ui.tabWidget.setCurrentIndex(2)
        else:
            self.ui.tabWidget.setTabVisible(0, True)
            self.ui.hider.hide()
            self.ui.pwd_submit.hide()
            self.ui.pwd_box.hide()
            self.ui.pwd_lbl.hide()
        if exists("Settings\\CCDDatabase"): 
            with open("Settings\\CCDDatabase", "r") as f: self.ccddatabase = f.readline()
        else:
            self.ccddatabase = ""
        if exists("Settings\\export"): 
            with open("Settings\\export", "r") as f: self.outputfolder = f.readline()
        else:
            self.outputfolder = ""
        if exists("Settings\\RCPDatabase"): 
            with open("Settings\\RCPDatabase", "r") as f: self.rcpdatabase = f.readline()
        else:
            self.rcpdatabase = ""
        if exists("Settings\\Zocdownload"): 
            with open("Settings\\Zocdownload", "r") as f: self.zocdownloadfolder = f.readline()
        else:
            self.zocdownloadfolder = ""
        if exists("Settings\\Downloadfolder"): 
            with open("Settings\\Downloadfolder", "r") as f: self.downloadfolder = f.readline()
        else:
            self.downloadfolder = ""
        log.debug("initvars function ran")
        
    def wizzard(self):                                                         # Initializes settings
        log.debug("wizzard function called")
        if not exists("Settings\\CCDDatabase"): self.select_setting(self.ccddatabase, "CCD Database File",True)
        if not exists("Settings\\export"): self.select_setting(self.outputfolder, "Output Folder",False)
        if not exists("Settings\\RCPDatabase"): self.select_setting(self.rcpdatabase, "RCP Database File",True)
        if not exists("Settings\\Zocdownload"): self.select_setting(self.zocdownloadfolder, "Zoc Download Folder",False)
        if not exists("Settings\\Downloadfolder"): self.select_setting(self.downloadfolder, "Downloads Folder",False)
        self.savefolders()
        self.refreshfolders()
        log.debug("wizzard function ran")

    def initui(self):                                                          # Various UI things not editable in Designer. Buttons disbaled by default
        log.debug("initui function called")
        #self.logoframe.pixmap.scaledToWidth(20)
        self.configtab = 1
        disabled = [self.ui.btn_gm_history, self.ui.btn_gm_yields, self.ui.btn_gm_onhand, self.ui.btn_pdf_remove, 
                    self.ui.btn_pdf_clear, self.ui.btn_pdf_up, self.ui.btn_pdf_down, self.ui.btn_pdf_split, 
                    self.ui.btn_pdf_combine]
        
        for button in disabled:
            button.setDisabled(True)

        self.ui.tabWidget.setCurrentIndex(0)   #Set the RCP tab to be default
        self.ui.outputbox.setText("Welcome to RoseUtils!\nClick on a report to get started")

        self.ui.gm_history_label.show()
        self.ui.gm_yields_label.show()
        self.ui.gm_onhand_label.show()

        self.ui.pwd_box.setEchoMode(qtw.QLineEdit.Password)# type: ignore
        self.setWindowIcon(QtGui.QIcon("lib/Logo.ico"))

        log.debug("initui function ran")

    def buttons(self):                                                         # Links all of the UI controls
        log.debug("buttons function called")

        #Buttons
        self.ui.btn_targetrcp.clicked.connect(lambda: threading.Thread(target=Target_Inventory.run,args=(self,)).start())    
        self.ui.btn_breaksrcp.clicked.connect(lambda: threading.Thread(target=Daily_DOR_Breaks.run,args=(self,)).start()) 
        self.ui.btn_new_hirercp.clicked.connect(lambda: New_Hire.run(self, self.zocdownloadfolder, self.rcpdatabase, self.outputfolder)) 
        self.ui.btn_drivosity.clicked.connect(lambda: Daily_Drivosity.run(self, self.downloadfolder, self.rcpdatabase, self.outputfolder)) 
        self.ui.btn_weekly_comprcp.clicked.connect(lambda: threading.Thread(target=Weeklycompfull.run,args=(self,)).start()) 
        self.ui.btn_weekly_dor.clicked.connect(lambda: Weekly_DOR_CSC.run(self, self.zocdownloadfolder, self.outputfolder)) 
        self.ui.btn_all3rcp.clicked.connect(self.all3rcp) 
        #self.ui.btn_all3ccd.clicked.connect(self.all3ccd)
        #self.ui.btn_weekly_compccd.clicked.connect(lambda: Weeklycompfull.run(self, self.zocdownloadfolder, self.outputfolder, "CCD"))
        #self.ui.btn_targetccd.clicked.connect(lambda: Target_Inventory.run(self, "CCD"))
        #self.ui.btn_breaksccd.clicked.connect(lambda: Daily_DOR_Breaks.run(self, self.zocdownloadfolder, self.ccddatabase, "CCD"))
        self.ui.btn_epp.clicked.connect(lambda: Epp.run(self)) 
        #self.ui.tabWidget.currentChanged.connect(self.tabchange)
        self.ui.btn_output_browse.clicked.connect(lambda: self.select_setting(self.outputfolder,"Output Folder",False))
        self.ui.btn_zocdownload_browse.clicked.connect(lambda: self.select_setting(self.zocdownloadfolder, "Zoc Download Folder",False))
        self.ui.btn_rcpdatabase_browse.clicked.connect(lambda: self.select_setting(self.rcpdatabase, "RCP Database File",True))
        self.ui.btn_ccddatabase_browse.clicked.connect(lambda: self.select_setting(self.ccddatabase, "CCD Database File",True))
        self.ui.btn_downloads_browse.clicked.connect(lambda: self.select_setting(self.downloadfolder, "Downloads Folder",False))
        self.ui.btn_output_set.clicked.connect(self.savefolders)
        self.ui.btn_zocdownload_set.clicked.connect(self.savefolders)
        self.ui.btn_rcpdatabase_set.clicked.connect(self.savefolders)
        self.ui.btn_ccddatabase_set.clicked.connect(self.savefolders)
        self.ui.btn_output_set.clicked.connect(self.savefolders)
        self.ui.btn_downloads_set.clicked.connect(self.savefolders)
        self.ui.btn_PA_Promo.clicked.connect(lambda: self.b.pa_promo(self)) 
        self.ui.btn_wizzard.clicked.connect(self.wizzard)
        self.ui.btn_reexport.clicked.connect(lambda: export_SQL.run(self))
        self.ui.check_delete.stateChanged.connect(self.deletecheck)
        self.ui.check_delete_2.stateChanged.connect(self.deletecheck)
        self.ui.btn_DDD.clicked.connect(lambda: threading.Thread(target=DDD_Dispatch_Times.run,args=(self,)).start()) 
        self.ui.radio_CCD.clicked.connect(lambda: self.setfran(False))
        self.ui.radio_RCP.clicked.connect(lambda: self.setfran(True))
        

        # GM tab buttons
        self.ui.btn_browse.clicked.connect(lambda: self.filepicker())
        self.ui.btn_clear.clicked.connect(self.historyClearButton)
        self.ui.btn_gm_history.clicked.connect(lambda: self.historybutton(self.weeklycompslist))
        self.ui.btn_gm_yields.clicked.connect(lambda: self.targetbutton(self.weeklycompslist))
        self.ui.btn_gm_onhand.clicked.connect(lambda: self.onhandbutton(self.weeklycompslist))
        self.ui.checkbox_GM.stateChanged.connect(self.gmbox)
        self.ui.checkbox_debug.stateChanged.connect(self.debug)
        self.ui.pwd_submit.clicked.connect(self.pwd_submitfunc)
        self.ui.pwd_box.returnPressed.connect(self.pwd_submitfunc)
        self.ui.check_custom_file_name.stateChanged.connect(self.gm_filename)
        self.ui.btn_gm_schedule_review.clicked.connect(self.filepicker_for_schedule_review)
        self.ui.btn_gm_schedule_help.clicked.connect(self.gm_schedule_help)
        log.debug("buttons function ran")

        
        # PDF Buttons
        self.ui.btn_pdf_browse.clicked.connect(self.pdfpicker)
        self.ui.btn_pdf_clear.clicked.connect(self.pdfClearButton)
        self.ui.btn_pdf_up.clicked.connect(self.move_up)
        self.ui.btn_pdf_down.clicked.connect(self.move_down)
        self.ui.btn_pdf_split.clicked.connect(self.pdf_split)
        self.ui.btn_pdf_combine.clicked.connect(self.pdf_combine)
        self.ui.pdf_listbox.currentRowChanged.connect(self.pdfboxchanged)
        self.ui.btn_pdf_remove.clicked.connect(lambda: self.ui.pdf_listbox.takeItem(self.ui.pdf_listbox.currentRow()))

        # timer buttons
        self.ui.btn_timer_start.clicked.connect(self.timer_start)
        self.ui.btn_timer_test.clicked.connect(self.timer_test)

        #Inventory cost tab buttons
        self.ui.btn_browse_3.clicked.connect(self.costreportpicker)
        self.ui.btn_clear_3.clicked.connect(self.costreportclear)
        self.ui.btn_runcostreport.clicked.connect(lambda: costreport.run(self))
    


    def setfran(self, x):
        self.rcp = x
    
    def DDD(self):                                                             # Runs the DDD Time report
        x = threading.Thread(target=DDD_Dispatch_Times.run,args=(self,))
        x.start()
      
    def costreportclear(self):                                                 # Clears the cost report list and listbox
        self.cost_report_list = []
        self.ui.file_listbox_3.clear()

    def costreportpicker(self):                                                # Opens file picker to select weekly comp files
        log.debug("filepicker function called")
        self.cost_report_list = []
        costlist, check = QFileDialog.getOpenFileNames(None, "QFileDialog.getOpenFileName()",
                        self.zocdownloadfolder,"RTF Files (*.rtf)")
        if check:
            if len(costlist) == 2:
                for i, file in enumerate(costlist):
                    self.ui.file_listbox_3.addItem(file)
                    self.cost_report_list.append(file)
            else:
                self.popup("Select exactly 2 reports",
                              QMessageBox.Warning,"Error")# type: ignore
            



        log.debug("filepicker function ran, returning "+str(costlist))

    def timer_test(self):                                                      # Beeps once to test the timer settings
        timer_freq = int(self.ui.timer_freq.text())
        timer_length = float(self.ui.timer_length.text()) * 1000
        
        winsound.Beep(timer_freq,int(timer_length))
        
    def timer_start(self):                                                     # Starts the timer
        self.timerthread = threading.Thread(target=self.timer_start_thread)
        self.timerthread.start()

    def timer_start_thread(self):
        if self.ui.timer_repeat_time.text() == "":
            self.popup("Enter the number of seconds between each beep",
                        QMessageBox.Warning,"Error")  # type: ignore
            return
        if self.ui.timer_repeat_num.text() == "":
            self.popup("Enter the number of times to beep",
                        QMessageBox.Warning,"Error")  # type: ignore
            return
        timer_freq = int(self.ui.timer_freq.text())
        timer_length = float(self.ui.timer_length.text()) * 1000
        timer_num = int(self.ui.timer_repeat_num.text())
        timer_wait = float(self.ui.timer_repeat_time.text())
        for i in range(timer_num):
            if not self.timer:
                return
            winsound.Beep(timer_freq,int(timer_length))
            time.sleep(timer_wait)

    def pdfboxchanged(self, i):                                                # Updates the 'enabled' of the up/down/remove buttons
        if i == 0:
            self.ui.btn_pdf_up.setDisabled(True)
        else:
            self.ui.btn_pdf_up.setDisabled(False)
        if i == self.ui.pdf_listbox.count()-1:
            self.ui.btn_pdf_down.setDisabled(True)
        else:
            self.ui.btn_pdf_down.setDisabled(False)

    def move_up(self):                                                         # moves the selected PDF up
        rowIndex = self.ui.pdf_listbox.currentRow()
        currentItem = self.ui.pdf_listbox.takeItem(rowIndex)
        self.ui.pdf_listbox.insertItem(rowIndex - 1, currentItem)
        self.ui.pdf_listbox.setCurrentRow(rowIndex -1)

    def move_down(self):                                                       # moves the selected PDF down
        rowIndex = self.ui.pdf_listbox.currentRow()
        currentItem = self.ui.pdf_listbox.takeItem(rowIndex)
        self.ui.pdf_listbox.insertItem(rowIndex + 1, currentItem)
        self.ui.pdf_listbox.setCurrentRow(rowIndex + 1)

    def pdf_split(self):                                                       # splits the selected PDF(s) into multiple files
        pdflist = [self.ui.pdf_listbox.item(i).text() for i in range(self.ui.pdf_listbox.count())]
        self.ui.outputbox.setText("Splitting your files now..")
        for file in pdflist:
            pdf = PdfFileReader(file)
            for page in range(pdf.getNumPages()):
                pdf_writer = PdfFileWriter()
                pdf_writer.addPage(pdf.getPage(page))

                with open(f"{self.outputfolder}{self.ui.pdf_name.text()}_{page+1}.pdf", 'wb') as out:
                    pdf_writer.write(out)
                self.ui.outputbox.append(f"Saved Page {page} of {file} as {self.outputfolder}{self.ui.pdf_name.text()}_{page+1}.pdf")
        
    def pdf_combine(self):                                                     # Merges PDFs
        pdflist = [self.ui.pdf_listbox.item(i).text() for i in range(self.ui.pdf_listbox.count())]
        merger = PyPDF2.PdfFileMerger()
        for file in pdflist:
            merger.append(file)
        merger.write(self.outputfolder + self.ui.pdf_name.text() + ".pdf")
        self.ui.outputbox.setText(f"Combined file saved as {self.outputfolder}{self.ui.pdf_name.text()}.pdf")
        self.ui.outputbox.append("Opening file now....")
        os.startfile(self.outputfolder + self.ui.pdf_name.text() + ".pdf")
        
    def debug(self,state):                                                     # Enables or disables debug mode
        log.debug("debug function called")
        if state == qtc.Qt.Checked: # type: ignore
            log.setLevel(logging.DEBUG)
            log.debug("debug mode turned on")
        if state != qtc.Qt.Checked: # type: ignore
            log.debug("debug mode turned off")
            log.setLevel(logging.INFO)
        log.debug("Debug mode turned on")

    def pwd_submitfunc(self):                                                  # Checks the submitted password
        log.debug(f"Password submitted: {str(self.ui.pwd_box.text())}")
        pwd = str(self.ui.pwd_box.text())
        self.ui.pwd_box.setText("")
        if pwd == "24bo":
            self.ui.checkbox_GM.setChecked(False)
            self.ui.checkbox_GM.setEnabled(True)
 
    def gm_filename(self, state):                                              # runs the logic for the custom filename checkbox
        if state == qtc.Qt.Checked:# type: ignore
            self.ui.gm_filename.setEnabled(True)
            self.gm_filename_checked = True
        else:
            self.ui.gm_filename.setEnabled(False)
            self.gm_filename_checked = False

    def filepicker_for_schedule_review(self):                                  # Opens file picker to select weekly comp files
        log.debug("filepicker function called")
        filelist , check = QFileDialog.getOpenFileNames(None, 
                                        "QFileDialog.getOpenFileName()",
                        self.zocdownloadfolder,"HTML Files (*.html)")
        if check:
            for i, file in enumerate(filelist):
                self.ui.file_listbox.insertItem(i,file[file.rfind("/")+1:])
            self.weeklycompslist = filelist
            self.ui.numitems_label.setText(str(len(filelist)))
        if len(filelist) == 1:
            Schedule_reviewer.run(self)
            self.ui.btn_gm_yields.setDisabled(False)
            self.ui.btn_gm_onhand.setDisabled(False)
            self.ui.gm_yields_label.hide()
            self.ui.gm_onhand_label.hide()
        else:
            self.ui.btn_gm_yields.setDisabled(True)
            self.ui.btn_gm_onhand.setDisabled(True)
            self.ui.gm_yields_label.show()
            self.ui.gm_onhand_label.show()
        log.debug("filepicker function ran, returning "+str(filelist))

    def gm_schedule_help(self):       ########## STILL NEED TO WRITE ########### Displays help text to the outputbox
        pass

    def deletecheck(self, state):                                              # Logic for the "Delete after report" checkbox
        if state == qtc.Qt.Checked:# type: ignore
            self.check_delete = True
            self.ui.outputbox.setText(" - - - - - - WARNING - - - - - - ")
            self.ui.outputbox.append("  Report files will be deleted after you run the report")
            self.ui.check_delete.setChecked(True)
            self.ui.check_delete_2.setChecked(True)
        else:
            self.check_delete = False
            self.ui.outputbox.setText("Report files will NOT be deleted after you run the report")
            self.ui.check_delete.setChecked(False)
            self.ui.check_delete_2.setChecked(False)

    def gmbox(self, state):                                                    # Logic for the GM Mode checkbox
        log.debug("gmbox function called")
        if state == qtc.Qt.Checked:# type: ignore
            with open("Settings\\GM", "w") as f:
                f.write("Hello")
            self.ui.tabWidget.setTabVisible(0, False)
            self.ui.hider.show()
            self.ui.pwd_submit.show()
            self.ui.pwd_box.show()
            self.ui.pwd_lbl.show()
            self.ui.checkbox_GM.setEnabled(False)
        if state != qtc.Qt.Checked:# type: ignore
            os.remove("Settings\\GM")
            self.ui.hider.hide()
            self.ui.pwd_submit.hide()
            self.ui.pwd_box.hide()
            self.ui.pwd_lbl.hide()
            self.ui.tabWidget.setTabVisible(0, True)
        log.debug("gmbox function ran")

    def filepicker(self):                                                      #Opens file picker to select weekly comp files
        log.debug("filepicker function called")
        filelist , check = QFileDialog.getOpenFileNames(None, "QFileDialog.getOpenFileName()",
                        self.zocdownloadfolder,"RTF Files (*.rtf)")
        if check:
            for i, file in enumerate(filelist):
                self.ui.file_listbox.insertItem(i,file[file.rfind("/")+1:])
            self.weeklycompslist = filelist
            self.ui.numitems_label.setText(str(len(filelist)))
        if len(filelist) == 1:
            self.ui.btn_gm_yields.setDisabled(False)
            self.ui.btn_gm_onhand.setDisabled(False)
            self.ui.gm_yields_label.hide()
            self.ui.gm_onhand_label.hide()
        else:
            self.ui.btn_gm_yields.setDisabled(True)
            self.ui.btn_gm_onhand.setDisabled(True)
            self.ui.gm_yields_label.show()
            self.ui.gm_onhand_label.show()
        if len(filelist) >= 1 and len(filelist) <=4:
            self.ui.btn_gm_history.setDisabled(False)
            self.ui.gm_history_label.hide()
        log.debug("filepicker function ran, returning "+str(filelist))
    
    def pdfpicker(self):                                                       #Opens file picker to select weekly comp files
        log.debug("filepicker function called")
        pdflist , check = QFileDialog.getOpenFileNames(None, "QFileDialog.getOpenFileName()",
                        self.outputfolder,"PDF Files (*.pdf)")
        if check:
            for i, file in enumerate(pdflist):
                self.ui.pdf_listbox.addItem(file)
            self.pdffilelist = pdflist
            self.ui.numitems_label_2.setText(str(len(pdflist)))
            self.ui.btn_pdf_remove.setDisabled(False)
            self.ui.btn_pdf_clear.setDisabled(False)
            self.ui.btn_pdf_split.setDisabled(False)
            self.ui.btn_pdf_combine.setDisabled(False)

        log.debug("filepicker function ran, returning "+str(pdflist))

    def historybutton(self,filelist):                                          # Runs the GM schedule history button
        log.debug("historybutton function called")
        if len(filelist) == 0:
            self.popup("Click on 'Browse' and select at least one weekly comparison file before clicking submit!",QMessageBox.Warning,"Error")# type: ignore
            return
        for file in filelist:
            if "DSHWKC" not in file:
                self.ui.outputbox.setText("Select ONLY weekly comparrison reports and try again.\nYou selected:")
                self.ui.outputbox.append(file)
                self.filelist = []
                self.ui.file_listbox.clear()
                return
        gm_weeklycomp.run(self, filelist)
        log.debug("historybutton function ran")

    def targetbutton(self,filelist):                                           # Runs the GM target yield button
        log.debug("targetbutton function called")
        if len(filelist) != 1:
            self.popup("Click on 'Browse' and select exactly ONE target inentory cost report!",QMessageBox.Warning,"Error")# type: ignore
            return
        for file in filelist:
            if "INVTAR" not in file:
                self.ui.outputbox.setText("Must select a target inventory cost report. Please try again.\nYou selected:")
                self.ui.outputbox.append(file)
                self.filelist = []
                self.ui.file_listbox.clear()
                return
        gm_Target_inv.run(self, filelist[0])
        log.debug("targetbutton function ran")

    def onhandbutton(self,filelist):                                           # Runs the GM On Hands button
        log.debug("targetbutton function called")
        if len(filelist) != 1:
            self.popup("Click on 'Browse' and select exactly ONE target inentory cost report!",QMessageBox.Warning,"Error")# type: ignore
            return
        for file in filelist:
            if "INVVAL" not in file:
                self.ui.outputbox.setText("Must select a target inventory cost report. Please try again.\nYou selected:")
                self.ui.outputbox.append(file)
                self.filelist = []
                self.ui.file_listbox.clear()
                return
        gm_on_hands.run(self, filelist[0])
        log.debug("targetbutton function ran")

    def historyClearButton(self):                                              #clears the list box showing files selected
        log.debug("historyclearbutton function called")
        self.weeklycomplist = []
        self.ui.file_listbox.clear()
        self.ui.numitems_label.setText("0")
        self.ui.btn_gm_history.setDisabled(True)
        self.ui.gm_history_label.show()
        self.ui.btn_gm_yields.setDisabled(True)
        self.ui.gm_yields_label.show()
        self.ui.gm_onhand_label.show()
        log.debug("historyclearbutton function ran")
    
    def pdfClearButton(self):                                                  # clears the list box showing files selected
        log.debug("pdfclearbutton function called")
        self.pdflist = []
        self.ui.pdf_listbox.clear()
        self.ui.numitems_label_2.setText("0")
        log.debug("pdfclearbutton function ran")

    def comments(self):                                                        ############# NO LONGER USED # Runs the Comments Report
        log.debug("comments function called")
        self.commentsfile = qtw.QFileDialog.getOpenFileName(self, 'Select the Comments File')
        self.commentsfile = self.commentsfile[0]
        Comments.run(self, self.outputfolder, self.commentsfile)
        log.debug("comments function ran")

    def refreshfolders(self):                                                  # refreshes the folders displayed on the ui settings tab
        log.debug("refreshfolders function ran")
        self.ui.output_edit.setText(self.outputfolder)
        self.ui.zocdownload_edit.setText(self.zocdownloadfolder)
        self.ui.rcpdatabase_edit.setText(self.rcpdatabase)
        self.ui.ccddatabase_edit.setText(self.ccddatabase)
        self.ui.downloads_edit.setText(self.downloadfolder)        
        log.debug("refreshfolders function ran")
    
    def popup(self, text, icon, title):                                        # Generic popup -> self.popup("_BODY_",QMessageBox.Information,"_TITLE_")
        log.debug(f"popup function called with text {text}, icon {icon} and title {title}")
        msg = QMessageBox()
        msg.setWindowTitle(title)
        msg.setText(text)
        msg.setIcon(icon)
        x = msg.exec_()
        log.debug("popup function ran")

    def all3rcp(self):                                                         # All 3 RCP Daily reports
        log.debug("all3rcp function called")
        self.ui.outputbox.setText("Running Reports...")
        threading.Thread(target=Daily_DOR_Breaks.run,args=(self,)).start()
        threading.Thread(target=Target_Inventory.run,args=(self,)).start()
        threading.Thread(target=Weeklycompfull.run,args=(self,)).start()
        log.debug("all3rcp function ran")

    def savefolders(self):                                                     # Dialog box to delect folder for settings
        log.debug("savefolder function called")
        with open("Settings\\CCDDatabase", "w") as f:
            f.write(self.ccddatabase)

        with open("Settings\\export", "w") as f:
            f.write(self.outputfolder)

        with open("Settings\\RCPDatabase", "w") as f:
            f.write(self.rcpdatabase)

        with open("Settings\\Zocdownload", "w") as f:
            f.write(self.zocdownloadfolder)

        with open("Settings\\Downloadfolder", "w") as f:
            f.write(self.downloadfolder)
        self.ui.outputbox.append("All Folders/Files Saved in configuration")
        log.debug("savefolders function ran")

    def select_setting(self, a:str, b:str, file: bool)-> str:                  # Opens folder picker and sets variable
        if file:
            a = qtw.QFileDialog.getOpenFileName(self, f'Select the {b}') # type: ignore
            if "CCD" in b: self.ccddatabase = a[0]
            elif "RCP" in b: self.rcpdatabase = a[0]
        else:
            a = qtw.QFileDialog.getExistingDirectory(self, f'Select {b}') + "/"
            if "Output Folder" in b: self.outputfolder = a               
            elif "Zoc " in b: self.zocdownloadfolder = a
            elif "Downloads Folder" in b: self.downloadfolder = a
        self.ui.outputbox.append(f"{b} Saved")
        self.refreshfolders()
        return a


'''
TODO:
    - Need to create a wrapper function for logging, logging function name and *args
    - 


'''

if __name__ == '__main__':
    app = qtw.QApplication([])

    with open('lib/style.qss', 'r') as f:
        style = f.read()
    app.setStyleSheet(style)

    widget = RoseUtils()
    widget.show()

    app.exec_()