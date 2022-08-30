# -*- coding: utf-8 -*-
from RoseUtils import Daily_DOR_Breaks, New_Hire, Target_Inventory, Weekly_DOR_CSC, Weeklycompfull, Daily_Drivosity, Epp, Comments, Release, gm_Target_inv, gm_weeklycomp
import sys
from PyQt5.QtWidgets import QTabWidget, QPushButton, QLabel, QLineEdit, QMenuBar, QMenu, QMainWindow, QApplication, QMessageBox, QFileDialog, QCheckBox # Change to * if you get an error
from PyQt5 import uic,QtWidgets,QtCore
import os
from os.path import exists
from subprocess import Popen, PIPE
import logging

log = logging.getLogger("Justin")
logging.basicConfig(level=logging.INFO, format='%(asctime)s %(levelname)s:%(name)s:%(message)s',filename='RoseUtils.log')
log.debug("Finished imports")



class MyGUI(QMainWindow):
    def __init__(self):
        super(MyGUI, self).__init__()
        uic.loadUi("lib\\Main_UI.ui",self)
        filelist = []
        self.initui()
        self.initvars()
        self.refreshfolders()
        self.show()
        self.buttons()
        self.setFixedSize(self.size())
        try:
            self.update()
        except Exception:
            log.error("UNABLE TO RUN UPDATER")
        log.debug("Finished __init__")

    def update(self):
        log.debug("Update function called")
        process = Popen(['git', 'pull', str('https://github.com/Justind-hub/RoseUtils')],stdout=PIPE, stderr=PIPE)
        stdout, stderr = process.communicate()
        log.debug(f"stdout = {stdout}")
        log.debug(f"stderr = {stderr}")
        if "Already up to date" not in str(stdout):
            self.popup("Update Downloaded!\nPlease quit the program and re-open to apply",QMessageBox.Information,"New Update Downloaded!")
        log.debug("Update function ran")

    def initvars(self):
        log.debug("initvars function called")
        if not exists("Settings\\"):
            os.mkdir("Settings")
            self.popup("Please set your folders and files on the Configuration tab",QMessageBox.Information,"First time set-up")
            self.tabWidget.setCurrentIndex(1)
        if exists("Settings\\GM"): 
            self.tabWidget.setTabVisible(0, False)
            self.hider.show()
            self.checkbox_GM.setChecked(True)
            self.checkbox_GM.set_enabled
            self.checkbox_GM.setEnabled(False)
            self.tabWidget.setCurrentIndex(2)
        else:
            self.tabWidget.setTabVisible(0, True)
            self.hider.hide()
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
        

    def wizzard(self):
        log.debug("wizzard function called")
        if not exists("Settings\\CCDDatabase"): self.ccddatabasefunc()
        if not exists("Settings\\export"): self.outputfolderfunc()    
        if not exists("Settings\\RCPDatabase"): self.rcpdatabasefunc()
        if not exists("Settings\\Zocdownload"): self.zocdownloadfolderfunc()
        if not exists("Settings\\Downloadfolder"): self.downloadfolderfunc()
        self.savefolders()
        self.refreshfolders()
        log.debug("wizzard function ran")

    def initui(self):
        log.debug("initui function called")
        #self.logoframe.pixmap.scaledToWidth(20)
        self.configtab = 1
        self.tabWidget.setCurrentIndex(0)   #Set the RCP tab to be default
        self.outputbox.setText("Click on a report to get started")
        self.btn_gm_history.setDisabled(True)
        self.gm_history_label.show()
        self.btn_gm_yields.setDisabled(True)
        self.gm_yields_label.show()
        log.debug("initui function ran")

    def buttons(self):
        log.debug("buttons function called")
        # Actions (Menu items)
        self.actionOutput_Folder.triggered.connect(self.outputfolderfunc)
        self.actionZocDownload_Folder.triggered.connect(self.zocdownloadfolderfunc)
        self.actionRCP_Database_Folder.triggered.connect(self.rcpdatabasefunc)
        self.actionCCD_Database_File.triggered.connect(self.ccddatabasefunc)
        self.actionDownloads_Folder.triggered.connect(self.downloadfolderfunc) 
        self.actionBETA_V1_2.triggered.connect(lambda: Release.r12(self, True)) 
        self.actionBETA_V1_1.triggered.connect(lambda: Release.r11(self, True)) 
        self.actionBETA_V1_25.triggered.connect(lambda: Release.r125(self, True)) 
        self.actionBETA_V1_3.triggered.connect(lambda: Release.r13(self, True)) 
        self.actionBETA_V1_3_5.triggered.connect(lambda: Release.r135(self, True)) 
        self.actionVersion_1_4.triggered.connect(lambda: Release.r14(self, True)) 
        #Buttons
        self.btn_targetrcp.clicked.connect(lambda: Target_Inventory.run(self, self.zocdownloadfolder, self.outputfolder, "RCP"))
        self.btn_breaksrcp.clicked.connect(lambda: Daily_DOR_Breaks.run(self, self.zocdownloadfolder, self.rcpdatabase, "RCP"))
        self.btn_new_hirercp.clicked.connect(lambda: New_Hire.run(self, self.zocdownloadfolder, self.rcpdatabase, self.outputfolder))
        self.btn_drivosity.clicked.connect(lambda: Daily_Drivosity.run(self, self.downloadfolder, self.rcpdatabase, self.outputfolder))
        self.btn_weekly_comprcp.clicked.connect(lambda: Weeklycompfull.run(self, self.zocdownloadfolder, self.outputfolder, "RCP"))
        self.btn_weekly_dor.clicked.connect(lambda: Weekly_DOR_CSC.run(self, self.zocdownloadfolder, self.outputfolder))
        self.btn_all3rcp.clicked.connect(self.all3rcp)
        self.btn_all3ccd.clicked.connect(self.all3ccd)
        self.btn_weekly_compccd.clicked.connect(lambda: Weeklycompfull.run(self, self.zocdownloadfolder, self.outputfolder, "CCD"))
        self.btn_targetccd.clicked.connect(lambda: Target_Inventory.run(self, self.zocdownloadfolder, self.outputfolder, "CCD"))
        self.btn_breaksccd.clicked.connect(lambda: Daily_DOR_Breaks.run(self, self.zocdownloadfolder, self.ccddatabase, "CCD"))
        self.btn_epp.clicked.connect(lambda: Epp.run(self, self.zocdownloadfolder, self.outputfolder))
        self.tabWidget.currentChanged.connect(self.tabchange)
        self.btn_output_browse.clicked.connect(self.outputfolderfunc)
        self.btn_zocdownload_browse.clicked.connect(self.zocdownloadfolderfunc)
        self.btn_rcpdatabase_browse.clicked.connect(self.rcpdatabasefunc)
        self.btn_ccddatabase_browse.clicked.connect(self.ccddatabasefunc)
        self.btn_downloads_browse.clicked.connect(self.downloadfolderfunc)
        self.btn_output_set.clicked.connect(self.outputfolderfunc2)
        self.btn_zocdownload_set.clicked.connect(self.savefolders)
        self.btn_rcpdatabase_set.clicked.connect(self.savefolders)
        self.btn_ccddatabase_set.clicked.connect(self.savefolders)
        self.btn_output_set.clicked.connect(self.savefolders)
        self.btn_downloads_set.clicked.connect(self.savefolders)
        self.btn_compliments.clicked.connect(self.comments)
        self.btn_wizzard.clicked.connect(self.wizzard)

        # GM specific buttons
        self.btn_browse.clicked.connect(self.filepicker)
        self.btn_clear.clicked.connect(self.historyClearButton)
        self.btn_gm_history.clicked.connect(lambda: self.historybutton(self.weeklycompslist))
        self.btn_gm_yields.clicked.connect(lambda: self.targetbutton(self.weeklycompslist))
        self.checkbox_GM.stateChanged.connect(self.gmbox)
        log.debug("buttons function ran")

    def gmbox(self, state):
        log.debug("gmbox function called")
        if state == QtCore.Qt.Checked:
            with open("Settings\\GM", "w") as f:
                f.write("Hello")
            self.tabWidget.setTabVisible(0, False)
            self.hider.show()
        if state != QtCore.Qt.Checked:
            os.remove("Settings\\GM")
            self.hider.hide()
            self.tabWidget.setTabVisible(0, True)
        log.debug("gmbox function ran")

    def filepicker(self): #Opens file picker to select weekly comp files
        log.debug("filepicker function called")
        filelist , check = QFileDialog.getOpenFileNames(None, "QFileDialog.getOpenFileName()",
                        self.zocdownloadfolder,"RTF Files (*.rtf)")
        if check:
            for i, file in enumerate(filelist):
                self.file_listbox.insertItem(i,file[file.rfind("/")+1:])
            self.weeklycompslist = filelist
            self.numitems_label.setText(str(len(filelist)))
        if len(filelist) == 1:
            self.btn_gm_yields.setDisabled(False)
            self.gm_yields_label.hide()
        else:
            self.btn_gm_yields.setDisabled(True)
            self.gm_yields_label.show()
        if len(filelist) >= 1 and len(filelist) <=4:
            self.btn_gm_history.setDisabled(False)
            self.gm_history_label.hide()
        log.debug("filepicker function ran, returning "+self.filelist)

    def historybutton(self,filelist):
        log.debug("historybutton function called with filelist "+self.filelist)
        if len(filelist) == 0:
            self.popup("Click on 'Browse' and select at least one weekly comparison file before clicking submit!",QMessageBox.Warning,"Error")
            return
        for file in filelist:
            if "DSHWKC" not in file:
                self.outputbox.setText("Select ONLY weekly comparrison reports and try again.\nYou selected:")
                self.outputbox.append(file)
                self.filelist = []
                self.file_listbox.clear()
                return
        gm_weeklycomp.run(self, filelist)
        log.debug("historybutton function ran")
    def targetbutton(self,filelist):
        log.debug("targetbutton function called with filelist "+self.filelist)
        if len(filelist) != 1:
            self.popup("Click on 'Browse' and select exactly ONE target inentory cost report!",QMessageBox.Warning,"Error")
            return
        for file in filelist:
            if "INVTAR" not in file:
                self.outputbox.setText("Must select a target inventory cost report. Please try again.\nYou selected:")
                self.outputbox.append(file)
                self.filelist = []
                self.file_listbox.clear()
                return
        gm_Target_inv.run(self, filelist[0])
        log.debug("targetbutton function ran")
    def historyClearButton(self): #clears the list box showing files selected
        log.debug("historyclearbutton function called")
        self.weeklycomplist = []
        self.file_listbox.clear()
        self.numitems_label.setText("0")
        self.btn_gm_history.setDisabled(True)
        self.gm_history_label.show()
        self.btn_gm_yields.setDisabled(True)
        self.gm_yields_label.show()
        log.debug("historyclearbutton function ran")
    
    def comments(self):
        log.debug("comments function called")
        self.commentsfile = QtWidgets.QFileDialog.getOpenFileName(self, 'Select the Comments File')
        self.commentsfile = self.commentsfile[0]
        Comments.run(self, self.outputfolder, self.commentsfile)
        log.debug("comments function ran")

    def refreshfolders(self):
        log.debug("refreshfolders function ran")
        self.output_edit.setText(self.outputfolder)
        self.zocdownload_edit.setText(self.zocdownloadfolder)
        self.rcpdatabase_edit.setText(self.rcpdatabase)
        self.ccddatabase_edit.setText(self.ccddatabase)
        self.downloads_edit.setText(self.downloadfolder)        
        log.debug("refreshfolders function ran")
    
    def popup(self, text, icon, title):
        log.debug(f"popup function called with text {text}, icon {icon} and title {title}")
        msg = QMessageBox()
        msg.setWindowTitle(title)
        msg.setText(text)
        msg.setIcon(icon)
        x = msg.exec_()
        log.debug("popup function ran")

    def tabchange(self, i):
        log.debug("tabchange function called")
        if i == self.configtab:
            self.outputbox.hide()
        else:
            self.outputbox.show()
        log.debug("tabchange function ran")

    def all3rcp(self):
        log.debug("all3rcp function called")
        self.outputbox.setText("Running Reports...")
        Daily_DOR_Breaks.run(self, self.zocdownloadfolder, self.rcpdatabase, "RCP")
        Target_Inventory.run(self, self.zocdownloadfolder, self.outputfolder, "RCP")
        Weeklycompfull.run(self, self.zocdownloadfolder, self.outputfolder, "RCP")
        log.debug("all3rcp function ran")

    def all3ccd(self):
        self.outputbox.setText("Running Reports...")
        Daily_DOR_Breaks.run(self, self.zocdownloadfolder, self.ccddatabase, "CCD")
        Target_Inventory.run(self, self.zocdownloadfolder, self.outputfolder, "CCD")
        Weeklycompfull.run(self, self.zocdownloadfolder, self.outputfolder, "CCD")


    def outputfolderfunc(self):
        self.outputfolder = QtWidgets.QFileDialog.getExistingDirectory(self, 'Select Output Folder') + "/"
        self.refreshfolders()
        self.outputbox.append("Output Folder Saved")
        return self.outputfolder

    def zocdownloadfolderfunc(self):
        self.zocdownloadfolder = QtWidgets.QFileDialog.getExistingDirectory(self, 'Select Zoc Download Folder') +"/"
        self.refreshfolders()
        self.outputbox.append("ZocDownload Folder Saved")
        return self.zocdownloadfolder

    def rcpdatabasefunc(self):
        self.rcpdatabase = QtWidgets.QFileDialog.getOpenFileName(self, 'Select the RCP Database File')
        self.rcpdatabase = self.rcpdatabase[0]
        self.outputbox.append("RCP Database File Saved")
        self.refreshfolders()
        return self.rcpdatabase

    def ccddatabasefunc(self):
        self.ccddatabase = QtWidgets.QFileDialog.getOpenFileName(self, 'Select the CCD Database File')
        self.ccddatabase = self.ccddatabase[0]
        self.outputbox.append("CCD Database File Saved")
        self.refreshfolders()
        return self.ccddatabase


    def savefolders(self):
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
        self.outputbox.append("All Folders/Files Saved in configuration")
        log.debug("savefolders function ran")

    def downloadfolderfunc(self):
        self.downloadfolder = QtWidgets.QFileDialog.getExistingDirectory(self, 'Select Downloads Folder') +"/"
        self.outputbox.append("Downloads Folder Saved")
        self.refreshfolders()

    def outputfolderfunc2(self):
        self.outputfolder = self.output_edit.text()
        self.refreshfolders()

    def zocdownloadfolderfunc2(self):
        self.zocdownloadfolder = self.zocdownload_edit.text()
        self.refreshfolders()

    def rcpdatabasefunc2(self):
        self.rcpdatabase = self.rcpdatabase_edit.text()
        self.refreshfolders()

    def ccddatabasefunc2(self):
        self.ccddatabase = self.ccddatabase_edit.text()
        self.refreshfolders()

    def downloadfolderfunc2(self):
        self.downloadfolder = self.downloads_edit.text()
        self.refreshfolders()







def main ():
    app = QApplication(sys.argv)
    window = MyGUI()
    app.exec_()
    
if __name__ == '__main__':
    main()