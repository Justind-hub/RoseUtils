from RoseUtils import Daily_DOR_Breaks, New_Hire, Target_Inventory, Weekly_DOR_CSC, Weeklycompfull, Daily_Drivosity, Epp
import sys
from PyQt5.QtWidgets import QTabWidget, QPushButton, QLabel, QLineEdit, QMenuBar, QMenu, QMainWindow, QApplication, QMessageBox # Change to * if you get an error
from PyQt5 import uic
from PyQt5 import QtWidgets
import os
from os.path import exists



class MyGUI(QMainWindow):
    def __init__(self):
        super(MyGUI, self).__init__()
        uic.loadUi("lib/Main_UI.ui",self)


        self.initvars()
        self.initui()
        self.show()
        self.buttons()
        self.setFixedSize(self.size())
        


    def initvars(self):
        
        if not exists("Settings/"):
            os.mkdir("Settings")
        if not exists("Settings/CCDDatabase"):
            self.ccddatabasefunc()
        else:
            with open("Settings/CCDDatabase", "r") as f: self.ccddatabase = f.readline()

        if not exists("Settings/export"):
            self.outputfolderunc()
        else:
            with open("Settings/export", "r") as f: self.outputfolder = f.readline()
            
        if not exists("Settings/RCPDatabase"):
            self.rcpdatabasefunc()
        else:
            with open("Settings/RCPDatabase", "r") as f: self.rcpdatabase = f.readline()

        if not exists("Settings/Zocdownload"):
            self.zocdownloadfolderunc()
        else:
            with open("Settings/Zocdownload", "r") as f: self.zocdownloadfolder = f.readline()

        if not exists("Settings/Downloadfolder"):
            self.downloadfolderfunc()
        else:
            with open("Settings/Downloadfolder", "r") as f: self.downloadfolder = f.readline()
        
        self.savefolders()
        self.refreshfolders()

        
        



    def initui(self):
        
        #self.logoframe.pixmap.scaledToWidth(20)
        self.tabWidget.setCurrentIndex(0)   #Set the RCP tab to be default
        self.outputbox.setText("Click on a report to get started")

    def buttons(self):
        #connect button clicks
        ######self.run_target.clicked.connect(lambda: self.targetbutton(self.weeklycompslist))
        self.actionOutput_Folder.triggered.connect(self.outputfolderfunc)
        self.actionZocDownload_Folder.triggered.connect(self.zocdownloadfolderfunc)
        self.actionRCP_Database_Folder.triggered.connect(self.rcpdatabasefunc)
        self.actionCCD_Database_File.triggered.connect(self.ccddatabasefunc)
        self.actionDownloads_Folder.triggered.connect(self.downloadfolderfunc)
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


    def refreshfolders(self):
        self.output_edit.setText(self.outputfolder)
        self.zocdownload_edit.setText(self.zocdownloadfolder)
        self.rcpdatabase_edit.setText(self.rcpdatabase)
        self.ccddatabase_edit.setText(self.ccddatabase)
        self.downloads_edit.setText(self.downloadfolder)        
    
    def popup(self, text, icon, title):
        msg = QMessageBox()
        msg.setWindowTitle(title)
        msg.setText(text)
        msg.setIcon(icon)
        x = msg.exec_()

    def tabchange(self, i):
        if i == 2:
            self.outputbox.hide()
        else:
            self.outputbox.show()

    def all3rcp(self):
        self.outputbox.setText("Running Reports...")
        Daily_DOR_Breaks.run(self, self.zocdownloadfolder, self.rcpdatabase, "RCP")
        Target_Inventory.run(self, self.zocdownloadfolder, self.outputfolder, "RCP")
        Weeklycompfull.run(self, self.zocdownloadfolder, self.outputfolder, "RCP")

    def all3ccd(self):
        self.outputbox.setText("Running Reports...")
        Daily_DOR_Breaks.run(self, self.zocdownloadfolder, self.ccddatabase, "CCD")
        Target_Inventory.run(self, self.zocdownloadfolder, self.outputfolder, "CCD")
        Weeklycompfull.run(self, self.zocdownloadfolder, self.outputfolder, "CCD")


    def outputfolderfunc(self):
        self.outputfolder = QtWidgets.QFileDialog.getExistingDirectory(self, 'Select Output Folder') + "/"
        self.refreshfolders()
        return self.outputfolder

    def zocdownloadfolderfunc(self):
        self.zocdownloadfolder = QtWidgets.QFileDialog.getExistingDirectory(self, 'Select Zoc Download Folder') +"/"
        self.refreshfolders()
        return self.zocdownloadfolder

    def rcpdatabasefunc(self):
        self.rcpdatabase = QtWidgets.QFileDialog.getOpenFileName(self, 'Select the RCP Database File')
        self.rcpdatabase = self.rcpdatabase[0]
        self.refreshfolders()
        return self.rcpdatabase

    def ccddatabasefunc(self):
        self.ccddatabase = QtWidgets.QFileDialog.getOpenFileName(self, 'Select the CCD Database File')
        self.ccddatabase = self.ccddatabase[0]
        self.refreshfolders()
        return self.ccddatabase


    def savefolders(self):

        with open("Settings/CCDDatabase", "w") as f:
            f.write(self.ccddatabase)

        with open("Settings/export", "w") as f:
            f.write(self.outputfolder)

        with open("Settings/RCPDatabase", "w") as f:
            f.write(self.rcpdatabase)

        with open("Settings/Zocdownload", "w") as f:
            f.write(self.zocdownloadfolder)

        with open("Settings/Downloadfolder", "w") as f:
            f.write(self.downloadfolder)

    def downloadfolderfunc(self):
        self.downloadfolder = QtWidgets.QFileDialog.getExistingDirectory(self, 'Select Downloads Folder') +"/"
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