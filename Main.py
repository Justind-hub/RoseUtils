from inspect import Attribute
from operator import attrgetter
from RoseUtils import Daily_DOR_Breaks, New_Hire, Target_Inventory, Weekly_DOR_CSC, Weeklycompfull, Daily_Drivosity

import sys
from PyQt5.QtWidgets import *
from PyQt5 import uic, QtGui
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
        
        if not exists("settings/"):
            os.mkdir("Settings")
        if not exists("Settings/CCDDatabase"):
            with open("Settings/CCDDatabase", "w") as f:
                f.write(self.ccddatabasefunc())
        else:
            with open("Settings/CCDDatabase", "r") as f: self.ccddatabase = f.readline()

        if not exists("Settings/export"):
            with open("Settings/export", "w") as f:
                f.write(self.outputfolderfunc())
        else:
            with open("Settings/export", "r") as f: self.outputfolder = f.readline()
            
        if not exists("Settings/RCPDatabase"):
            with open("Settings/RCPDatabase", "w") as f:
                f.write(self.rcpdatabasefunc())
        else:
            with open("Settings/RCPDatabase", "r") as f: self.rcpdatabase = f.readline()

        if not exists("Settings/Zocdownload"):
            with open("Settings/Zocdownload", "w") as f:
                f.write(self.zocdownloadfolderfunc())
        else:
            with open("Settings/Zocdownload", "r") as f: self.zocdownloadfolder = f.readline()

        if not exists("Settings/Downloadfolder"):
            with open("Settings/Downloadfolder", "w") as f:
                f.write(self.downloadfolderfunc())
        else:
            with open("Settings/Downloadfolder", "r") as f: self.downloadfolder = f.readline()
        
        self.outputbox.setText("Click on a report to get started")

        



    def initui(self):
        
        #self.logoframe.pixmap.scaledToWidth(20)
        self.tabWidget.setCurrentIndex(0)   #Set the RCP tab to be default

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
               
    
    def popup(self, text, icon, title):
        msg = QMessageBox()
        msg.setWindowTitle(title)
        msg.setText(text)
        msg.setIcon(icon)
        x = msg.exec_()




    def outputfolderfunc(self):
        self.outputfolder = QtWidgets.QFileDialog.getExistingDirectory(self, 'Select Output Folder') + "/"
        return self.outputfolder

    def zocdownloadfolderfunc(self):
        self.zocdownloadfolder = QtWidgets.QFileDialog.getExistingDirectory(self, 'Select Zoc Download Folder') +"/"
        return self.zocdownloadfolder

    def rcpdatabasefunc(self):
        self.rcpdatabase = QtWidgets.QFileDialog.getOpenFileName(self, 'Select the RCP Database File')
        self.rcpdatabase = self.rcpdatabase[0]
        return self.rcpdatabase

    def ccddatabasefunc(self):
        self.ccddatabase = QtWidgets.QFileDialog.getOpenFileName(self, 'Select the CCD Database File')
        self.ccddatabase = self.ccddatabase[0]
        return self.ccddatabase

    def downloadfolderfunc(self):
        self.downloadfolder = QtWidgets.QFileDialog.getExistingDirectory(self, 'Select Downloads Folder') +"/"
        return self.downloadfolder
    
            
def main ():
    app = QApplication(sys.argv)
    window = MyGUI()
    app.exec_()

if __name__ == '__main__':
    main()