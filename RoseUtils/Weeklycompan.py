import openpyxl
from rtf_reader import read_rtf as openrtf
#from RoseUtils.rtf_reader import read_rtf as openrtf
import pandas as pd

from os import listdir
from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, Border, Side, PatternFill
from datetime import datetime
import traceback
from time import sleep, perf_counter

zocdownloadfolder = "c:/zocdownload/"
output = "c:/users/justi/onedrive/desktop/"
mon = []
tue = []
wed = []
thu = []
fri = []
sat = []
sun = []
report = openrtf(zocdownloadfolder+"DSHWKC.rtf")





print()



'''
                              Monday                       Tuesday                       Wednesday                     Thursday         '
                    Del C/O     DV     IN  Prd   Del C/O     DV     IN  Prd   Del C/O     DV     IN  Prd   Del C/O     DV     IN  Prd'
0:00 am - 10:59 am    1   0   0.99   1.00    2     0   0   0.88   0.90    0     1   3   0.93   2.00   23     0   0   0.94   1.00    0'


Monday-Thursday: Rows 7 through 19
Friday-Sunday: Rows 28 through 40

'''