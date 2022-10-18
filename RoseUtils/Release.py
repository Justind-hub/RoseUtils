def aboutme(gui) -> None:
    gui.ui.outputbox.setText('''Coming soon!''')

def r12(gui,x):
    text = '''Release Version 1.2
    Added the Comments Functionality.
        1. Download some compliments from SMG
        2. Click the "Compliments" Button
        3. Select the comments file you downloaded
    Report will output as an html file to your specificied output folder'''
    
    if x:
        gui.ui.outputbox.setText(text)
    else:
        gui.ui.outputbox.append(text)

    r11(gui, False)


def r11(gui,x):
    text = '''Release Version 1.1
    Added the EPP Functionality
        1. Download Yields and Value reports from ZOC
        2. Click the "EPP" Button
    Report will output to "EPP.xlsx" in your output directory'''

    if x:
        gui.ui.outputbox.setText(text)
    else:
        gui.ui.outputbox.append(text)     

    


def r125(gui,x):  
    text = '''Release Version 1.25
    Deleted the RCP and CCD tabs and moved all the reports on to one.
    Other misc codebase changes to make it easier to add more in the future


Release Version 1.2
    Added the Comments Functionality.
        1. Download some compliments from SMG
        2. Click the "Compliments" Button
        3. Select the comments file you downloaded
    Report will output as an html file to your specificied output folder

'''

    if x:
        gui.ui.outputbox.setText(text)
    else:
        gui.ui.outputbox.append(text) 
    
    r11(gui, False)



 

def r13(gui,x):  
    text = '''Release Version 1.3.0
    Added current GM Utilities
    Changed name RoseUtils, merging with other program to only be 1 moving forward
    
    Yields Report:
        Download target Yields from ZOC, click the button. 
        Copy and paste yields into ACO
    Schedule History:
        Download up to 4 weekly comparison reports from ZOC, then click the button
        Will export the file needed for the "Scheduler History" spreadsheet

    '''
    if x:
        gui.ui.outputbox.setText(text)
    else:
        gui.ui.outputbox.append(text) 

    r125(gui, False)


def r135(gui,x):  
    text = '''Release Version 1.3.5
    Rewrote Schedule History Report to generate entire export file instead of a file to be referenced by another excel file.
    Should be the last update before final version

    '''
    if x:
        gui.ui.outputbox.setText(text)
    else:
        gui.ui.outputbox.append(text) 

    r13(gui, False)


def r14(gui,x):  
    text = '''Release Version 1.4.0
First NON-BETA Release!!
    All functionality has been tested on store PCs 
    and is confirmed to be working.

    '''
    if x:
        gui.ui.outputbox.setText(text)
    else:
        gui.ui.outputbox.append(text) 

    r13(gui, False)

def r14_1(gui,x):  
    text = '''Release Version 1.4.1
    Bugfix - "Browse" button was crashing the program

    '''
    if x:
        gui.ui.outputbox.setText(text)
    else:
        gui.ui.outputbox.append(text) 

    r14(gui, False)

def r14_2(gui,x):  
    text = '''Release Version 1.4.2
    Added Re-Export Database to DM Reports tab and tooltips to buttons

    '''
    if x:
        gui.ui.outputbox.setText(text)
    else:
        gui.ui.outputbox.append(text) 

    r14_1(gui, False)

def r15(gui,x):  
    text = '''Release Version 1.5.0
    Added sales to schedule history report. 
    Rows 17 and 47 now show each day's sales divided by 100 (4385 shows as '44') as well as the high low and average for each day. Should be helpful with projected sales!

    '''
    if x:
        gui.ui.outputbox.setText(text)
    else:
        gui.ui.outputbox.append(text) 

    r14_2(gui, False)

def r16(gui,x):  
    text = '''Release Version 1.6.... PDF Tools!!

    Add as many PDFs to the list as you like by selecting "Browse", you can rearange them however you like
    Once you're satisfied with the order of the PDFs, enter a name for your new file into the box above the list of PDFs. You can then either "Combine" or "Spit".
      - Split: Each individual page of each PDF will be saved as a new file in your Output Folder, with the page number appended
      - Combine: All of the PDFs selected will be combined into one file in the order shown in the box and saved as as a new file

      

    '''
    if x:
        gui.ui.outputbox.setText(text)
    else:
        gui.ui.outputbox.append(text) 

    r15(gui, False)


def r17(gui,x):  
    text = '''Release Version 1.7... Lots of small things..

    - Added an "On Hand Amounts" report for ACOs to accompany the yields report. Items that are missing are by design, not a bug :)
    - Slightly changed the auto updater to now quit the program and force you to reopen it if an update is downloaded
    - Added some external libraries to the repo to get around some "pip" issues. Now only have to pip install Pandas and PyQT5
    - Some various refactoring for performance increases. Some daily DM reports now run about 20% faster.
    - Started the process of re-writing the StripRTF library
      

    '''
    if x:
        gui.ui.outputbox.setText(text)
    else:
        gui.ui.outputbox.append(text) 

    r16(gui, False)


def r19(gui,x):  
    text = '''Release Version 1.9... Added the timer tab
    I put this together very quickly for my own use
    Enter how many times you want the timer to repeat, feel free to change the frequency or duration, then enter the number of seconds in the between in each beep
    The entire GUI will freeze while the timer is running. If you have to close out, click on it a bunch until it completely freezes and then use windows to kill the process


    '''
    if x:
        gui.ui.outputbox.setText(text)
    else:
        gui.ui.outputbox.append(text) 

    r17(gui, False)

def r110(gui,x):  
    text = '''Release Version 1.10... Commodity Price Change Analysis.

    Download any two weeks Actual Inventory Cost Reports from ZOC
    Select them with the "Browse" button, then click Run Report.
    The file will output to your desktop
    '''
    if x:
        gui.ui.outputbox.setText(text)
    else:
        gui.ui.outputbox.append(text) 

    r19(gui, False)

def r113(gui,x):  
    text = '''Release Version 1.13... File cleanup!
    
    If you check the "Delete files after report" button, the downloaded RTF or CSV files will be automatically deleted when you're done
    '''
    if x:
        gui.ui.outputbox.setText(text)
    else:
        gui.ui.outputbox.append(text) 

    r110(gui, False)

def r20(gui,x):  
    text = '''Release Version 2.0... Loads of small updates and very few release notes!
    Well I skipped 1.4-1.16, but here's the gist of it:
    Added a couple new buttons and corresponding reports, most notably the PA report and the DDD Dispatch Time report
    Made all of the DM reports work for both franchises
    TONS of changes under the hood to make future updates easier, and now the window is resizable!
    '''
    if x:
        gui.ui.outputbox.setText(text)
    else:
        gui.ui.outputbox.append(text) 

    r113(gui, False)