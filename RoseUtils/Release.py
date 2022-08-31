def r12(gui,x):
    text = '''Release Version 1.2
    Added the Comments Functionality.
        1. Download some compliments from SMG
        2. Click the "Compliments" Button
        3. Select the comments file you downloaded
    Report will output as an html file to your specificied output folder'''
    
    if x:
        gui.outputbox.setText(text)
    else:
        gui.outputbox.append(text)

    r11(gui, False)


def r11(gui,x):
    text = '''Release Version 1.1
    Added the EPP Functionality
        1. Download Yields and Value reports from ZOC
        2. Click the "EPP" Button
    Report will output to "EPP.xlsx" in your output directory'''

    if x:
        gui.outputbox.setText(text)
    else:
        gui.outputbox.append(text)     

    


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
        gui.outputbox.setText(text)
    else:
        gui.outputbox.append(text) 
    
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
        gui.outputbox.setText(text)
    else:
        gui.outputbox.append(text) 

    r125(gui, False)


def r135(gui,x):  
    text = '''Release Version 1.3.5
    Rewrote Schedule History Report to generate entire export file instead of a file to be referenced by another excel file.
    Should be the last update before final version

    '''
    if x:
        gui.outputbox.setText(text)
    else:
        gui.outputbox.append(text) 

    r13(gui, False)


def r14(gui,x):  
    text = '''Release Version 1.4.0
First NON-BETA Release!!
    All functionality has been tested on store PCs 
    and is confirmed to be working.

    '''
    if x:
        gui.outputbox.setText(text)
    else:
        gui.outputbox.append(text) 

    r13(gui, False)

def r14_1(gui,x):  
    text = '''Release Version 1.4.1
    Bugfix - "Browse" button was crashing the program

    '''
    if x:
        gui.outputbox.setText(text)
    else:
        gui.outputbox.append(text) 

    r14(gui, False)

def r14_2(gui,x):  
    text = '''Release Version 1.4.2
    Added Re-Export Database to DM Reports tab and tooltips to buttons

    '''
    if x:
        gui.outputbox.setText(text)
    else:
        gui.outputbox.append(text) 

    r14_1(gui, False)