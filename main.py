import socket
import tkinter.messagebox
import serial
import sys
import glob
import time
import os
import webbrowser
import multiprocessing
from collections import deque
from queue import Queue
import pandas as pd
import threading
from tkinter import *
from tkinter import filedialog
import tkinter.font as font
import serial.tools.list_ports
import barcode
from barcode.writer import ImageWriter
import win32print
import win32com.client
import keyboard
from PIL import Image
import numpy as np
import io
import ctypes
from bs4 import BeautifulSoup
import wmi
import usb.core
from responsive_voice import ResponsiveVoice


def scanComPort():
    if sys.platform.startswith('win'):
        ports = ['COM%s' % (i + 1) for i in range(256)]
    elif sys.platform.startswith('linux') or sys.platform.startswith('cygwin'):
        # this excludes your current terminal "/dev/tty"
        ports = glob.glob('/dev/tty[A-Za-z]*')
    elif sys.platform.startswith('darwin'):
        ports = glob.glob('/dev/tty.*')
    else:
        raise EnvironmentError('Unsupported platform')

    result = []
    for port in ports:
        try:
            s = serial.Serial(port)
            s.close()
            result.append(port)
        except (OSError, serial.SerialException):
            pass
    return result


def openWifiPort():
    ipAddress = "192.168.86.22"
    port = 8899

    try:
        # with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as s:
        #     s.connect((ipAddress, port))
        s = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
        s.settimeout(0.5)
        s.connect((ipAddress, port))
        weight = s.recv(1024)
        s.close()

    except socket.error as err:
        weight = "wifi not connected"

    return weight


def openComPort(comport):
    try:
        ser2 = serial.Serial()
        ser2.baudrate = 9600
        ser2.port = comport
        ser2.timeout = None  # Use None if you want it to wait forever. 0 doesn't read anything. Third Option is 0.5 seconds
        ser2.open()  # This open the com port
    except (OSError, serial.SerialException, AttributeError, serial.serialutil.SerialException, TypeError) as error:
        ser2 = "FAILED TO OPEN COM PORT"

    return ser2


def readComPort(comport):
    try:
        # if comport.inWaiting() > 0:
        #     print("INSIDE")
        #     ser = comport.readline()
        #ser = comport.readlines()
        ser = comport.read_until()  # Test Read until which reads till \n Use this for Mainstream customer
    except (OSError, serial.SerialException, AttributeError, serial.serialutil.SerialException) as error:
        ser = "FAILED TO READ COM PORT"

    return ser


def updateWeight(weight):
    try:
        weightOutput.insert(END, weight[:-5])  # Use -5 to remove \r\n. Should work on all OP900 & 909.Test other indicators
    except TypeError as error:
        weightOutput.insert(END, weight)  # Use -5 to remove \r\n. Should work on all OP900 & 909.Test other indicators
        #print(error)

    weightOutput.tag_configure("here", justify='center')
    weightOutput.tag_add("here", "1.0", "end")

    return weight


class commandButtons:
    def catchErrors(self, comPort, commandLetter):
        try:
            return comPort.write(commandLetter)
        except (OSError, serial.SerialException, AttributeError, serial.serialutil.SerialException) as error:
            #print(error)
            pass

    def zero(self, comPort):
        return self.catchErrors(comPort, b'Z')
        #return comPort.write(b'Z')

    def units(self, comPort):
        return self.catchErrors(comPort, b'C')
        #return comPort.write(b'C')

    def tare(self, comPort):
        return self.catchErrors(comPort, b'T')
        #return comPort.write(b'T')

    def print(self, comPort):
        return self.catchErrors(comPort, b'P')
        #return comPort.write(b'P')

    def read(self, comPort):
        return self.catchErrors(comPort, b'R')

    def gross(self, comPort):
        return self.catchErrors(comPort, b'G')
        #return comPort.write(b'G')  # Check gross weight in net weighing mode


def openFolder(inputs):
    filename = tkinter.filedialog.askopenfilename(initialdir="/",
                                          title="Select a File",
                                          filetypes=(("Text files",
                                                      "*.xlsx*"),
                                                     ("all files",
                                                      "*.*")))
    # filename= filedialog.askdirectory() # Open Folder
    if filename != "":
        fileEntry.delete(0, END)
        fileEntry.insert(0, filename)
        # Uncommeted Below for C18 = 2 -------------------------------------------------------------------------
        dataFrame = pd.read_excel(filename, na_values="Missing")
        #dataFrame.columns[5]  # List of all columns
        #dataFrame['Box No.'].tolist()



        for i in range(len(dataFrame.columns)):  #Used to go through each column
            values = dataFrame[dataFrame.columns[i]].tolist()
            inputs[i] = deque()
            for j in reversed(range(len(values))):
                inputs[i].appendleft(values[j])


    insertWeightToOutput(inputs)  # Send inputs to screen
    # Uncommeted for C18 = 2 --------------------------------------------------------------------------------


def helpButton():
    helpWindow = Toplevel(window)
    helpWindow.geometry("700x400")
    helpWindow.title("Help selecting com ports and format issues.")
    helpWindow.configure(background="black")
    helpWindow.attributes('-topmost', 'true')

    label = Label(helpWindow, text="1) Plug in RS232 Dongle to PC and Power cable if needed. Power on Scale\n\n"
                                   "2) First select the correct Com Port of Indicator / Readout.\n\n"
                                   "3) Either Select file to save to or create text file in address bar \n\n"
                                   "4) For first time of software use hit the settings button\n\n"
                                   "5) Click printer com port and select com port of printer.\n\n"
                                   "6) If no com ports appear close menu. Replug cables and run again.\n\n"
                                   "7) Then exit menu. Enter in your Inputs. Then hit print. Files save to excel.\n\n",
                  bg="black", fg="white", font=myFont)

    label.config(wraplength="700")
    label.place(x=0, y=0)

    url = "https://youtu.be/MjL-30WFqMM"
    myLink = Label(helpWindow, text="Youtube Video on how to use Software", cursor='hand2',
                   font="none 18 bold", bg="black", fg="white")
    myLink.place(x=100, y=350)
    myLink.bind('<Button-1>', lambda x: webbrowser.open_new(url))


def openTextFileForPrinter():
    try:
        file = open('Printer Com Port.txt', 'r+')  # Opens Printer Com Port text file
        printerPort = file.readline().strip()
    except OSError as e:
        file = open('Printer Com Port.txt', 'w+')  # Creates file if it doesn't exist
        printerPort = "NOTHING"

    return printerPort


def settingsWindow():
    settings = Toplevel(window)
    settings.geometry("400x300")
    settings.resizable(0, 0)
    settings.title("Settings")
    settings.configure(background="black")
    settings.attributes('-topmost', 'true')

    printerPort = openTextFileForPrinter()

    printPort = StringVar()
    textVar = StringVar()
    printComPort = scanComPort()
    printPort.set(printerPort)

    comPortEntry = OptionMenu(settings, printPort, *printComPort)
    comPortEntry.place(x=158, y=100)


    def saveButton():
        file = open('Printer Com Port.txt', 'w+')  # Opens Printer Com Port text file
        file.writelines(printPort.get())
        textVar.set(f'{printPort.get()} SAVED FOR PRINTER')
        file.close()


    saveSettingsButton = Button(settings, text="SAVE", width=8, height=1, command=saveButton, font="none 14 bold")
    saveSettingsButton.place(x=145, y=150)


    saveText = Label(settings, textvariable=textVar, bg="black", fg="white", font=myFont)
    textVar.set(f'{printerPort} SAVED FOR PRINTER')
    saveText.place(x=70, y=220)



def printLog(inputs):  # Used when print button is pressed
    extraSpace = ' '

    saveToPreviousCustomersList(nameBar.get() + extraSpace)
    inputs[0].append(nameBar.get() + extraSpace)
    inputs[1].append(batchNoBar.get() + extraSpace)
    inputs[2].append(packDateBar.get() + extraSpace)
    inputs[3].append(destination4.get() + extraSpace)
    if destination5.get() == "Use Scale Weight":
        #inputs[4].append(weightOutput.get('1.1', 'end-1c'))  # Test with 999999 lbs
        try:
            inputs[4].append(float(weightReading.decode()[1:-8].lstrip()))  # Below line might be a problem with wireless or maybe just press print on indicator to Print
            #inputs[4].append(str(weightReading.decode().split()[1]))
        except AttributeError as error:
            inputs[4].append(weightReading)
    else:
        inputs[4].append(destination5.get() + extraSpace)


    if not inputs[5]:  # Box No that is set to 1 if first run or increment
        inputs[5].append(1)
    else:
        inputs[5].append(int(inputs[5][-1]) + 1)

    df = pd.DataFrame({'Name': inputs[0], 'Batch No.': inputs[1], 'Pack Date': inputs[2], 'Size': inputs[3],
                       'Net Weight (lb)': inputs[4], 'Box No.': inputs[5]})


    dataToExcel = pd.ExcelWriter(fileEntry.get(), engine='xlsxwriter')  # Need to catch error for invalid file type
    df.to_excel(dataToExcel, sheet_name='Sheet1', index=False)


    # Section below is to format excel
    workbook = dataToExcel.book
    worksheet = dataToExcel.sheets['Sheet1']

    column_settings = [{'header': column} for column in df.columns]
    (max_row, max_col) = df.shape

    worksheet.add_table(0, 0, max_row, max_col-1, {'columns': column_settings})

    for i, col in enumerate(df.columns):
        column_len = max(df[col].astype(str).str.len().max(), len(col) + 5)
        worksheet.set_column(i, i, column_len)

    worksheet.insert_image('G2', 'generated_barcode.png', {'x_scale': 0.5, 'y_scale': 0.5})  # Insert barcode image into excel

    # This is a test for sending image to L2 printer
    # image = Image.open("generated_barcode.png")
    # img_byte_arr = io.BytesIO()
    # image.save(img_byte_arr, format='PNG')
    # img_byte_arr = img_byte_arr.getvalue()
    # selectPrintComPort.write(img_byte_arr)


    dataToExcel.close()

    printer = openTextFileForPrinter()
    printerComPort = openComPort(printer)

    insertWeightToOutput(inputs)  # Send inputs to screen
    t412AndL2(inputs, printerComPort)  # Send inputs to printer
    try:
        printerComPort.close()
    except AttributeError as err:
        pass

    return inputs


def insertWeightToOutput(inputs):  # Function to insert weight to Output section
    logOutput.delete('1.0', 'end')
    for j in range(len(inputs[0])):
        for i in range(len(inputs)):
            logOutput.insert(INSERT, str(inputs[i][j]) + '| ')  # Inserts output into logOutput section
        logOutput.insert('end', ' \r\n_________________________________________________________ \r\n')
    logOutput.see('end')  # Scrolls text to the bottom


def t412AndL2(printQue, selectPrintComPort):
    try:
        for i in range(len(printQue)):  # Goes through every input and sends to printer
            #print(printQue[i][-1])
            value = (printQue[i][-1])  # Gets most recent inputs to print
            try:
                if i == 4:  # This adds units of measurements to weight
                    selectPrintComPort.write((str(value) + f' {weightReading.decode().split()[2]}\r\n').encode())
                else:
                    selectPrintComPort.write((str(value) + ' \r\n').encode())
                status.set(" ")
            except AttributeError as error:
                status.set("Printer Not Connected. Check Settings Button.")
                pass

    except (OSError, serial.SerialException) as error:
        #print(error)
        pass



## READ WEIGHT FUNCTION
allPorts = deque()

connectionType = scanComPort()  # View Available Com Ports
#connectionType.append("Internet Option")
command = commandButtons()  # This is used for command request of scales. Only used with C18=3 on scale readout


# Create Main Window
window = Tk()
window.title("New Data Logging Software")
window.geometry("1200x700")
window.minsize(1000, 600)
window.maxsize(1200, 700)
window.configure(background="black")
window.lift()

clickComPort = StringVar()
# if connectionType[0] == "INTERNET OPTION":  # Uncomment for Internet Option
#     clickComPort.set("PLUG IN USB CABLE")
# else:
#     clickComPort.set("SELECT COM PORT")
clickComPort.set("SELECT COM PORT")


myFont = font.Font(family='Helvetica', size=14, weight='bold')  # Font for buttons
myFont2 = font.Font(family='Helvetica', size=52, weight='bold')
myFont3 = font.Font(family='Helvetica', size=18, weight='bold')
weightOutput = Text(window, width=31, height=1, background='white', foreground="#009933", font=myFont2, bd='4')
weightOutput.tag_configure('here', justify='center')
weightOutput.place(x=5, y=0)


comPortDropDown = OptionMenu(window, clickComPort, *connectionType)
comPortDropDown.place(x=5, y=100)
comPortDropDown.config(bg='white', font=myFont)
# Add section for Wi-Fi connect

xAxisCommandButton = 300
yAxisCommandButton = 160
unitsButton = Button(window, text="UNITS", command=lambda: command.units(com2), background='white', width=5)
unitsButton.place(x=xAxisCommandButton, y=yAxisCommandButton)
unitsButton['font'] = myFont

tareButton = Button(window, text="TARE", command=lambda: command.tare(com2), background='white', width=5)
tareButton.place(x=xAxisCommandButton+80, y=yAxisCommandButton)
tareButton['font'] = myFont

zeroButton = Button(window, text="ZERO", command=lambda: command.zero(com2), background='white', width=5)
zeroButton.place(x=xAxisCommandButton+160, y=yAxisCommandButton)
zeroButton['font'] = myFont

printButton = Button(window, text="GROSS", command=lambda: command.gross(com2), background='white', width=5)
printButton.place(x=xAxisCommandButton+240, y=yAxisCommandButton)
printButton['font'] = myFont


# Section for Inputs
def getPreviousCustomersList():
    previousCustomerTextFileLocation = 'Previous Customers.txt'

    try:
        if os.path.exists(previousCustomerTextFileLocation):
            append_write = 'r'  # append if already exists
        else:
            append_write = 'w'  # make a new file if not

        file = open(previousCustomerTextFileLocation, append_write)
        lines = file.readlines()
        file.close()

        for line in lines:
            previousCustomers.append(line.strip())

        return lines

    except (OSError, serial.SerialException, AttributeError, serial.serialutil.SerialException, TypeError) as error:
        pass


def saveToPreviousCustomersList(name='Strand Seafoods LLC'):
    previousCustomerTextFileLocation = 'Previous Customers.txt'

    try:  # Only append customer if it is a new customer
        previous = getPreviousCustomersList()  # Gets list of previous customers
        match = False
        for item in previous:  # Search previous customer for matches
            if name.lower() in item.lower():
                match = True  # If it finds match it will raise flag
                break

        if not match:  # If this is new customer it will be appended
            file = open(previousCustomerTextFileLocation, 'a')
            file.writelines(name + '\n')
            file.close()
            status.set(f'{name} saved to Previous Customer text file')
            previous.append(name)
            updateListBox(previous)

    except (OSError, serial.SerialException, AttributeError, serial.serialutil.SerialException, TypeError) as error:
        #status.set(f'{name} could not be saved.')
        updateListBox(previous)
        file = open(previousCustomerTextFileLocation, 'w')
        file.writelines(name + '\n')
        file.close()
    # Need to make sure it works when filled doesn't exist


def updateListBox(data):  # Update the Listbox
    nameListBox.delete(0, END)

    for item in data:
        nameListBox.insert(END, item)

    nameListBox.see(END)


def filloutListBox(e):  # Update List Box when clicked
    nameBar.delete(0, END)  # Delete whatever is in entry box

    nameBar.insert(0, nameListBox.get(ACTIVE).strip())


def check(e):
    # Grab what is typed
    previousCustomers = getPreviousCustomersList()  # Gets list of previous customers
    typed = nameBar.get()
    if typed == '':  # Update list Box when empty
        data = previousCustomers
    else:
        data = []
        for item in previousCustomers:  # Search previous customer for matches
            if typed.lower() in item.lower():
                data.append(item)

    updateListBox(data)  # Updated List box with selected items


xAxisForInputs = 900
xAxisDelta = 125
yAxisForInputs = 270
status = tkinter.StringVar()
destination1 = tkinter.StringVar()
destination2 = tkinter.StringVar()
destination3 = tkinter.StringVar()
destination4 = tkinter.StringVar()
destination5 = tkinter.StringVar()
destination6 = tkinter.StringVar()
pieces = tkinter.StringVar()

options = ["2.5-5 lb", "5-UP lb", "7-UP lb"]
weightOptions = ["Use Scale Weight", "750", "1500"]
previousCustomers = []
getPreviousCustomersList()
destination4.set(options[2])
destination5.set(weightOptions[0])

statusText = Label(window, textvariable=status, bg="black", fg="white", font=myFont)

inputBarText = Label(window, text="INPUTS", bg="black", fg="white", font=myFont)

nameBarText = Label(window, text="Name:", bg="black", fg="white", font=myFont)

nameBar = Entry(window, textvariable=destination1, width=26, bg="white", font=myFont)

nameListBox = Listbox(window, width=26, height=5, bg="white", font=myFont)  # List box of previous name

batchNoText = Label(window, text="Batch No.:", bg="black", fg="white", font=myFont)

batchNoBar = Entry(window, textvariable=destination2, width=26, bg="white", font=myFont)

packDateText = Label(window, text="Pack Date:", bg="black", fg="white", font=myFont)

packDateBar = Entry(window, textvariable=destination3, width=26, bg="white", font=myFont)

sizeText = Label(window, text="Size:", bg="black", fg="white", font=myFont)

sizeDropDown = OptionMenu(window, destination4, *options)
sizeDropDown.config(bg='white', font=myFont)

netWeightText = Label(window, text="Net Weight:", bg="black", fg="white", font=myFont)

#netWeightBar = Entry(window, textvariable=destination5, width=35, bg="white")
netWeightBar = OptionMenu(window, destination5, *weightOptions)
netWeightBar.config(bg='white', font=myFont)


# If you want no inputs comment below section out ---------------------------------------------------------------------
delta = 120
#nameListBox.bind("<<ListboxSelect>>", filloutListBox)  # Create a binding on the List Box when Clicked
nameListBox.bind("<Button-1>", filloutListBox)  # Create a binding on the List Box when Clicked
nameBar.bind("<KeyRelease>", check)  # Bind when typing into nameBar
statusText.place(x=xAxisForInputs-500, y=yAxisForInputs-50-delta)
inputBarText.place(x=xAxisForInputs+50, y=yAxisForInputs-delta)
nameBarText.place(x=xAxisForInputs-xAxisDelta, y=yAxisForInputs+50-delta)
nameBar.place(x=xAxisForInputs, y=yAxisForInputs+50-delta)
nameBar.insert(END, 'Strand Seafoods LLC')
nameListBox.place(x=xAxisForInputs, y=yAxisForInputs+50-82)
updateListBox(previousCustomers)  #Updates List Box
batchNoText.place(x=xAxisForInputs-xAxisDelta, y=yAxisForInputs+100)
batchNoBar.place(x=xAxisForInputs, y=yAxisForInputs+100)
packDateText.place(x=xAxisForInputs-xAxisDelta, y=yAxisForInputs+150)
packDateBar.place(x=xAxisForInputs, y=yAxisForInputs+150)
sizeText.place(x=xAxisForInputs-xAxisDelta, y=yAxisForInputs+200)
sizeDropDown.place(x=xAxisForInputs, y=yAxisForInputs+200)
netWeightText.place(x=xAxisForInputs-xAxisDelta, y=yAxisForInputs+250)
netWeightBar.place(x=xAxisForInputs, y=yAxisForInputs+250)

# If you want no inputs comment above section out ---------------------------------------------------------------------



boxNoText = Label(window, text="Box No.:", bg="black", fg="white", font=myFont)
#boxNoText.grid(row=6, column=10, columnspan=1000, sticky=NW, padx="300", pady="118")

boxNoBar = Entry(window, textvariable=destination6, width=26, bg="white")
#boxNoBar.grid(row=6, column=10, columnspan=1000, sticky=NW, padx="400", pady="120")


trackingNumberText = Label(window, text="Tracking #:", bg="black", fg="white", font=myFont)
#trackingNumberText.grid(row=6, column=10, columnspan=100, sticky=NW, padx="100", pady="190")

piecesBar = Entry(window, textvariable=pieces, width=26, bg="white", font=myFont)
#piecesBar.grid(row=6, column=10, columnspan=100, sticky=NW, padx="200", pady="260")


# Section for inputs. This will be sent to excel and/or printer

nameBarQue = deque()
batchNoBarQue = deque()
packDateBarQue = deque()
sizeDropDownQue = deque()
netWeightBarQue = deque()
boxNumberQue = deque()
# nameBarQue.append(nameBar.get() + ' \r\n')
# batchNoBarQue.append(batchNoBar.get() + ' \r\n')
# packDateBarQue.append(packDateBar.get() + ' \r\n')
# sizeDropDownQue.append(destination4.get() + ' \r\n')
# netWeightBarQue.append(destination5.get() + ' \r\n')

inputsList = [nameBarQue, batchNoBarQue, packDateBarQue, sizeDropDownQue, netWeightBarQue, boxNumberQue]



# Output Section and button Below that
logOutput = Text(window, width=58, height=13, background='white', font=myFont3)  # Output results go here
logOutput.place(x=0, y=250)


heightForButton = 650
# Button to select file
selectFileButton = Button(window, text="SELECT FILE", width=10, command=lambda: openFolder(inputsList), font=myFont, background='white')
selectFileButton.place(x=5, y=heightForButton)

# Section for entering file destination
fileEntry = Entry(window, width=50, bg="white", font=myFont, background='white')
fileEntry.place(x=138, y=heightForButton+7)
excelFile = "C:" + os.path.join(os.environ["HOMEPATH"], "Desktop") + '\Strand Seafoods LLC.xlsx'  # Target file destination
fileEntry.insert(0, excelFile)

# Help Button should have link to youTube video in the Menu
helpButton = Button(window, text="HELP", width=5, command=helpButton, font=myFont, background='white')
helpButton.place(x=696, y=heightForButton)

# Settings window
settingsButton = Button(window, text="SETTINGS", width=8, command=settingsWindow, font=myFont, background='white')
settingsButton.place(x=768, y=heightForButton)

# Print Button
printButton = Button(window, text="PRINT", width=6, font=myFont, command=lambda: printLog(inputsList), background='white')
printButton.place(x=950, y=heightForButton)



# printer = serial.Serial()
# printer.port = 'COM6'
# printer.baudrate = 9600
# printer.timeout = 0.25
#printer.open()


def dismiss():
    window.grab_release()
    if tkinter.messagebox.askokcancel("Quit", "Do you want to quit"):
        window.destroy()

# window.protocol("WM_DELETE_WINDOW", dismiss)


def updateComPortOrWifi(connectionType, comPortDropDown):
    #if clickComPort.get() == "Internet Option" or clickComPort.get() == "Plug in USB Cable":
    if clickComPort.get():
        comPortDropDown.children["menu"].delete(0, "end")
        connectionType = scanComPort()
        connectionType.append("Internet Option")

        for v in connectionType:
            comPortDropDown.children["menu"].add_command(label=v, command=lambda veh=v: clickComPort.set(veh))

    return connectionType


start = time.time()
beg = time.time()
com2 = ''

# Section below for log + print ----------------------------------------------------------------------------------------
wasPrintButtonPressedOnReadout = False
weightReading = ''
weightStack = []
dateStack = deque()
timeStack = deque()
netWeightStack = deque()
tareWeightStack = deque()
grossWeightStack = deque()
weightStack.append(dateStack)
weightStack.append(timeStack)
weightStack.append(netWeightStack)
weightStack.append(tareWeightStack)
weightStack.append(grossWeightStack)


def initializePressPrintToExcel():
    thread = threading.Thread(target=readPort)   # Uncomment these two lines for C18 2 Log to Excel
    thread.daemon = True
    thread.start()      # Uncomment these two lines for C18 2 Log to Excel


def readPort():
    global wasPrintButtonPressedOnReadout
    global weightReading
    global com2
    global start

    reading = None
    #print("Thread Created")
    while True:
        if com2 != '':
            try:
                if not com2.is_open:
                    break
            except (OSError, serial.SerialException, AttributeError, serial.serialutil.SerialException) as error:
                pass


            try:
                command.read(com2)
                reading = readComPort(com2)
                # if reading:  # Only tested this if statement for C18 = 2.
                #     # This verifies if the reading are all byte data of weight.
                #     # Sometimes wireless scales adds charters.
                #     for i in range(len(reading)):
                #         try:
                #             reading[i].decode()
                #         except (UnicodeDecodeError, AttributeError):
                #             del reading[i]
                if any(x in reading.decode()[1:] for x in ['PT']) or keyboard.is_pressed('f6'):  # This checks if Print button is pressed
                    wasPrintButtonPressedOnReadout = True
                    time.sleep(1)  # Sleep to prevent double press

            except (OSError, serial.SerialException, AttributeError, serial.serialutil.SerialException,
                    UnicodeDecodeError) as error:
                #reading = ''
                pass
            weightReading = reading


def c18_2LogToExcel(inputs):
    df = pd.DataFrame({'Date': inputs[0], 'Time': inputs[1], 'Net Weight (lb)': inputs[2],
                       'Tare Weight (lb)': inputs[3], 'Gross Weight (lb)': inputs[4]})

    dataToExcel = pd.ExcelWriter(fileEntry.get(), engine='xlsxwriter')  # Need to catch error for invalid file type
    df.to_excel(dataToExcel, sheet_name='Sheet1', index=False)

    # Section below is to format excel
    workbook = dataToExcel.book
    worksheet = dataToExcel.sheets['Sheet1']

    column_settings = [{'header': column} for column in df.columns]
    (max_row, max_col) = df.shape

    worksheet.add_table(0, 0, max_row, max_col - 1, {'columns': column_settings})

    for i, col in enumerate(df.columns):
        column_len = max(df[col].astype(str).str.len().max(), len(col) + 5)
        worksheet.set_column(i, i, column_len)

    dataToExcel.close()


def pressPrintToExcel():
    global weightReading

    if len(weightReading) >= 1:  # This section for C18 = 2 -------------------------------------------------
        global weightStack

        for i in range(len(weightReading)):  # Not sure if this can be done without for loop
            logOutput.insert(INSERT, weightReading[i].decode())
            logOutput.see('end')  # Scrolls text to the bottom


        if len(weightReading) == 5:  # If they pressed print once there will be 5 elements coming from readout.

            dateStack.append(weightReading[0].decode().split()[1])
            timeStack.append(weightReading[1].decode().split()[1])
            grossWeightStack.append(float(weightReading[2].decode().split()[1]))
            tareWeightStack.append(' ')
            netWeightStack.append(' ')

        if len(weightReading) == 7:  # If they pressed tare there will be 7 elements coming from readout
            dateStack.append(weightReading[0].decode().split()[1])
            timeStack.append(weightReading[1].decode().split()[1])
            grossWeightStack.append(float(weightReading[2].decode().split()[1]))
            tareWeightStack.append(float(weightReading[3].decode().split()[1]))
            netWeightStack.append(float(weightReading[4].decode().split()[1]))


        weightReading = ''

        inputsWeights = [dateStack, timeStack, netWeightStack, tareWeightStack, grossWeightStack]
        c18_2LogToExcel(inputsWeights)
        # This section for C18 = 2 -------------------------------------------------------------




previousComPort = clickComPort.get()


# End log + Print ------------------------------------------------------------------------------------------------------
initializePressPrintToExcel()  # Initialize Thread for reading

number = '01993999901234503101000262131503102141457354'  # This is sample GS1-Code128 barcode
#my_code = EAN13(number, writer=ImageWriter())

barcode_format = barcode.get_barcode_class('gs1_128')
my_barcode = barcode_format(number, writer=ImageWriter())
my_barcode.save("generated_barcode")

# printers = win32print.EnumPrinters(win32print.PRINTER_ENUM_LOCAL)  # List out all printers
# for printer in printers:
#     print(printer)


wmi = win32com.client.GetObject('winmgmts:')
controllers = wmi.InstancesOf('Win32_USBControllerDevice')  # This sees all USB Controllers in device manager

# for controller in controllers:
#     print(controller.Dependent)
#     print("\n")

#printerName = 'USB\\VID_0FE6&PID_811E\\0E08F5C20506010073170000550801B6'
printerName = r'\\?\usb#vid_0fe6&pid_811e#0e08f5c20506010073170000550801b6#{28d78fad-5a12-11d1-ae5b-0000f803a8c2}'
#handle = win32print.OpenPrinter(win32print.GetDefaultPrinter())  # This opens default windows printer

with open('C:\\Users\\Cap\\PycharmProjects\\New Data Logging Software\\newDataLoggingSoftware\\test.xml', 'rb') as f:
    data = f.read()

#handle = ctypes.windll.kernel32.CreateFileW(printerName, ctypes.c_uint32(0x80000000), 0, None, ctypes.c_uint32(3), 0, None)
#xmlData = b'?xml version"1.0" encoding="UTF-8"?><root><element>data</element></root>'
#bufferSize = len(data)
#bytesWritten = ctypes.c_ulong(0)
#success = ctypes.windll.kernel32.WriteFile(handle, data, bufferSize, ctypes.byref(bytesWritten), None)

printer = usb.core.find(idVendor='0x0fe6', idProduct='0x811e')
#deviceId = printer.dev.dev

#print(printer)
#print(deviceId)

# deviceHandle = printer.open()
# deviceHandle.write(data)
# deviceHandle.close()

engine = ResponsiveVoice(lang=ResponsiveVoice.SPANISH_ES)
engine.say("""

Hola mi nombre es Maria

""",
           gender=ResponsiveVoice.FEMALE,
           rate=0.45, mp3_file='testFile.mp3')


def refreshMenu():
    global connectionType
    global comPortDropDown
    global start
    global com2
    global weightReading
    global previousComPort
    global wasPrintButtonPressedOnReadout


    if previousComPort != clickComPort.get():
        if com2 != '':  # This if statement is a test
            try:
                com2.close()  # Tries to close open com port. Sometimes might not work.
            except (OSError, serial.SerialException, AttributeError, serial.serialutil.SerialException) as error:
                pass

            com2 = openComPort(clickComPort.get())
            previousComPort = clickComPort.get()
            initializePressPrintToExcel()

        else:
            com2 = openComPort(clickComPort.get())
            previousComPort = clickComPort.get()

    if wasPrintButtonPressedOnReadout:  # If they press f1 or button it will print label
        wasPrintButtonPressedOnReadout = False  # This is the check flag being turned off
        printLog(inputsList)

    updateWeight(weightReading)  # Update weight to screen

    #pressPrintToExcel()

    # with concurrent.futures.ThreadPoolExecutor() as executor:
    #     th1 = executor.submit(openComPort, 'COM5')
    #     print(th1.result())

    #print(scanComPort())
    #weight = openWifiPort()
    #print(weight)

    # if (time.time() - start) > 10:  # This will update com ports after set period of time
    #     start = time.time()
    #     for p in serial.tools.list_ports.comports():
    #         #print(p.device)  # Print available com ports. Faster than my function but not tested with mac or linux yet.
    #         pass
        #print(serial.tools.list_ports.comports())
        # port = updateComPortOrWifi(connectionType, comPortDropDown)  # Also have it when com port open to not run

    window.update()
    weightOutput.delete(0.0, END)


while True:
    refreshMenu()

window.mainloop()