#Author: Gus Tahara-Edmonds
#Date: Oct-Nov 2019
#Purpose: Reads data from analog inputs on a U or T series Labjack. Originally developed
#for recording data during tests for UBC Rocket. Has remapping, user settings for frequency, 
#enabled pins, etc., and lots of error checking. Uses Kivy for app UI, saves data to Excel file. 
#Hopefully adding stream mode support for faster input rates soon. 

import os #reading/writing files
import sys 
import ljm #for labjack
import ctypes #for popup boxes
from threading import Thread
import xlsxwriter #for excel

import time
msCurrentTime = lambda: int(round(time.time() * 1000))

#GUI imports
from kivy.app import App
from kivy.core.window import Window
from kivy.uix.label import Label
from kivy.uix.popup import Popup
from kivy.uix.button import Button
from kivy.uix.checkbox import CheckBox 
from kivy.uix.textinput import TextInput
from kivy.uix.gridlayout import GridLayout
from kivy.config import Config

 #Window settings (height width, color)
Config.set('graphics', 'width', '1300')               
Config.set('graphics', 'height', '750')
Config.write()
DEFAULT = (0, 0, 0, 1)
HIGHLIGHT = (0, 0, 0.2, 1)

global handle                           #reference to labjack
global reading                          #whether or not the labjack is currently reading data
reading = False

global pinCount                         #number of analog sensors
pinCount = 6

#GUI element references for each pin
global GUI_sensorNames                  
GUI_sensorNames = [0] * pinCount
global GUI_enabledPins               
GUI_enabledPins = [0] * pinCount
global GUI_freqs                       
GUI_freqs = [0] * pinCount  
global GUI_useStreams                
GUI_useStreams = [0] * pinCount
global GUI_outputValues               
GUI_outputValues = [0] * pinCount
global GUI_units                        
GUI_units = [0] * pinCount

global GUI_m
GUI_m = [0] * pinCount
global GUI_b
GUI_b = [0] * pinCount

global m
m = [0] * pinCount
global b
b = [0] * pinCount

ains = [0, 1, 2, 3, 4, 5]           #ain value for each pin (i.e. AIN0, AIN1, etc.)
global ainEnableds                  #whether or not each pin is enabled
ainEnableds = [False, False, False, False, False, False]
global freqs                        #frequency value for each pin to be read in at
freqs = [1, 1, 1, 1, 1, 1]
global useStreams                   #whether or not each pin is using stream mode
useStreams = [False, False, False, False, False, False]

global write2File                   #whether or not data should be written to excel file
write2File = False

#excel file references 
global workbook
global worksheet

'''-------------------------------GUI---------------------------------------'''
class MyGrid(GridLayout):
    #this function essentially just sets up all the GUI elements. it also links the buttons to the checkbox/button functions below
    def __init__(self, **kwargs):   
        Window.clearcolor = DEFAULT
        
        super(MyGrid, self).__init__(**kwargs)
        
        #Header
        self.cols = 1
        self.add_widget(Label(text="Rocket Testing Interface", font_size=24, size_hint_y=0.1))

        #Power output config
        #region
        self.add_widget(Label(text="Power Output", font_size=18, size_hint_y=0.05))
        header = GridLayout(cols=6, size_hint_y=0.12)

        header.add_widget(Label(text=""))
        header.add_widget(Label(text=""))
        header.add_widget(Label(text="DAC0"))
        header.add_widget(Label(text="DAC1"))
        header.add_widget(Label(text=""))
        header.add_widget(Label(text=""))

        header.add_widget(Label(text=""))
        header.add_widget(Label(text=""))
        self.powerSupply1 = TextInput(text="0", multiline = False)
        header.add_widget(self.powerSupply1)
        self.powerSupply2 = TextInput(text="0", multiline = False)
        header.add_widget(self.powerSupply2)
        header.add_widget(Label(text=""))
        header.add_widget(Label(text=""))

        header.add_widget(Label(text=""))
        header.add_widget(Label(text=""))
        header.add_widget(Label(text=""))
        header.add_widget(Label(text=""))
        header.add_widget(Label(text=""))
        header.add_widget(Label(text=""))
        self.add_widget(header)
        #endregion

        #Channel input config
        #region
        main = GridLayout(cols=3, size_hint_y=0.4)
        main.add_widget(Label(text="Channel Input", font_size=18, size_hint_y=0.15))
        main.add_widget(Label(text="Remapping", font_size=18, size_hint_y=0.15))
        main.add_widget(Label(text="Channel Output", font_size=18, size_hint_y=0.15))

        inConfig = GridLayout(cols=3)
        inConfig.add_widget(Label(text="Channel", font_size=16))
        inConfig.add_widget(Label(text="Sensor name", font_size=16))
        inConfig.add_widget(Label(text="Enabled?", font_size=16))

        global GUI_sensorNames
        global GUI_enabledPins
        for i in range(pinCount):
            inConfig.add_widget(Label(text="AIN" + str(ains[i]), size_hint_x=0.1))

            GUI_sensorNames[i] = TextInput(text="-", disabled=True)
            inConfig.add_widget(GUI_sensorNames[i])

            GUI_enabledPins[i] = CheckBox(active = False, size_hint_x=0.1)
            GUI_enabledPins[i].bind(active = self.onCheck_PIN)
            inConfig.add_widget(GUI_enabledPins[i]) 

        main.add_widget(inConfig)
        #endregion

        #Remap config
        #region
        remapConfig = GridLayout(cols=4)
        remapConfig.add_widget(Label(text=""))
        remapConfig.add_widget(Label(text="m"))
        remapConfig.add_widget(Label(text="b"))
        remapConfig.add_widget(Label(text=""))

        remapConfig.add_widget(Label(text=""))
        remapConfig.add_widget(Label(text=""))
        remapConfig.add_widget(Label(text=""))
        remapConfig.add_widget(Label(text=""))

        for i in range(6):
            remapConfig.add_widget(Label(text=""))
            GUI_m[i] = TextInput(text="1", multiline=False, disabled=True)
            remapConfig.add_widget(GUI_m[i])
            GUI_b[i] = TextInput(text="0", multiline=False, disabled=True)
            remapConfig.add_widget(GUI_b[i])
            remapConfig.add_widget(Label(text=""))

        main.add_widget(remapConfig)
        #endregion

        #Channel output config
        #region
        outConfig = GridLayout(cols=4)
        outConfig.add_widget(Label(text="Frequency", font_size=16))
        outConfig.add_widget(Label(text="High Freq?", font_size=16))
        outConfig.add_widget(Label(text="Output Value", font_size=16))
        outConfig.add_widget(Label(text="Units", font_size=16))

        self.thermoFreq = TextInput(text="-", multiline = False)
        outConfig.add_widget(self.thermoFreq)
        self.thermoStream = CheckBox(active = False, size_hint_x=0.1)
        outConfig.add_widget(self.thermoStream)
        self.thermoOutput = Label(text="0")
        outConfig.add_widget(self.thermoOutput)
        outConfig.add_widget(Label(text=""))

        global GUI_freqs
        global GUI_useStreams
        global GUI_outputValues
        global GUI_units
        for i in range(pinCount):
            GUI_freqs[i] = TextInput(text="1", multiline=False, disabled=True)
            outConfig.add_widget(GUI_freqs[i])

            GUI_useStreams[i] = CheckBox(active = False, size_hint_x=0.1)
            GUI_useStreams[i].bind(active = self.onCheck_UseHF)
            outConfig.add_widget(GUI_useStreams[i])

            GUI_outputValues[i] = Label(text="0")
            outConfig.add_widget(GUI_outputValues[i])

            GUI_units[i] = TextInput(text="-", multiline=False, disabled=True)
            outConfig.add_widget(GUI_units[i])

        main.add_widget(outConfig)
        self.add_widget(main)
        #endregion

        #Bottom buttons and file config
        #region
        bottom = GridLayout(cols=2, size_hint_y=0.25)
        bottom.add_widget(Label(text=""))
        bottom.add_widget(Label(text=""))

        bottom.add_widget(Label(text="High Frequency Value"))
        self.highFrequencyValue = TextInput(text="-", multiline=False, disabled=False)
        bottom.add_widget(self.highFrequencyValue)

        bottom.add_widget(Label(text="Write to Excel File?"))
        self.exportBox = CheckBox(active = False)
        self.exportBox.bind(active = self.onCheck_Write2File)
        bottom.add_widget(self.exportBox) 

        bottom.add_widget(Label(text="File name:"))
        self.fileName = TextInput(text="-", multiline=False, disabled=False)
        bottom.add_widget(self.fileName)

        self.testButton = Button(text="User Input is Valid?", font_size=20)
        self.testButton.bind(on_press=self.onPressed_CheckValid)
        bottom.add_widget(self.testButton)

        self.createFileButton = Button(text="Create Excel File", font_size=20)
        self.createFileButton.bind(on_press=self.onPressed_CreateFile)
        bottom.add_widget(self.createFileButton)

        self.readButton = Button(text="Read Data", font_size=20)
        self.readButton.bind(on_press=self.onPressed_Read)
        bottom.add_widget(self.readButton)

        self.exitButton = Button(text="Save File & Exit", font_size=20)
        self.exitButton.bind(on_press=self.onPressed_Exit)
        bottom.add_widget(self.exitButton)

        self.add_widget(bottom)
        #endregion

        #Init main program
        updateUI(self)
        init(self)
                             
    #--------------------------Checkbox Event Setup---------------------------'''
    #region
    #we subscribe the checking of a textbox element to these functions so program knows when...
    #AIN enabled/disabled
    def onCheck_PIN(self, checkboxInstance, isActive):
        global ainEnableds

        for i in range(pinCount):
            ainEnableds[i] = GUI_enabledPins[i].active

        updateUI(self)
    
    #stream mode enabled/disabled
    def onCheck_UseHF(self, checkboxInstance, isActive): 
        global useStreams

        for i in range(pinCount):
            useStreams[i] = GUI_useStreams[i].active

        updateUI(self)
    #Write To File enabled/disabled
    def onCheck_Write2File(self, checkboxInstance, isActive):
        global write2File
        write2File = isActive
        updateUI(self)
    #endregion
    
    #-----------------------------------Buttons-------------------------------'''   
    #region
    def onPressed_Read(self, instance):   #when Read Data pressed, start reading from the labjack
        global reading
        reading = not reading

        if(reading):
            if not start(self, True):
                reading = False
                return

            start(self, False)

        updateUI(self)

    def onPressed_CheckValid(self, instance):   #when Check Valid is pressed, make sure all the user input checks out
        if start(self, True):
            popup("Ready to Go", "All input appears valid")
        
        updateUI(self)

    def onPressed_CreateFile(self, instance):   #when Create File is pressed, create a new excel file and save the old one
        global workbook    
        
        try:
            workbook.close()
        except:
            pass

        workbook = xlsxwriter.Workbook(self.fileName.text + '.xlsx')
        popup("Success", "Excel file created")

    def onPressed_Exit(self, instance): #when Exit is pressed, quit program
        quit()
    #endregion

'''-------------------------------Helper Functions-------------------------------'''  
#region
#this is called whenever any GUI elements are updated. performance was not an issue so 
#basically just updates everything to avoid any errors or inconsistencies
def updateUI(self):
    self.powerSupply1.disabled = reading
    self.powerSupply2.disabled = reading

    disableHighFreq = True
    for i in range(pinCount):
        GUI_enabledPins[i].disabled = reading
        GUI_sensorNames[i].disabled = reading or not ainEnableds[i]
        GUI_freqs[i].disabled = reading or not ainEnableds[i] or useStreams[i]
        GUI_useStreams[i].disabled = reading or not ainEnableds[i]
        GUI_outputValues[i].disabled = not ainEnableds[i]
        GUI_units[i].disabled = reading or not ainEnableds[i]

        GUI_m[i].disabled = reading or not ainEnableds[i]
        GUI_b[i].disabled = reading or not ainEnableds[i]

        if ainEnableds[i] and useStreams[i]:
            disableHighFreq = False

    self.highFrequencyValue.disabled = reading or disableHighFreq
    self.fileName.disabled = reading or not write2File
    self.exportBox.disabled = reading

    self.testButton.disabled = reading
    self.createFileButton.disabled = reading or not write2File
    self.exitButton.disabled = reading

    if reading:
        self.readButton.text = "Stop read"
        Window.clearcolor = HIGHLIGHT

    else:
        self.readButton.text = "Read Data"
        Window.clearcolor = DEFAULT

#sets the analog output value
def setDAC(dacNum, volts):
    name = "DAC" + str(dacNum)
    ljm.eWriteName(handle, name, volts)

#given a slope m and y-int b, remap incoming value
def remap(value, m, b):
        return (value * m) + b

#when quit button is pressed close stuff down, save excel file, and quit
def quit():
    reading = False
    App.get_running_app().stop() 

    try:
        workbook.close()
    except:
        pass

    try:
        ljm.close(handle)
    except:
        pass

def popup(title, text):
    ctypes.windll.user32.MessageBoxW(0, text, title, 1)
#endregion

'''-------------------------------Main Functions-------------------------------''' 
#region
#set up labjack data when starting program (this function is called right after GUI elements are created)
def init(self):
    global handle

    try:
        handle = ljm.openS("ANY", "ANY", "ANY")
    except:
        popup("Error", "Could not find a connected Labjack")
        return

    info = ljm.getHandleInfo(handle)
    print("Opened a LabJack with Device type: %i, Connection type: %i,\n"
      "Serial number: %i, IP address: %s, Port: %i,\nMax bytes per MB: %i" %
      (info[0], info[1], info[2], ljm.numberToIP(info[3]), info[4], info[5]))

#main function that runs to collect the data. The steps this function does are:
#1. Set analog output (DAC) values to be on.
#2. Create and format excel file 
#3. Run through each of the pin's values to make sure nothing is incorrect
#4. Start seperate timed threads which read from the analog inputs (AINs)
#*this function can also just do the error checking
def start(self, test):
    #Analog output init 
    #region
    try:
        handle
    except:
        popup("Error", "Labjack is not connected")
        return False

    dac0Volts = 0
    dac1Volts = 0
    try:
        dac0Volts = float(self.powerSupply1.text)
    except:
        popup("Error", "DAC0 voltage input is invalid")
        return False

    try:
        dac1Volts = float(self.powerSupply2.text)
    except:
        popup("Error", "DAC1 voltage input is invalid")
        return False

    if not test:
        setDAC(0, dac0Volts)
        setDAC(1, dac1Volts)
    #endregion

    #File output init
    #region
    global worksheet

    if write2File:
        try:
           workbook
        except:
            popup("Error", "No Excel file created")
            return False

        if not test:
            worksheet = workbook.add_worksheet()
            x = 0
            for i in range(6):
                 if ainEnableds[i]:
                    name = GUI_sensorNames[i].text
                    unit = GUI_units[i].text
                    worksheet.write(0, x * 3, name + " (" + unit + ")")
                    worksheet.write(0, x * 3 + 1, "< Time (ms) for " + name)
                    x += 1

    #endregion
    
    #Pin error checking
    rates = [1, 1, 1, 1, 1, 1]
    streamUsed = False
    for i in range(pinCount):
        if ainEnableds[i]:
            if useStreams[i]:
                streamUsed = True
            else:
                try:
                    r = 1.0 / float(GUI_freqs[i].text)
                    rates[i] = max(r, 0.001)
                except:
                    popup("Error", "PIN" + str(i) + " update rate is invalid")
                    return False 

    if streamUsed:
        try:
            int(self.highFrequencyValue)
        except:
            popup("Error", "High frequency value is invalid")
            return False

    #Start reading at inputted interval
    streamChannelNames = []
    x = 0
    if not test: 
        for i in range(pinCount):
            if ainEnableds[i]:
                if useStreams[i]:
                    streamChannelNames += ["AIN" + str(i)]
                else:
                    Thread(target=onIntervalReadChannel, args = (GUI_outputValues[i], ains[i], rates[i], x)).start()
                    x += 1
                    
    #init remapping data
    global m, b
    for i in range(pinCount):
        if ainEnableds[i]:
            try:
                m[i] = int(GUI_m[i].text)
                b[i] = int(GUI_b[i].text)
            except:
                popup("Error", "Remapping input for AIN" + str(i) + " is invalid")
                return False

    return True

#runs on seperate thread at a certain time interval based on user-inputted frequency
#reads from analog input on the Labjack, remaps them, and writes them to the excel file (if desired)
def onIntervalReadChannel(labelOutput, ainNum, interval, columnIndex):
    data = []
    times = []
    startTime = msCurrentTime()
    skippedCount = 0
    count = 0

    ljm.startInterval(ainNum, int(1000000 * interval))  # Delay between readings (in microseconds

    while reading:
        result = 0;
            
        try:
            result = ljm.eReadName(handle, "AIN" + str(ainNum))
        except:
            print("No Labjack found")
            return


        result = remap(result, m[ainNum], b[ainNum])
        result = round(result, 4)
        labelOutput.text = str(result)

        data += [result]
        times += [msCurrentTime() - startTime]
        count += 1

        skippedCount += ljm.waitForNextInterval(ainNum)

    if write2File:
        for i in range(count):
            worksheet.write_number(i + 1, columnIndex * 3, data[i])
            worksheet.write_number(i + 1, columnIndex * 3 + 1, times[i])

    print("Skipped: " + str(skippedCount))
#endregion

'''--------------------------Initializes Program----------------------------'''    
class MyApp(App):
    def build(self):
        return MyGrid()

if __name__ == "__main__":
    MyApp().run()
