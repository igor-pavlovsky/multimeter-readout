# -*- coding: utf-8 -*-

"""
Application to read and save data from 
Velleman DVM345DI or compatible multimeter

Runs on Windows, tested to work with Python 3+

@author: Igor Pavlovsky
"""

import winreg #to check serial ports
import itertools #to check serial ports
import serial
import threading #for timed com port check
import sys #for stdout flush
import datetime #for timestamp
from PySide import QtGui, QtCore
import time #for delay/sleep and tic/toc function         

# get a list of COM ports available on a Windows machine    
def checkPorts():          
    ComPorts = []
    path = 'HARDWARE\\DEVICEMAP\\SERIALCOMM'
    try:
        key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, path)
        for i in itertools.count():
            try:
                val = winreg.EnumValue(key, i)
                if str(val[0]).count("Device"):
                    ComPorts.append(str(val[1]))
            except EnvironmentError:
                break
        ComPorts.sort()
    except WindowsError:
        print ("COM Ports Not Found")
        ComPorts.append("-Ports N/A-")
    
    return ComPorts
 
# read a line of characters from COM port buffer   
def readLine(ser):
    dat = ""
    while 1:
        ch = chr(int.from_bytes(ser.read(), byteorder='big'))
        # this hack is to detect timeout
        if(ch == '\0') or (len(ch) == 0):  
            return str('Timeout')     
        dat += ch
        if(ch == '\r'):  
            break
    return str(dat[3:10])

# Initialize COM port using device protocol settings
def initComport(COM):

    portName = "COM" + str(COM)
    print ("Connecting to " + portName)
    sys.stdout.flush()
    try:
        COMport = serial.Serial(portName)
        COMport.baudrate = 600
        COMport.bytesize = 7
        COMport.parity = serial.PARITY_NONE
        COMport.stopbits = serial.STOPBITS_TWO
        COMport.timeout = 3
        COMport.setDTR(True)
        time.sleep(.5)
        COMport.setRTS(False)
        time.sleep(.5)
        COMport.flush()
        COMport.reset_input_buffer()
        time.sleep(.5)
        COMport.write(('\r').encode())
        time.sleep(.1)
        print ("Port ", COM, "is open\n")
        sys.stdout.flush()
        return COMport # pySerial object
    except serial.serialutil.SerialException:
        print ("Port not found")
        sys.stdout.flush()
        return 0
  
class ProgramSignals(QtCore.QObject):
    timer_signal = QtCore.Signal(int)
    
#instantiate the signal class
program_signals = ProgramSignals()


class Main_Prg(QtGui.QMainWindow,  QtCore.QObject):
    
    global COMport
    
    def __init__(self):
        super(Main_Prg, self).__init__()
        program_signals.timer_signal.connect(self.checkStatus)  
        self.ComPortList = checkPorts()  
        self.Port = None
        self.initUI()
    
    # define GUI controls
    def initUI(self):
        self.status = ""
        self.fname = ""
        self.saveAction = 0
        self.f = None
        self.lastsaved = "test.csv"
        
        self.setFont(QtGui.QFont("Fantasy", 12, QtGui.QFont.Bold))  
        self.setFixedSize(170,150)
        self.move(300,300)
        self.setWindowTitle('DVM345DI')
        self.setWindowFlags(self.windowFlags() | QtCore.Qt.CustomizeWindowHint)
        self.setWindowFlags(self.windowFlags() & ~QtCore.Qt.WindowMinMaxButtonsHint) 
        
        self.combo = QtGui.QComboBox(self)
        self.combo.addItems(self.ComPortList)
        self.combo.setGeometry(20, 20, 130, 25)
        
        self.btn1 = QtGui.QPushButton('CONNECT', self)
        self.btn1.setToolTip('Connect via COM port')
        self.btn1.setGeometry(20, 50, 130, 25)  
        self.btn1.clicked.connect(self.initPort)
        
        self.lbl1 = QtGui.QLabel(self)
        self.lbl1.move(20, 80)
        self.lbl1.setText("Value")

        self.ed1 = QtGui.QLineEdit(self)
        self.ed1.setGeometry(70, 80, 80, 25)
        self.ed1.setEnabled(False)
        self.ed1.setText("")
        
        self.btn2 = QtGui.QPushButton('SAVE', self)
        self.btn2.setToolTip('Save data as csv file')
        self.btn2.setGeometry(20, 110, 130, 25)  
        self.btn2.clicked.connect(self.doSaving)
        self.btn2.setDisabled(True)

        self.skinDisconnected() 
        self.show()    
    
    # change window skin if device is connected
    def skinConnected(self):
        self.btn1.setDisabled(True)
        self.btn1.setText("CONNECTED")
        self.btn2.setDisabled(False)
        self.setStyleSheet("""
        QMainWindow{      
             font: 14px;
             font: bold;
             color: #000088; /* font */
             background: #ffeecc; /* bisque */
             }
        QPushButton {      
             font: 14px;
             font: bold;
             }
        QLineEdit {      
             font: 14px;
             font: bold;
             color: #000088; /* font */
             background:white;
             }
         QLabel {      
             font: 14px;
             font: bold;
             color: #000000; /* font */
             }
         QComboBox {      
             font: 14px;
             font: bold;
             color: #000000; /* font */
             }
        """
        )
       
     # change window skin if device is disconnected
    def skinDisconnected(self):
        self.btn1.setDisabled(False)
        self.btn1.setText("CONNECT")
        self.btn2.setDisabled(True)
        self.setStyleSheet("""
        QMainWindow{      
             font: 14px;
             font: bold;
             color: #000000; /* font */
             background: #bbaaaa;
             }
        QPushButton {      
             font: 14px;
             font: bold;
             }
        QLineEdit {      
             font: 14px;
             font: bold;
             color: #000088; /* font */
             background:white;
             }
         QLabel {      
             font: 14px;
             font: bold;
             color: #000000; /* font */
             }
         QComboBox {      
             font: 14px;
             font: bold;
             color: #000000; /* font */
             }
        """
        )
            
    # Initialize com port
    def initPort(self):    
        
        selected_port = self.ComPortList[self.combo.currentIndex()] #e.g. 'COM2'
        
        COM = int(selected_port[3:]) # COM is a number
        
        self.Port =  initComport(COM)
        
        if (self.Port):
            self.skinConnected()
            time.sleep(.1)
            # start readout timer loop
            self.initCheck()
        
    # Define s signal function to read port
    def initCheck(self):
        program_signals.timer_signal.emit(0)
        
    # Check if new data is available
    def checkStatus(self):
        
        # read data and display in the QLineEdit widget
        self.status = self.watchReadout(self.Port)
        
        if (self.status == 'Timeout'):
            self.timeOut()
            self.skinDisconnected()
        else:
            # restart the polling timer
            threading.Timer(1.0, self.initCheck).start()
            
        self.ed1.setText(str(self.status))
        # save data if needed
        if (self.fname):
           _time = datetime.datetime.now().strftime('%x %X')
           line = _time + ',' + self.status  + '\n'
           print (line)
           self.f.write(line)
           sys.stdout.flush()    
    
    # Read data from port
    def watchReadout(self, port):
        # allow non-blocking application
        time.sleep(0.2)
        # request readout
        if port.isOpen():
            port.write(('\r').encode())
        time.sleep(0.2)
        # read data
        if port.isOpen():        
            self.status = str(readLine(port))
            return self.status
        else:
            return "0"      
    
    # save data line upon receiving
    def doSaving(self):
        # change state of the controls
        if (self.saveAction):
            self.saveAction = 0
            if self.fname and not self.f.closed:
                self.f.close()
                self.fname = ""
            self.btn2.setText("SAVE")
        else:
            self.saveAction = 1
            self.btn2.setText("STOP")
            self.showSaveDialog()
    
    # self-explanatory
    def showSaveDialog(self):
        self.fname, _ = QtGui.QFileDialog.getSaveFileName(self, 'Save file', self.lastsaved,
                   ("CSV files (*.csv);;All files (*.*)"))
        if (self.fname):
            self.lastsaved = self.fname
            self.f = open(self.fname, 'a+b')
    
    # process port timeout
    def timeOut(self):
        msgBox = QtGui.QMessageBox.warning(self, 'Exit Monitor',
            "Device not connected", QtGui.QMessageBox.Ok, QtGui.QMessageBox.Ok)
        
        if msgBox == QtGui.QMessageBox.Ok:
            if self.Port is not None:
                self.Port.close()
                print ("COM Port closed due to read timeout")
                sys.stdout.flush()  
            if self.fname and not self.f.closed:
                self.f.close()

    # close the application       
    def closeEvent(self, event): 
        # show closing dialog box
        reply = QtGui.QMessageBox.question(self, 'Exit Monitor',
            "Are you ready to quit?", QtGui.QMessageBox.Yes | 
            QtGui.QMessageBox.No, QtGui.QMessageBox.No)
        
        if reply == QtGui.QMessageBox.Yes:
            if self.Port is not None:
                self.Port.close()
                print ("COM Port closed")
            if self.fname and not self.f.closed:
                self.f.close()
                print ("Data file closed")
            event.accept()
        else:
            event.ignore() 


def main():       
    # usual stuff to initiate and quit GUI
    QtGui.QApplication.setStyle("plastique")
    qApp = QtGui.QApplication.instance()
    if qApp is None:
        qApp = QtGui.QApplication(sys.argv)
        
    _ = Main_Prg()
    
    qApp.exec_()

if __name__ == '__main__':
    main()