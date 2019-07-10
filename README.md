**Velleman DVM345DI multimeter readout with Python and PySide based GUI**

This repository provides a Python 3 script with a purpose to read and save data from a Velleman DVM345DI multimeter. 
This device is rebranded as CSI345, Global Specialties Pro-70, McVoice M-345pro, Mastech MAS345, Sinometer MAS345, 
and I expect that all of these would work with the app as well as Velleman. One unusual feature of this multimeter 
interface is that it has a very low communication baud rate of 600, which is unavailable as a settings parameter in 
many popular terminals such as Windows Hyperterminal.

![alt text](http://igorpavlovsky.com/wp-content/uploads/2019/06/velleman-1-e1559505392696-225x300.jpg)

The python script uses a call to the Windows register to obtain a list of available COM ports, and the script as written 
is not expected to work under MacOS out of the box. The app uses PySide to create a widget based GUI. If you have PyQt 
framework installed, you may need to take care of possible differences between the two Qt implementations. There is also 
some output to console for debugging, which can be removed if you don't need it.

![alt text](http://igorpavlovsky.com/wp-content/uploads/2019/06/velleman_4.png)
![alt text](http://igorpavlovsky.com/wp-content/uploads/2019/06/velleman_3.png)

The data can be saved in a csv file (default) with an Excel readable timestamp in the first column, such that the data 
could be opened in one click, and plotted easily as a function of time:

6/3/2019 22:57:49, -0.001

6/3/2019 22:57:51, 0.053

6/3/2019 22:57:52, 0.004

6/3/2019 22:57:54, -0.066


In order to transfer the application to another computer, I had most luck converting this code to *.exe file using cx_Freeze, 
saving it into a folder and then creating a desktop shortcut. A few simultaneous instances of the app launched as exe files 
worked flawlessly on a single Windows 7 machine. 
