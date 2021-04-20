# Scannen
This application drives a Canon DR_M260 in order to scan, process(Invoicing &amp; Emailing Delivery Confirmation ) and archive delivery documents.
***********************************************************************************************
![alt text](https://github.com/dotH55/Scannen/blob/main/Image.png?raw=true)

Scannen.py: Implements a GUI using Kivy

Scan.py: This program scans until the device is out of paper.
It also convert the resulting png file into PDFs while 
renaming then based on each file's barcode

ProcessDR.py: This python program would authentify a PDF file using its name. The File name 
should be formated as scan_"Order#"--"Location"_0001.pdf. Example: scan_43160--Gro.pdf.
After a successful authentification, the program would get a serial number & email address
from the server the server. It would then Insert the packing slip into the company database. 
The final database work would be to release Packing slips into invoices. The program also 
emails a comfirmation email to the customer.It also has a method that deletes files 
older than 3 years.
A window notification appears when the program starts. In case of an incorrect file name,
A window notication would appear. The program also opens the folder where the file is located.


Requirements
pip install pypiwin32
pip install PyPDF2
pip install pyodbc
pip install --upgrade pip wheel setuptools
pip install docutils pygments pypiwin32 kivy.deps.sdl2 kivy.deps.glew
pip install kivy.deps.gstreamer
pip install kivy.deps.angle
pip install pygame
pip install kivy

Installation
1- Double Click Scannen 2.0.exe and 
   Follow the instructions

3- To start program, double click Scannen.exe

Notes

Exe with cmd
pyinstaller -F Packing_Slip.py --hidden-import=pypiwin32 ^ --hidden-import=PyPDF2 ^ 
--hidden-import=watchdog ^ --hidden-import=pyodbc ^ --hidden-import=win32api

Exe without cmd
pyinstaller -F Scannen.pyw --hidden-import=pypiwin32 ^ --hidden-import=PyPDF2 ^ --hidden-import=watchdog ^ --hidden-import=pyodbc ^ --hidden-import=win32api ^ --hidden-import=kivy ^ --hidden-import=pygame ^ --hidden-import=kivy.deps.angle ^ --hidden-import=kivy.deps.gstreamer ^ --hidden-import=pygments ^ --hidden-import=kivy.deps.sdl2 ^ --hidden-import=kivy.deps.glew ^ --hidden-import=docutils
