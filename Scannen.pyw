# ======================================================================
# ----------------------|  File: Scannen.py   |-------------------------
# ----------------------|   Author: dotH55    |-------------------------
# ======================================================================

# Description: This program constructs the Graphical User Interface of 
# Scannen. Other functions such as Scan & ProcessDR are invoked 
# externaly (os.system(Python Scan.py)).

# Imports
from win32api import *
from win32gui import *
from email import encoders
from datetime import datetime
from PyPDF2 import PdfFileReader
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from watchdog.observers import Observer
from email.mime.multipart import MIMEMultipart
from watchdog.events import FileSystemEventHandler
import threading, win32con, zipfile, logging, struct, ctypes
import email, imaplib, pyodbc, datetime, PyPDF2, win32process
import re, os, sys, ssl, time, email, smtplib, shutil, subprocess

# Kivy Imports
from kivy.config import Config
Config.set('graphics', 'fullscreen', 'fake')
#Config.set('graphics', 'position', 'custom')
#Config.set('graphics', 'top', '300')
#Config.set('graphics', 'left', '300')
from kivy.app import App
from kivy.uix.image import Image
from kivy.uix.label import Label
from kivy.uix.widget import Widget
from kivy.uix.button import Button
from kivy.core.window import Window
from kivy.uix.textinput import TextInput
from kivy.uix.boxlayout import BoxLayout
from kivy.uix.gridlayout import GridLayout
from kivy.graphics import Color, Rectangle
from kivy.uix.floatlayout import FloatLayout

# Import Classes
import WindowsBalloonTip

# Window Notification
NOTIFICATION = WindowsBalloonTip.WindowsBalloonTip()

class ImageLayout(FloatLayout):

    def __init__(self,**args):
        super(ImageLayout, self).__init__(**args)

        with self.canvas.before:
            Color(0, 0, 0, 0)
            self.rect=Rectangle(size=self.size,pos=self.pos)

        self.bind(size=self.updates,pos=self.updates)
    def updates(self,instance,value):
        self.rect.size=instance.size
        self.rect.pos=instance.pos

class Scannen(App):

    image = Image(source = "Icons\\DSI1.png")

    def build(self):

        # Init ImageLayout
        imageLayout = ImageLayout()
        imageLayout.add_widget(self.image)

        # Init Root
        root = BoxLayout(orientation='horizontal')
        root.add_widget(imageLayout)

        # Buttons
        scan = Button(background_normal = "Icons\\Scan.png", size_hint=(None, None))
        procressDR = Button(background_normal = "Icons\\ProcessDR.png", size_hint=(None, None))
        prev = Button(background_normal = "Icons\\LeftArrow.png", size_hint=(None, None))
        nex = Button(background_normal = "Icons\\RightArrow.png", size_hint=(None, None))
        search = Button(background_normal = "Icons\\SearchDR.png", size_hint=(None, None))
        process = Button(background_normal = "Icons\\Process.png", size_hint=(None, None))
        close = Button(background_normal = "Icons\\Close.png", size_hint=(None, None))
        
        # Bind Buttons
        scan.bind(on_press=self.InitScan)
        procressDR.bind(on_press=self.ProcessDRs)
        #search.bind(on_press=self.Search)
        #process.bind(on_press=self.Function)
        prev.bind(on_press=self.PreviousPicture) 
        nex.bind(on_press=self.NextPicture)
        close.bind(on_press=self.Close) 

        # Init Panel (GridLayout)
        panel = GridLayout(size_hint=(.5, .8))
        panel.cols = 2

        # Add buttons to panel
        panel.add_widget(prev)
        panel.add_widget(nex)
        panel.add_widget(scan)
        panel.add_widget(procressDR)
        panel.add_widget(search)
        panel.add_widget(close)
        panel.add_widget(process)

        # Add panel to root
        root.add_widget(panel)
        self.image.keep_ratio= True
        self.image.allow_stretch = False
        
        return root
    # End of Build

    # Start Scanning
    def InitScan(self, instance):
        NOTIFICATION.ShowWindow("Scannen", "Scanning...")
        os.system("python Scan.pyw")
    
    # Process Delivery Reports
    def ProcessDRs(self, instance):
        os.system("python ProcessDR.pyw")

    def Search(self, instance):
        os.system("python SearchDR.py")
    
    def Function(self, instance):
        os.system("python Function.py")

    # Change display
    def NextPicture(self, instance):
        self.image.source = "Icons\\DSI1.png"
    
    # Change display
    def PreviousPicture(self, instance):
        self.image.source = "Icons\\DSI2.png"

    # Teminate program
    def Close(self, instance):
        quit()

if __name__ == "__main__":
    try:
        Scannen().run()
    except Exception as e:
        print(e.args) 
        print(e.__cause__)