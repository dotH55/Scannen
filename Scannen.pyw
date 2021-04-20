# ==============================================================================================================================================
# ----------------------------------------------------------------------------------------------------------------------------------------------
# -------------------------------------------------------------------File: Scannen.py------------------------------------------------------
# ------------------------------------------------------------------------Author: dotH55--------------------------------------------------------
# ----------------------------------------------------------------------------------------------------------------------------------------------
# ==============================================================================================================================================

# Description: This python program would authentify a PDF file using its name. The File name should be formated as scan_"Order#"--"Location".
# pdf. Example: scan_43160--Gro.pdf. After a successful authentification, the program would get a serial number & email address from the 
# server the server. It would then Insert the packing slip into the company database. The final database work would be to release Packing slips
# into invoices. The program also emails a comfirmation email to the customer. It also has a method that deletes files older than 3 years.
# Processed files are first stored locally before being moved to the external Z drive. A window notification appears when the program starts.
# In case of an incorrect file name, a window notication would appear. The program also opens the folder where the file is located.

# Imports
import sys, os 
from PIL import Image
from win32api import *
from win32gui import *
from email import encoders
from datetime import datetime
from PyPDF2 import PdfFileReader
from pyzbar.pyzbar import decode
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
import threading, win32con, zipfile, logging, struct, ctypes
import email, imaplib, pyodbc, datetime, PyPDF2, win32process
import re, ssl, time, email, smtplib, shutil, random, subprocess

# =====================================================================================================================================
# -------------------------------------------------------------File: GlobalVariables.py----------------------------------------------
# ----------------------------------------------------------------Author: dotH55-------------------------------------------------------
# =====================================================================================================================================

# Description: This File defines all global variables used in the 
# Scannnen program.

# WIA DEVICE TYPE
WIA_DEVICE_UNSPECIFIED = 0
WIA_DEVICE_SCANNER     = 1
WIA_DEVICE_CAMERA      = 2
WIA_DEVICE_VIDEO       = 3

# WIA IMAGE BIAS
WIA_MAXIMUM_QUALITY = 131072
WIA_MINIMUM_SIZE    = 65536

# WIA IMAGE INTENT
WIA_INTENT_UNSPECIFIED = 0
WIA_INTENT_COLOR       = 1
WIA_INTENT_GRAY        = 2
WIA_INTENT_TEXT        = 4

# WIA Format
WIA_FORMAT_PNG  = "{B96B3CAF-0728-11D3-9D7B-0000F81EF32E}"
WIA_FORMAT_BMP  = "{B96B3CAB-0728-11D3-9D7B-0000F81EF32E}"
WIA_FORMAT_GIF  = "{B96B3CB0-0728-11D3-9D7B-0000F81EF32E}"
WIA_FORMAT_JPEG = "{B96B3CAE-0728-11D3-9D7B-0000F81EF32E}"
WIA_FORMAT_TIFF = "{B96B3CB1-0728-11D3-9D7B-0000F81EF32E}"

# WIA Commands
WIA_COMMAND_TAKE_PICTURE     = "{AF933CAC-ACAD-11D2-A093-00C04F72DC3C}"
WIA_COMMAND_SYNCHRONIZE      = "{9B26B7B2-ACAD-11D2-A093-00C04F72DC3C}"
WIA_COMMAND_DELETE_ALL_ITEMS = "{E208C170-ACAD-11D2-A093-00C04F72DC3C}"
WIA_COMMAND_CHANGE_DOCUMENT  = "{04E725B0-ACAE-11D2-A093-00C04F72DC3C}"
WIA_COMMAND_UNLOAD_DOCUMENT  = "{1F3B3D8E-ACAE-11D2-A093-00C04F72DC3C}"

# WIA EVENT
WIA_EVENT_DEVICE_CONNECTED     = "{A28BBADE-64B6-11D2-A231-00C04FA31809}"
WAI_EVENT_DEVICE_DISCONNECTED  = "{143E4E83-6497-11D2-A231-00C04FA31809}"
WIA_EVENT_ITEM_CREATED         = "{4C8F4EF5-E14F-11D2-B326-00C04F68CE61}"
WIA_EVENT_ITEM_DELETED         = "{1D22A559-E14F-11D2-B326-00C04F68CE61}"
WIA_EVENT_SCAN_EMAIL_IMAGE     = "{C686DCEE-54F2-419E-9A27-2FC7F2E98F9E}"
WIA_EVENT_SCAN_FAX_IMAGE       = "{C00EB793-8C6E-11D2-977A-0000F87A926F}"
WIA_EVENT_SCAN_FILM_IMAGE      = "{9B2B662C-6185-438C-B68B-E39EE25E71CB}"
WIA_EVENT_SCAN_IMAGE           = "{A6C5A715-8C6E-11D2-977A-0000F87A926F}"
WIA_EVENT_SCAN_IMAGE_2         = "{FC4767C1-C8B3-48A2-9CFA-2E90CB3D3590}"
WIA_EVENT_SCAN_IMAGE_3         = "{154E27BE-B617-4653-ACC5-0FD7BD4C65CE}"
WIA_EVENT_SCAN_IMAGE_4         = "{A65B704A-7F3C-4447-A75D-8A26DFCA1FDF}"
WIA_EVENT_SCAN_OCR_IMAGE       = "{9D095B89-37D6-4877-AFED-62A297DC6DBE}"
WIA_EVENT_SCAN_PRINT_IMAGE     = "{B441F425-8C6E-11D2-977A-0000F87A926F}"

# SERVER
ORG_EMAIL = ""
FROM_EMAIL = "auto_alerts" + ORG_EMAIL
FROM_PWD = ""
SMTP_SERVER = "imap.gmail.com"
SMTP_PORT = 993
PORT = 465

# MY_EMAIL
IT_SUPPORT_EMAIL = ""
IT_SUPPORT_PASSWORD = ""

# ATHENS EMAIL
ATHENS_EMAIL = ""
ATHENS_PASSWORD = ""

# GROVETOWN EMAIL
GROVETOWN_EMAIL = ""
GROVETOWN_PASSWORD = ""

# Path to (Z:) drive where files are stored
PATH_TO_Z = "\\\\athnas\\Shared Data\\Packing_Slips\\Images\\"
PATH_TO_Z_LOG = "\\\\athnas\\Shared Data\\Packing_Slips\\Temp\\Delivery_Report_Log.txt"
PATH_FROM = "\\\\athnas\\Shared Data\\Packing_Slips\\ToProcess\\"
PATH_TO_TEMP_FILE = PATH_FROM + "Temp.png"

# Path to local drive where files are stored
PATH_FROM_LOCAL = "C:\\\\Temp_Augusta\\To_Process\\"
PATH_TO_LOCAL_UNVERIFIED = "C:\\\\Temp_Augusta\\UN_Process\\"

#DB Connection String
SERVER = ''
DATABASE = ''
USERNAME = ''
PASSWORD = ''


# =====================================================================================================================================
# -------------------------------------------------------------File: WindowsBalloonTip.py----------------------------------------------
# ----------------------------------------------------------------Author: dotH55-------------------------------------------------------
# =====================================================================================================================================

# Window Notification
# Window Notification Class
class WindowsBalloonTip:
    def __init__(self):
        message_map = {
                win32con.WM_DESTROY: self.OnDestroy,
        }
        # Register the Window class.
        wc = WNDCLASS()
        self.hinst = wc.hInstance = GetModuleHandle(None)
        wc.lpszClassName = "PythonTaskbar"
        wc.lpfnWndProc = message_map # could also specify a wndproc.
        self.classAtom = RegisterClass(wc)

    def ShowWindow(self,title, msg):
        # Create the Window.
        style = win32con.WS_OVERLAPPED | win32con.WS_SYSMENU
        self.hwnd = CreateWindow( self.classAtom, "Taskbar", style, \
                0, 0, win32con.CW_USEDEFAULT, win32con.CW_USEDEFAULT, \
                0, 0, self.hinst, None)
        UpdateWindow(self.hwnd)
        iconPathName = os.path.abspath(os.path.join( sys.path[0], "balloontip.ico" ))
        icon_flags = win32con.LR_LOADFROMFILE | win32con.LR_DEFAULTSIZE
        try:
           hicon = LoadImage(self.hinst, iconPathName, \
                    win32con.IMAGE_ICON, 0, 0, icon_flags)
        except:
          hicon = LoadIcon(0, win32con.IDI_APPLICATION)
        flags = NIF_ICON | NIF_MESSAGE | NIF_TIP
        nid = (self.hwnd, 0, flags, win32con.WM_USER+20, hicon, "tooltip")
        Shell_NotifyIcon(NIM_ADD, nid)
        Shell_NotifyIcon(NIM_MODIFY, \
                         (self.hwnd, 0, NIF_INFO, win32con.WM_USER+20,\
                          hicon, "Balloon  tooltip",msg,200,title))
        # self.show_balloon(title, msg)
        DestroyWindow(self.hwnd)

    def OnDestroy(self, hwnd, msg, wparam, lparam):
        # The commented code deletes the popup
        #nid = (self.hwnd, 0)
        #Shell_NotifyIcon(NIM_DELETE, nid)
        #PostQuitMessage(0) # Terminate the app.
        pass

# =====================================================================================================================================
# -------------------------------------------------------------File: Packing_Slip.py---------------------------------------------------
# ----------------------------------------------------------------Author: dotH55-------------------------------------------------------
# =====================================================================================================================================

# Window Notification
# Let user know program has started
NOTIFICATION = WindowsBalloonTip()

# try connecting to server
# Install ODBC driver if an exception occurs
try:
    CNXN = pyodbc.connect('DRIVER={ODBC Driver 13 for SQL Server};SERVER=' + SERVER +';DATABASE=' + DATABASE + ';UID=' + USERNAME + ';PWD=' + PASSWORD)
    CURSOR = CNXN.cursor()
except:
    # Window Notification
    # Inform user to install driver
    NOTIFICATION.ShowWindow("PyODBC Driver is missing", "Follow Instruction to Install Driver")
    os.system("msodbcsql.msi")

def main():

    # Start Program
    NOTIFICATION.ShowWindow("Scannen", "Processing...")

    # Check if paths are valid
    CheckPaths()

    # Collect Garbage
    if(datetime.datetime.today().day == 30):
        NOTIFICATION.ShowWindow("Scannen", "Trash Day: Scannen will take longer than usual")
        GarbageCollector()

    # List
    AthensStrList = ""
    GrovetownStrList = ""

    # Get initial # of files in PATH_FROM_LOCAL
    initialCount = len(os.listdir(PATH_FROM_LOCAL))
    processedFiles = 0

    # Iterate thru PATH_FROM_LOCAL
    for filename in os.listdir(PATH_FROM_LOCAL):
        # Verify filename
        # Send an alert if negative
        fileInfo = None
        try:
            fileInfo = GetFileInfo(PATH_FROM_LOCAL + filename)
            # fileInfo[0]: Order#
            # fileInfo[1]: Location
            # fileInfo[2]: Filename
        except:
            NOTIFICATION.ShowWindow("Scannen", "Barcode Error!")
            UnverifiedFiles(filename)
            continue

        if(not AuthentifyFile(fileInfo[2])):
            #SendMail(IT_SUPPORT_EMAIL, PATH_FROM_LOCAL + filename, TYPE = "ALERT")
            # Toast a notification & Open folder
            NOTIFICATION.ShowWindow("Scannen", "Authentification Error!")
            UnverifiedFiles(filename)
            continue
        else:
            pdfFile = PATH_TO_Z + fileInfo[2]
            image = Image.open(PATH_FROM_LOCAL + filename)
            image.save(pdfFile, "pdf", resolution = 100.0, save_all = True)
            image.close()

            # Reassign Location & add the file path to a list that will be called by a database method later
            if (fileInfo[1] == "ATHENS"):
                AthensStrList += ", " + fileInfo[0]
            else:
                GrovetownStrList += ", " + fileInfo[0]

            # Database Work
            # Get serial number & Email from the first and make an Insert into the database on the second
            serialNumAndEmail = GetSerialNumber(fileInfo[0], fileInfo[1])
            InsertPackingSlip(fileInfo[0], serialNumAndEmail[0], fileInfo[1], PATH_TO_Z + fileInfo[2])

            # Send Confirmation Email
            if(not serialNumAndEmail[1] == ""):
                try:
                    SendMail(serialNumAndEmail[1], PATH_TO_Z + fileInfo[2], TYPE = fileInfo[2])
                except:
                    NOTIFICATION.ShowWindow(fileInfo[0], "Has an Incorrect Email Address")
                #print("Sent to: " + serialNumAndEmail[1])

            # Log
            RecordLog(fileInfo[2], serialNumAndEmail[1])

            # Remove PNG File
            os.remove(PATH_FROM_LOCAL + filename)

            # Increment processedFiles
            processedFiles += 1

            

    # Del Comment at the beginning of str
    AthensStrList = AthensStrList[2:len(AthensStrList)]
    GrovetownStrList = GrovetownStrList[2:len(GrovetownStrList)]

    # Release
    ReleasePackingSlips(AthensStrList, GrovetownStrList)

    # Inform user
    NOTIFICATION.ShowWindow("Scannen", "Scanned: " + str(initialCount) + "\nProcessed: " + str(processedFiles))

    # Open UN_Processed Folders if there were some errors
    if(not(len(os.listdir(PATH_TO_LOCAL_UNVERIFIED)) == 0)):
        subprocess.call("explorer C:\\Temp_Augusta\\UN_Process\\", shell = True)
        NOTIFICATION.ShowWindow("Scannen", \
            "Using Format \"scan_Order#--Loc.pdf\"\nRename & Move Files to\n\"C:\\Temp_Augusta\\To_Process\\\"\nEx: scan_54441--Ath.pdf, scan_44475--Gro.pdf")

# def main()

# @param -> pathToFile: location in memory
# @retur -> Filename: Consists of the order# & location in the format "scan_Order#--Loc.pdf"
# This method obtains an order# & location from GetOrder() & getLocation() in order to construct
# a filename It uses RegEx.
def GetFileInfo(pathToImage):
    string = str(decode(Image.open(pathToImage)))
    orderNumber = GetOrder(string)
    location = GetLocation(string)
    name = "scan_" + orderNumber + "--" + GetLoc(string) + ".pdf"
    return [orderNumber, location, name]

# @param -> filename 
# @retur -> Order#: Uses Regex to obtain an Order#
def GetOrder(filename):
    return re.search("data=b\'(.*?)--", filename).group(1)

# @param -> filename 
# @retur -> Loc: Uses Regex to obtain an location in full
def GetLocation(filename):
    tempString = re.search("--(.*?)\', ", filename).group(1)
    if(tempString[0] == "A"):
        return "ATHENS"
    else:
        return "GROVETOWN"

# @param -> filename 
# @retur -> Loc: Uses Regex to obtain an location abbreviated 
def GetLoc(filename):
    tempString = re.search("--(.*?)\', ", filename).group(1)
    if(tempString[0] == "A"):
        return "Ath"
    else:
        return "Gro"

# @param -> ORDER_NUMBER 
# @param -> LOC
# @retur -> list[SerialNumber, Email]
# This method retrieves serial numbers from database
def GetSerialNumber(ORDER_NUMBER, LOC):
    # sql code
    sql = "DECLARE	@pSN varchar(50), @pEmailOut varchar(max) \r\n" \
        "EXEC [dbo].[usp_Insert_PackSlip_Get_SN] @pInvNumber = N'" + ORDER_NUMBER \
        + "', @pLocation = N'" + LOC + "', @pSN = @pSN OUTPUT, @pEmailOut = @pEmailOut OUTPUT \r\n" \
        "SELECT	@pSN as N'@pSN', @pEmailOut as N'@pEmailOut' \r\n"
    temp = CURSOR.execute(sql).fetchone()
    # temp[0]: Serial Number
    # temp[1]: email address
    return [temp[0],  temp[1]]

# @void
# @param -> ORDER_NUMBER 
# @param -> SERIAL_NUMBER
# @param -> LOCATION 
# @param -> NEW_FILENAME
# This method inserts File paths into database
def InsertPackingSlip(ORDER_NUMBER, SERIAL_NUMBER, LOCATION, NEW_FILENAME):
    sql = "EXEC [usp_Insert_PackSlip_Scans] @pInvNumber = N'" + ORDER_NUMBER  + "',@pSN = N'" + SERIAL_NUMBER  + "', @pNewFileName = N'" + NEW_FILENAME \
        + "', @pLocation = N'" + LOCATION + "'\r\n"
    CURSOR.execute(sql)
    CURSOR.commit()

# @void
# @param -> ATH_ORDER_NUMBER_LIST 
# @param -> GRO_ORDER_NUMBER_LIST
# This method releases packing slips into invoices
def ReleasePackingSlips(ATH_ORDER_NUMBER_LIST, GRO_ORDER_NUMBER_LIST):
    # sql code
    sql1 = "EXEC [dbo].[usp_Auto_Release_Scanned_PackSlip] @pInvNumber_List = N'" + ATH_ORDER_NUMBER_LIST + "', @pLocation = N'" + "ATHENS" + "'\r\n"
    sql2 = "EXEC [dbo].[usp_Auto_Release_Scanned_PackSlip] @pInvNumber_List = N'" + GRO_ORDER_NUMBER_LIST + "', @pLocation = N'" + "GROVETOWN" + "'\r\n"
    CURSOR.execute(sql1)
    CURSOR.commit()
    CURSOR.execute(sql2)
    CURSOR.commit()

# @param -> FILENAME 
# @retur -> Bool
# This method uses Regex to authenticate a file name
def AuthentifyFile(FILENAME):
    if(re.match(r'^scan_\d{1,}--Ath.pdf$', FILENAME) or re.match(r'^scan_\d{1,}--Gro.pdf$', FILENAME)):
        return True
    else:
        return False

# @void
# This method deletes files older than 3 years
def GarbageCollector():
    if(datetime.datetime.today().day == 30):
        for filename in os.listdir(PATH_TO_Z):
            pdfFile = PdfFileReader(open(PATH_TO_Z + filename, "rb"))
            pdfText = str(pdfFile.getDocumentInfo())
            if re.search('D:(.+?)-', pdfText):
                year = re.search('D:(.+?)-', pdfText).group(1)[0:4]
                month = re.search('D:(.+?)-', pdfText).group(1)[5:6]
                day = re.search('D:(.+?)-', pdfText).group(1)[7:8]
                if (int(datetime.date.today().year - datetime.date(int(year), int(month), int(day)).year) > 3):
                    os.remove(PATH_TO_Z + filename)

# @void
# This method logs processed work
def RecordLog(Path_, Email_):
    f = open(PATH_TO_Z_LOG, "a+")
    f.write("Report Time: " + time.ctime() + "\r\nReport Path: " + Path_ + "\r\nEmailed To: " + Email_ + "\r\n\r\n")

# @void
# This method makes sure that "C:\\\\Temp_Augusta\\To_Process\\" exits
# and creates it if not
def CheckPaths():
    try:
        os.makedirs(PATH_FROM_LOCAL)
        os.makedirs(PATH_TO_LOCAL_UNVERIFIED)
    except:
        pass

# @void
# This method moves an unverified file to PATH_TO_LOCAL_UNVERIFIED
def UnverifiedFiles(filename):
    pdfFile = PATH_TO_LOCAL_UNVERIFIED + str(random.randint(1, 10000)) + ".pdf"
    image = Image.open(PATH_FROM_LOCAL + filename)
    image.save(pdfFile, "PDF", resolution = 100.0, save_all = True)
    image.close()
    os.remove(PATH_FROM_LOCAL + filename)

# @void
# @param -> RECEIVER_EMAIL 
# @param -> PDF_PATH 
# @param -> ALERT a Bool to differenciate between an alert and confirmation email
# This method sends emails
def SendMail(RECEIVER_EMAIL, PDF_PATH, TYPE):

    # Message
    message = MIMEMultipart("alternative")
    message["To"] = RECEIVER_EMAIL
    message["From"] = "Duplicating Systems Inc."
    senderEmail = ""
    senderPassword = ""
    
    if(TYPE == "ALERT"):
        # Send an Alert
        message["Subject"] = "AUTO_ALERT"
        senderEmail = IT_SUPPORT_EMAIL
        senderPassword = IT_SUPPORT_PASSWORD
        html = """\
        <html>
        PACKING SLIP ERROR
            <body>
                    <p>This is an automated alert.<br>
                    File: {PDF_PATH}</p>
                    <p>Duplicating Systems, Inc.<br>
                    177 Newton Bridge Road<br>
                    Athens, Georgia 60607<br>
                    Office: 706-546-1220<br>
                    Fax   : 706-353-2133<br>
                </p>
            </body>
        </html>"""

    elif(TYPE == "ATHENS"):
        # Send a Confirmation 
        message["Subject"] = "Delivery Confirmation"
        senderEmail = ATHENS_EMAIL
        senderPassword = ATHENS_PASSWORD
        html = """\
        <html>
        Delivery Report
            <body>
                <p>Duplicating Systems would like to thank you for your continued support.<br>
                <br>
                Continuing to improve our service and communication, this email is a notification of shipment <br>
                to your location.<br>
                <br>
                This is for your records only.<br><br>
                If you have any questions, please contact us.<br><br>
                    This is an automated email from.<br>
                    <a href="https://duplicatingsystems.com/"> Duplicating Systems, Inc.</a></p>
                    <p>Duplicating Systems, Inc.<br>
                    177 Newton Bridge Road<br>
                    Athens, Georgia 60607<br>
                    Office: 706-546-1220<br>
                    Fax   : 706-353-2133<br>
                </p>
            </body>
        </html>
        """

    else:
        # Send a Confirmation 
        message["Subject"] = "Delivery Confirmation"
        senderEmail = GROVETOWN_EMAIL
        senderPassword = GROVETOWN_PASSWORD
        html = """\
        <html>
        Delivery Report
            <body>
                <p>Duplicating Systems would like to thank you for your continued support.<br>
                <br>
                Continuing to improve our service and communication, this email is a notification of shipment <br>
                to your location.<br>
                <br>
                This is for your records only.<br><br>
                If you have any questions, please contact us.<br><br>
                    This is an automated email from.<br>
                    <a href="https://duplicatingsystems.com/"> Duplicating Systems, Inc.</a></p>
                    <p>Duplicating Systems, Inc.<br>
                    177 Newton Bridge Road<br>
                    Athens, Georgia 60607<br>
                    Office: 706-546-1220<br>
                    Fax   : 706-353-2133<br>
                </p>
            </body>
        </html>
        """

    with open(PDF_PATH, "rb") as attachment:
        PDF_FILE = MIMEBase("application", "octet-stream")
        PDF_FILE.set_payload(attachment.read())

    # Encode file in ASCII character
    encoders.encode_base64(PDF_FILE)

    # Add header as key/value pair to attachment part
    PDF_FILE.add_header(
        "Content-Disposition", 
        f"attachment; filename = Delivery_Document.pdf",
    )
    
    message.attach(MIMEText(html, "html"))
    message.attach(PDF_FILE)
    context = ssl.create_default_context()
    with smtplib.SMTP_SSL("smtp.gmail.com", PORT, context = context) as server:
        server.login(senderEmail, senderPassword)
        server.sendmail(senderEmail, RECEIVER_EMAIL, message.as_string())

# if __name__ == "__main__":
#     try:
#         main()
#     except Exception:
#         exc_type, exc_obj, exc_tb = sys.exc_info()
#         fname = os.path.split(exc_tb.tb_frame.f_code.co_filename)[1]
#         print(exc_type, fname, exc_tb.tb_lineno)

if __name__ == "__main__":
    main()

# End of File


        