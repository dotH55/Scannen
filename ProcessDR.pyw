# File: ProcessDR.py
# Author: dotH55
# Description: This python program would authentify a PDF file using its name. The File name 
# should be formated as scan_"Order#"--"Location"_0001.pdf. Example: scan_43160--Gro.pdf.
# After a successful authentification, the program would get a serial number & email address
# from the server the server. It would then Insert the packing slip into the company database. 
# The final database work would be to release Packing slips into invoices. The program also 
# emails a comfirmation email to the customer. It also has a method that deletes files 
# older than 3 years. Processed files are first stored locally before being moved to the external Z drive.
# A window notification appears when the program starts. In case of an incorrect file name,
# A window notication would appear. The program also opens the folder where the file is located.

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

# Import Classes
from GlobalVariables import *
import WindowsBalloonTip

# Window Notification
# Let user know program has started
NOTIFICATION = WindowsBalloonTip.WindowsBalloonTip()

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

# Start of main
NOTIFICATION.ShowWindow("Scannen", "Processing...")


def main():

    # Collect Garbage
    GarbageCollector()

    # List
    AthensStrList = ""
    GrovetownStrList = ""

    # iterate thru PATH_FROM_LOCAL
    for filename in os.listdir(PATH_FROM_LOCAL):
        # Verify filename
        # Send an alert if negative
        if(not AuthentifyFile(filename)):
            #SendMail(IT_SUPPORT_EMAIL, PATH_FROM_LOCAL + filename, TYPE = "ALERT")
            # Toast a notification & Open folder
            NOTIFICATION.ShowWindow("Delivery Report", "Error: " + PATH_FROM_LOCAL + filename)
            subprocess.call("explorer " +  PATH_FROM_LOCAL, shell=True)
            continue
        else:
            orderNumber = GetOrderNumber(filename)
            orderLocation = GetOrderLocation(filename)
            
            # Reassign Location & add the file path to a list that will be called for a database method later
            if (orderLocation.upper() == 'ATH'):
                AthensStrList += ", " + orderNumber
                orderLocation = 'ATHENS'
            else:
                GrovetownStrList += ", " + orderNumber
                orderLocation = 'GROVETOWN'

            # Database Work
            # Get serial number & Email from the first and make an Insert into the database on the second
            serialNumAndEmail = GetSerialNumber(orderNumber, orderLocation)
            InsertPackingSlip(orderNumber, serialNumAndEmail[0], orderLocation, PATH_TO_Z + filename)
            #print("\nOrder: " + orderNumber + "\nSerial Number: " + serialNumAndEmail[0])

            # Send Confirmation Email
            if(not serialNumAndEmail[1] == ""):
                SendMail(serialNumAndEmail[1], PATH_FROM_LOCAL + filename, TYPE = orderLocation)
                #print("Sent To: " + serialNumAndEmail[1])
            
            # Move Files to Images (PATH_TO_Z)
            shutil.move(PATH_FROM_LOCAL + filename, PATH_TO_Z + filename)

    # Del Comment at the beginning of str
    AthensStrList = AthensStrList[2:len(AthensStrList)]
    GrovetownStrList = GrovetownStrList[2:len(GrovetownStrList)]

    # Release
    ReleasePackingSlips(AthensStrList, GrovetownStrList)

# def run()

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
# @retur -> OrderNumber
# This method gets order number by Regex Substring"ING"
def GetOrderNumber(FILENAME):
    return re.search("scan_(.*?)--", FILENAME).group(1)

# @param -> FILENAME 
# @retur -> Location
# This method gets order location by Regex Substring"ING"
def GetOrderLocation(FILENAME):
    return re.search("--(.*?).pdf", FILENAME).group(1)

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

if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        print(e.args)
        print(e.__cause__)

# End of File