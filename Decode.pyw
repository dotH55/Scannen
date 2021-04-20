# File: Decode.py
# Author: dotH55
# Description: This program reads the barcode 
# located on each file then returns a string which is 
# used to rename scanned documents.

import os, re
from PIL import Image
from pyzbar.pyzbar import decode

# @param -> pathToFile: location in memory
# @retur -> Filename: Consists of the order# & location in the format "scan_Order#--Loc.pdf"
# This method obtains an order# & location from GetOrder() & getLocation() in order to construct
# a filename It uses RegEx.
def GetFilename(pathToImage):
    string = str(decode(Image.open(pathToImage)))
    return "scan_" + GetOrder(string) + "--" + GetLocation(string) + ".pdf"

# @param -> filename 
# @retur -> Order#: Uses Regex to obtain an Order#
def GetOrder(filename):
    return re.search("data=b\'(.*?)--", filename).group(1)

# @param -> filename 
# @retur -> Loc: Uses Regex to obtain an location
def GetLocation(filename):
    tempString = re.search("--(.*?)\', ", filename).group(1)
    if(tempString[0] == "A"):
        return "Ath"
    else:
        return "Gro"