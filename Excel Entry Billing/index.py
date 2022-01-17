from pynput.mouse import Button, Controller
from pynput.keyboard import Key, Controller as KeyBoardController
from pynput import keyboard
import pyttsx3 
from PIL import Image 
import pytesseract  
import pyautogui
from PIL import Image
import xlrd
import pandas
from ahk import AHK
import time
import xlwt
from xlwt import Workbook
pytesseract.pytesseract.tesseract_cmd = r'C:\\Program Files\\Tesseract-OCR\\tesseract' 
mouse = Controller()
keyboardController = KeyBoardController()
ahk = AHK()

barcodeString = ""

def charToAscii(digit):
        ch = ord(digit)
        if ch >= 65 and ch <= 90:
            return ch - 55
        elif ch >= 97 and ch <= 122:
            return ch - 61
        elif ch >= 48 and ch <= 57:
            return ch - 48
        elif ch == 91:
            return ch - 29
        elif ch == 93:
            return ch - 30

def lescoConsumerCalc(barcode):
    r1 = 4398046511104 * charToAscii(barcode[11])
    r2 = 68719476736 * charToAscii(barcode[12])
    r3 = 1073741824 * charToAscii(barcode[13])
    r4 = 16777216 * charToAscii(barcode[14])
    r5 = 262144 * charToAscii(barcode[15])
    r6 = 4096 * charToAscii(barcode[16])
    r7 = 64 * charToAscii(barcode[17])
    r8 = charToAscii(barcode[18])
    result =r1 +r2 +r3 +r4 +r5 +r6 +r7 +r8
    result = str(result)
    if len(result) == 14:
        return result + "u"
    elif len(result) == 13:
        return "0" + result + "u"
    elif len(result) == 15:
        return result[1:15] + "r"
 
def lescoAmountCalc(barcode):
    amount = barcode[1:6]
    r1 = 16777216 * charToAscii(amount[0])
    r2 = 262144 * charToAscii(amount[1])
    r3 = 4096 * charToAscii(amount[2])
    r4 = 64 * charToAscii(amount[3])
    r5 = charToAscii(amount[4])
    result = r1 + r2 + r3 + r4 + r5
    return result

def lescoDateCalc(barcode):
    date = barcode[26:29]
    day =  str(charToAscii(date[0]))
    month = str(charToAscii(date[1]))
    year = str(charToAscii(date[2]))
    finalDate = day + month+ year;
    return finalDate


def inputBarcode():
    def on_press(key):
        try:
            global barcodeString
            barcodeString = barcodeString + str(key.char)
        except AttributeError:
            pass
    def on_release(key):
        if len(barcodeString) == 16:
            if barcodeString[15] == "N":
                return False
        if len(barcodeString) == 30:
            if barcodeString[-1] == "#":
                return False
        if len(barcodeString) == 43 or len(barcodeString) == 41:
            if barcodeString[2] == "-":
                return False
        if len(barcodeString) == 50:
            return False
    with keyboard.Listener(
            on_press=on_press,
            on_release=on_release) as listener:
        listener.join()

def consumberNumber(barcode):
    # print(barcode)
    if len(barcode) == 16:
        return barcode[0:8]
    elif len(barcode) == 30:
        lescoConsumer = lescoConsumerCalc(barcode) 
        return lescoConsumer
    elif len(barcode) == 50:
        return barcode[5:16]
    elif len(barcode) == 43 or len(barcode) == 41:
        x = barcode.split("-")
        return x[1][1:14]

def consumerCompany(barcode):
    # print(barcode)
    if len(barcode) == 16:
        return "LWASA"
    elif len(barcode) == 30:
        return "LESCO"
    elif len(barcode) == 50:
        return "SNGPL"
    elif len(barcode) == 43 or len(barcode) == 41:
        return "PTCL"

def consumerDate(barcode):
    # print(barcode)
    if len(barcode) == 16:
        return "NotAvailable"
    elif len(barcode) == 30:
        lescoDate = lescoDateCalc(barcode) 
        return lescoDate
    elif len(barcode) == 50:
        return barcode[26:32];
    elif len(barcode) == 41 or len(barcode) == 43:
        x = barcode.split("-")
        return x[3]
    
def consumerAmount(barcode):
    if len(barcode) == 16:
        return int(barcode[8:15])
    elif len(barcode) == 30:
        lescoConsumer = lescoAmountCalc(barcode) 
        return int(lescoConsumer)
    elif len(barcode) == 50:
        return int(barcode[36:42])
    elif len(barcode) == 43 or len(barcode) == 41:
        x = barcode.split("-")
        return int(x[4])

def autoBilling(consumer,amount,company,date,i):
    print(i)
    sheet1.write(i, 0, consumer)
    sheet1.write(i, 1, amount)
    sheet1.write(i, 2, company)
    sheet1.write(i, 3, date)    
    wb.save('bills.xls')
    # bill = consumer
    # billAmount = amount
    # time.sleep(0.5)
    # keyboardController.press(Key.up)
    # keyboardController.type(bill)
    # keyboardController.press(Key.right)
    # keyboardController.type(str(billAmount))
    # keyboardController.press(Key.right)
    # keyboardController.type(str(company))
    # keyboardController.press(Key.right)
    # keyboardController.type(str(date))
    # keyboardController.press(Key.down)
    # keyboardController.press(Key.home)
    


wb = Workbook()
# add_sheet is used to create sheet.
sheet1 = wb.add_sheet('Sheet 1')
i = 0
while True:
    barcodeResponse = inputBarcode()
    if not barcodeResponse:
        consumer = consumberNumber(barcodeString);
        amount = consumerAmount(barcodeString);
        company = consumerCompany(barcodeString);
        date = consumerDate(barcodeString);
        i = i + 1
        autoBilling(consumer,amount,company,date,i)
        barcodeString = ""
        print("----------------------------------")