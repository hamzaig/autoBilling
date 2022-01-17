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
    
def consumerAmount(barcode):
    if len(barcode) == 16:
        return int(barcode[8:15])
    elif len(barcode) == 30:
        lescoConsumer = lescoAmountCalc(barcode) 
        return int(lescoConsumer)
    elif len(barcode) == 50:
        return int(barcode[36:42])

def autoBilling(consumer,amount):
    # time.sleep(5)
    # print("Hello")
    bill = consumer
    billAmount = amount
    # print(type(bill))
    # print(type(billAmount))
    # print(int(bill))
    # print(int(billAmount))
    ### time.sleep(5)
    keyboardController.press(Key.ctrl)
    keyboardController.press("a")
    keyboardController.release('a')
    keyboardController.release(Key.ctrl)
    keyboardController.press(Key.ctrl)
    keyboardController.press("a")
    keyboardController.release('a')
    keyboardController.release(Key.ctrl)
    keyboardController.type(bill)
    keyboardController.press(Key.tab)
    keyboardController.release(Key.tab)
    keyboardController.press(Key.tab)
    keyboardController.release(Key.tab)
    keyboardController.type('u')
    keyboardController.press(Key.tab)
    keyboardController.release(Key.tab)
    keyboardController.type('03327229422')
    keyboardController.press(Key.tab)
    keyboardController.release(Key.tab)
    keyboardController.press(Key.space)
    keyboardController.release(Key.space)
    keyboardController.press(Key.tab)
    keyboardController.release(Key.tab)
    keyboardController.press(Key.tab)
    keyboardController.release(Key.tab)
    keyboardController.press(Key.space)
    keyboardController.release(Key.space)
    while True:
    #     ### time.sleep(2)
        try:
            myScreenshot = pyautogui.screenshot()
            myScreenshot.save('amount.png')
            nadraAmount = Image.open('amount.png')
            nadraAmountCrop = nadraAmount.crop((480, 246, 650, 288))
        ###     nadraAmountCrop.show()
            nadraAmountCrop.save("amount.png", 'JPEG')
            result = pytesseract.image_to_string(nadraAmountCrop)  
            ### print(result)
            nadraAmountInt = int(float(result.replace(',',"")))
            if nadraAmountInt == billAmount:
                ahk.mouse_move(1030,200,speed=1)
                mouse.press(Button.left)
                mouse.release(Button.left)
                ahk.mouse_move(1100,100,speed=1)
                mouse.press(Button.left)
                mouse.release(Button.left)
                ahk.mouse_move(830,300,speed=1)
                mouse.press(Button.left)
                mouse.release(Button.left)
                print(bill,"=>",billAmount)
                break
            else:
                print("Bill Not Paid")
        except:
            newMyScreenshot = pyautogui.screenshot()
            newMyScreenshot.save('noBill.png')
            noBill = Image.open('noBill.png')
            noBillCrop = noBill.crop((410, 350, 595, 400))
            ### nadraAmountCrop.show()
            newResult = pytesseract.image_to_string(noBillCrop)
            if(newResult[0] == "n"):
                mouse.position = (1100, 100)
                mouse.press(Button.left)
                mouse.release(Button.left)
                mouse.position = (830, 300)
                mouse.press(Button.left)
                mouse.release(Button.left)
                mouse.press(Button.left)
                mouse.release(Button.left)
                mouse.press(Button.left)
                mouse.release(Button.left)
                print(bill,"=>",billAmount,"Not Be Paid")
                break
            else:
                continue


while True:
    barcodeResponse = inputBarcode()
    if not barcodeResponse:
        consumer = consumberNumber(barcodeString);
        amount = consumerAmount(barcodeString);
        autoBilling(consumer,amount)
        barcodeString = ""
        # # autoBilling()
        # # print(barcodeString)
        # barCodeStr = str(barcodeString)
        # # print(len(barCodeStr))
        # print(barcodeString)
        # print(consumberNumber(barcodeString))

       
        

# print(barcodeResponse)
# 


# ### Lesco = ''
# ### Lwasa = ''
# ### Sngpl = ''
# ### if len(barcodeString) == 16:
# ###     Lwasa = x
# ### elif len(x) == 30:
# ###     print(x)
# ### elif len(x) == 50:
# ###     print(x)




