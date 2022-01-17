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

bills = pandas.read_excel('bill.xlsx', engine='openpyxl')
print(bills.head)

def total_balance():
    myScreenshot = pyautogui.screenshot()
    myScreenshot.save('totalBalance.png')
    nadraAmount = Image.open('totalBalance.png')
    nadraAmountCrop = nadraAmount.crop((480, 246, 650, 288))
    nadraAmountCrop.show()
    nadraAmountCrop.save("amount.png", 'JPEG')
    result = pytesseract.image_to_string(nadraAmountCrop)  
    print(result)
    nadraAmountInt = int(float(result.replace(',',"")))

for row in bills.iterrows():
    bill = row[1][0]
    billAmount = row[1][1]
    # ahk.mouse_position(866,617)
    # ahk.mouse_move(900,700,speed=15)
    # time.sleep(5)
    ahk = AHK()
    ahk.send(bill);
    ahk.key_press("tab")
    ahk.key_press("tab")
    ahk.key_press("u")
    ahk.key_press("tab")
    ahk.send("03327229422");
    ahk.key_press("tab")
    ahk.key_press("space")
    ahk.key_press("tab")
    ahk.key_press("tab")
    ahk.key_press("space")
    while True:
        # time.sleep(2)
        try:
            myScreenshot = pyautogui.screenshot()
            myScreenshot.save('amount.png')
            nadraAmount = Image.open('amount.png')
            nadraAmountCrop = nadraAmount.crop((480, 246, 650, 288))
        #     nadraAmountCrop.show()
            nadraAmountCrop.save("amount.png", 'JPEG')
            result = pytesseract.image_to_string(nadraAmountCrop)  
            # print(result)
            nadraAmountInt = int(float(result.replace(',',"")))
            if nadraAmountInt == billAmount:
                ahk.mouse_move(1030,200,speed=1)
                ahk.click()
                ahk.click()
                ahk.mouse_move(1100,100,speed=1)
                ahk.click()
                ahk.mouse_move(830,300,speed=5)
                ahk.click()
                ahk.click()
                ahk.click()
                print(bill,"=>",billAmount)
                break
            else:
                print("Bill Not Paid")
        except:
            newMyScreenshot = pyautogui.screenshot()
            newMyScreenshot.save('noBill.png')
            noBill = Image.open('noBill.png')
            noBillCrop = noBill.crop((410, 350, 595, 400))
            # nadraAmountCrop.show()
            newResult = pytesseract.image_to_string(noBillCrop)
            if(newResult[0] == "n"):
                ahk.mouse_move(1100,100,speed=1)
                ahk.click()
                ahk.mouse_move(830,300,speed=5)
                ahk.click()
                ahk.click()
                ahk.click()
                print(bill,"=>",billAmount,"Not Be Paid")
                break
            else:
                continue

