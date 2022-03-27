'''
Tickets with qr codes generating system
It will create qr codes for all the members in the excel file.
It takes a design of the tickets and then overlays the qr code on it.
'''
from ctypes.wintypes import RGB
import qrcode
from PIL import Image
from PIL import ImageFont
from PIL import ImageDraw
import openpyxl
#FIXME handling of data redundancy just use google sheets for this
sheet_name="sheet.xlsx"
wb = openpyxl.load_workbook(sheet_name)
audience = wb['Sheet1']
idCol=audience['A']
nameCol=audience['B']
emailCol=audience['C']
phoneCol=audience['D']
universityCol=audience['E']
facultyCol=audience['F']
yearOfGraduationCol=audience['H']
lenOfCol=len(nameCol)


qr = qrcode.QRCode(
                version=1,
                error_correction=qrcode.constants.ERROR_CORRECT_L,
                #TODO may be change the size of the qr code
                box_size=13,
                border=2)



def createQrs():
   for i in range(0,lenOfCol):
        qr.clear()
        print(idCol[i].value)
        qr.add_data(str(nameCol[i].value)+' & '+str(phoneCol[i].value)+' & '+str(emailCol[i].value)+' & '+str(universityCol[i].value)+' & '+str(facultyCol[i].value)+' & '+str(yearOfGraduationCol[i].value))
        qr.make_image()
        img = qr.make_image()
        #name of ticket design
        img2=Image.open(r"ticketDesign.png")
        img2.load()
        #TODO x,y on which the qr will be positioned
        img2.paste(img,(1912,1508)) 
        img2.save("qrs/"+str(idCol[i].value)+".png")
        #clear the qr code
        qr.clear()
       
createQrs()

