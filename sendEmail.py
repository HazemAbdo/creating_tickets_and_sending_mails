'''
system that sends emails to a list of members from excel sheet
each email is customized with a member's information and a ticket with a unique QR code
'''
# pip install openpyxl     pip install secure-smtplib     pip install email  pip install sockets
from email.message import EmailMessage
import smtplib
import openpyxl
import socket

#configuration for email server
socket.getaddrinfo('localhost', 8080)
#TODO change the email and password to your own
your_email = ""
your_password = ""
# establishing connection with gmail
#TODO Don't change the above line
server = smtplib.SMTP_SSL('smtp.gmail.com', 465)
server.ehlo()
server.login(your_email, your_password)



# reading the spreadsheet
wb = openpyxl.load_workbook("sheet.xlsx")
#audience is the first sheet in the excel file
audience = wb['Sheet1']
idCol=audience['A']
nameCol=audience['B']
emailCol=audience['C']
lenOfCol=len(nameCol)



# iterate through the records
for i in range(0,lenOfCol):
    message=EmailMessage()
    #TODO change subject and messaage content(ASK Hala and Noran)
    message['Subject'] = 'Macathon 3.0 registration ticket' 
    message['From'] = your_email
    message['To'] = emailCol[i].value
    name = nameCol[i].value
    email = emailCol[i].value
    message.set_content(f"Dear {name},\n\nHope you are doing well.")
    message.add_attachment(open("qrs/"+str(idCol[i].value)+".png",'rb').read(), maintype='image', subtype='png', filename=str(idCol[i].value)+".png")
    server.send_message(message)

   

# close the smtp server
server.close()
