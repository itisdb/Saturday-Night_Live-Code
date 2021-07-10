import smtplib
from email.message import EmailMessage
import pandas as pd
import codecs

df = pd.read_excel('/content/drive/MyDrive/Colab Notebooks/emails.xlsx') # can also index sheet by name or fetch all sheets
mylist = df['Emails'].tolist()

#taking the email as the input
emailid=input("Enter the Email ID: ")
emailpasswd=input("Enter the Email Password: ")

#taking the subject as the input
subjectemail=input("Enter the Subject of the email: ")

msg = EmailMessage()
msg['Subject']=subjectemail
msg['from'] = emailid
msg['to'] = mylist

#enter the html here but make sure it is in one line
path="/content/drive/MyDrive/Colab Notebooks/email.html" 
file=codecs.open(path,"r")
htmlcode=file.read()
msg.add_alternative(htmlcode,subtype='html')


with smtplib.SMTP_SSL('smtp.gmail.com',465) as smtp:
    smtp.login(emailid,emailpasswd)
    smtp.send_message(msg)

print("messgae sent successfully")
