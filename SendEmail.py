import win32com.client
import datetime

def sendmail(today, finalresult2):
    outlook = win32com.client.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)

    mail.To = "email"
    #mail.To = ""
    mail.Subject = "Inconsistencies discovered"
    mail.body = "Please note the following Inconsistencies detailed in the attachment."
    mail.Attachments.Add(finalresult2)
    #mail.CC =
    mail.Send()
