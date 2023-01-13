import win32com.client
import  os


def mail_with_cc_attach(To,CC,Subject,Boby,attachment):
    ol=win32com.client.Dispatch("outlook.application")
    olmailitem=0x0 #size of the new email
    newmail=ol.CreateItem(olmailitem)
    newmail.Subject=Subject
    if '@' in To and '@' in CC:
        newmail.To=To
        newmail.CC=CC
        newmail.Body=Boby
        file=os.path.isfile(attachment)
        if file:
            newmail.Attachments.Add(attachment)
            newmail.Send()
            return("mail sent sucessfully")
        else:
            return("please provide correct file ")
        
    else:
        return("enter correct mail ID's")


def mail_with_attach(To,Subject,Boby,attachment):
    ol=win32com.client.Dispatch("outlook.application")
    olmailitem=0x0 #size of the new email
    newmail=ol.CreateItem(olmailitem)
    newmail.Subject= Subject
    if '@' in To :
        newmail.To=To
        newmail.Body=Boby
        file=os.path.isfile(attachment)
        if file:
            newmail.Attachments.Add(attachment)
            newmail.Send()
            return("mail sent sucessfully")
        else:
            return("please provide correct file ")
    else:
        return("enter correct mail ID's")
    

def mail_with_cc(To,CC,Subject,Boby):
    ol=win32com.client.Dispatch("outlook.application")
    olmailitem=0x0 #size of the new email
    newmail=ol.CreateItem(olmailitem)
    newmail.Subject= Subject
    if '@' in To and '@' in CC:
        newmail.To=To
        newmail.CC=CC
        newmail.Body=Boby
        newmail.Send()
        return("mail sent sucessfully")
    else:
        return("enter correct mail ID's")


def mail_without_cc_attach(To,Subject,Boby):
    ol=win32com.client.Dispatch("outlook.application")
    olmailitem=0x0 #size of the new email
    newmail=ol.CreateItem(olmailitem)
    newmail.Subject=Subject
    if '@' in To:
        newmail.To=To
        #newmail.CC=CC
        newmail.Body=Boby
        #newmail.Attachments.Add(attachment)
        newmail.Send()
        return("mail sent sucessfully")
    else:
        print("enter correct mail's")
        return("enter correct mail ID's")



if __name__ == "__main__":
    mail_without_cc_attach()
    mail_with_attach()
    mail_with_cc()
    mail_with_cc_attach()