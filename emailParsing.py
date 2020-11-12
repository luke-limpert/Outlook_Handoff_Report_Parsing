import win32com.client
import win32com
import os
import sys

f = open("testfile.txt","w+")

outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
accounts= win32com.client.Dispatch("Outlook.Application").Session.Accounts;

def emailleri_al(folder):
    messages = folder.Items
    a=len(messages)
    if a>0:
        for message2 in messages:
             try:
                subject = message2.Subject
                subject2 = subject.strip("")
                subject2 = subject2.lower()
                body = message2.Body
                body2 = body.strip()
                if 'Hand Off' in str(subject2):
                    print(subject, file=f)
                    print(body2, file=f)
                if 'handoff' in str(subject2):
                    print(subject, file=f)
                    print(body2, file=f)
                if 'Handoff' in str(subject2):
                    print(subject, file=f)
                    print(body2, file=f)
             except:
                print("Error")
                print(account.DeliveryStore.DisplayName)
                pass

             try:
                message2.Save
                message2.Close(0)
             except:
                 pass



for account in accounts:
    global inbox
    inbox = outlook.Folders(account.DeliveryStore.DisplayName)
    print("****Account Name**********************************",file=f)
    print(account.DisplayName,file=f)
    print(account.DisplayName)
    print("***************************************************",file=f)
    folders = inbox.Folders

    for folder in folders:
        print("****Folder Name**********************************", file=f)
        print(folder, file=f)
        print("*************************************************", file=f)
        emailleri_al(folder)
        a = len(folder.folders)

print("Finished Succesfully")
