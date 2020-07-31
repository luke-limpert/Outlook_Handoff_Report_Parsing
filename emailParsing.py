import win32com.client
import win32com
import os
from os import path
import sys

f = open("testfile.txt","w+")

outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
inbox = outlook.GetDefaultFolder(6)

a = f.read()

messages = inbox.Items
HandoffMessages = messages.find("Handoff")

for HandoffMessages in a:
    print HandoffMessages.ReceivedTime
    print HandoffMessages.Subject.encode('ascii', 'ignore')
    print HandoffMessages.Body.encode('ascii', 'ignore')

f.close()
