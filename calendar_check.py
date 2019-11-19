# -*- coding: utf-8 -*-
#! python3
"""
Spyder Editor

This is a temporary script file.
"""

import win32com.client
import datetime
my_date = datetime.date.today()
selection=input("Which floor do you want? {'8' or '9'}: ")

if selection==8:
    namelist=["Room.804@luxresearchinc.com","Room.805@luxresearchinc.com","BOS.Boardroom80inchtv@luxresearchinc.com","Room.807A@luxresearchinc.com","Room.807B@luxresearchinc.com","Room.815@luxresearchinc.com","Room.817@luxresearchinc.com","Room.822@luxresearchinc.com"]
elif selection==9:
    namelist=["Room.901@luxresearchinc.com","Room.902@luxresearchinc.com","Room.903@luxresearchinc.com","Room.904@luxresearchinc.com","Room.905@luxresearchinc.com"]
    

#zeros are free, one is busy
#evaluates the hours in the future starting from the morning 
roomdict={}

obj_outlook = win32com.client.Dispatch('Outlook.Application')
obj_Namespace = obj_outlook.GetNamespace("MAPI")
room_total=len(namelist)

for name in namelist:
    key=name[0:8]
    obj_Recipient = obj_Namespace.CreateRecipient(name)
    str_Free_Busy_Data = obj_Recipient.FreeBusy(my_date, 30)
    #print(str_Free_Busy_Data)
    bipple=datetime.datetime.now().time()
    hour=bipple.hour
    minute=bipple.minute
    #print(hour)
    #print(minute)
    if minute<30:
        n=(hour*2)
    else:
        n=(hour*2)+1
    #print(n)
    busystatus=str(str_Free_Busy_Data[n:n+3])
    if busystatus=="000":
        roomdict[key]="Wide Open"
    elif busystatus=="001":
        roomdict[key]="Open for <1 hour"
    elif busystatus=="010":
        roomdict[key]="Open for <30 minutes"
    elif busystatus=="011":
        roomdict[key]="Open for <30 minutes"
    elif busystatus=="100":
        roomdict[key]="Open Soon"
    else:
        roomdict[key]="Busy"
    

print "{:<10} {:<20}".format('Room','Status')
for v in roomdict:
    label=v
    status=roomdict[v]
    print "{:<10} {:<20}".format(label,status)