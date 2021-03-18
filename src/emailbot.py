# -*- coding: utf-8 -*-
"""
Created on Wed Dec 30 17:37:47 2020

@author: sayers
"""

import os
import datetime as dt
import win32com.client as win32com
from admin import read_json
from emaildata import refresh_lists
outlook = win32com.Dispatch("Outlook.Application").GetNamespace("MAPI")

#TODO set up replacement function (dictionary call) for off-names (Greg, Linda)


# setup range for outlook to search emails (so we don't go through the entire inbox)
#lastWeekDateTime = dt.datetime.now() - dt.timedelta(days = 29)
#lastWeekDateTime = lastWeekDateTime.strftime('%m/%d/%Y %H:%M %p')  #<-- This format compatible with "Restrict"
def saveAttachments(email:object):
        for attachedFile in email.Attachments: #iterate over the attachments
                try:
                        filename = attachedFile.FileName
                        attachedFile.SaveAsFile("s:\\desktop\\Jan_Return\\"+filename) #Filepath must exist already
                except Exception as e:
                        print(e)


def movemessage(message,outfolder):
    message.move(outfolder)

def movemail(emplname,outfolder,infolder=None):
    emplname=emplname.lower()
    if infolder:
        infolder=infolder
    else:
        infolder=outlook.GetDefaultFolder(6)
    messages=infolder.items
    for message in [m for m in messages if m.Class==43]:
        if emplname in message.Sender.Name.lower():
            movemessage(message,outfolder)
            
def massmove(outfolder,infolder=None,group=None,subjstr=None):
    timeframe = dt.datetime.now() - dt.timedelta(days = 7)
    if infolder:
        infolder=infolder
    else:
        infolder=outlook.GetDefaultFolder(6)
    messages= infolder.items
    messages.Sort("[ReceivedTime]", True)
    messages=messages.Restrict("[ReceivedTime] >= '" +timeframe.strftime('%m/%d/%Y %H:%M %p')+"'")
    
    if group:   #this handles the chair problem but this function is not as good as it could be
        for message in [m for m in messages if m.Class==43]:
            if message.Sender.Name in group:
                movemessage(message,outfolder)
    if subjstr:
        for message in [m for m in messages if m.Class==43]:
            if subjstr in message.subject:
                movemessage(message,outfolder)
    #please add error handling and conditional for group and subjetstr
def subfoldermove(folderstr,folderstr2,listobj):
    infolder=outlook.GetDefaultFolder(6)
    outfolder=infolder.Folders(folderstr).Folders(folderstr2)
    print(f"now moving items into {folderstr2}")
    massmove(outfolder,infolder=infolder,group=listobj)
def general_move():
    recipients=read_json('Y://Program Data//emaildata.json')
    for k,v in recipients.items():
        subfoldermove("Internal Communication",k,v)


def restrictmail(infolder,subject=None,senton=None,receivedon=None,senders=None,title=None,outfolder=None,specific=None):
    messages=infolder.items
    if subject:
        messages=messages.restrict(f'[subject] = "{subject}"')
    if senton:
        messages=messages.restrict(f"[SentOn] > '{senton}'")
    if receivedon:
        messages=messages.restrict("[ReceivedTime] >= '" + receivedon +"'")
    if senders:
        messages=messages.restrict(f"[Sender.Name] = '{senders}")
    if outfolder:
        for message in messages:
            movemessage(message,outfolder)
        print(f"Messages moved to {outfolder}")
        return('')
    #if specific:
     #   return([getattr(message,specific) for message in messages])
    if specific:
        return([getattr(message,specific) if specific in dir(message) else message for message in messages])
    return(messages)

       
def letter_download(date1,title):
    x=restrictmail(outlook.GetDefaultFolder(6),subject=title,senton=date1)
    for m in x:
        saveAttachments(m)
