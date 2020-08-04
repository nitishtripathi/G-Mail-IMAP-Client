import sys
import imaplib
import getpass
import email
import email.header
import datetime
import csv 
import logging
import re

print("Welcome to Gmail-Scraper\nPlease go to this link"
      "\nhttps://www.google.com/settings/security/lesssecureapps\nand"
      " disable the 'Allow less secure apps' \n"
      "option for the Email ID you wanna scrape\n\nPlease wait for 'Done'"
      " to indicate end of process \nThank you.")


EMAIL_ACCOUNT = input("Email: ")
PASSWORD = getpass.getpass()
EMAIL_FOLDER = "INBOX"
scrapedemail = []

def getEmails(str):
    str = str.replace(".png@",'')
    str = str.replace(".jpg@",'')
    regex = r'([\w0-9._-]+@[\w0-9._-]+\.[\w0-9_-]+)'
    newstr = re.findall(regex, str, re.M|re.I)
    newstr = list(set(newstr))
    newstr = ', '.join(newstr)
    return newstr

def phoneNum(str):
    phnum = []
    for i in re.findall(r'[\+\(][1-9][0-9 .\-\(\)]{8,}[0-9]', str):
        phnum.append(i)
    phnum = list(set(phnum))
    phnum = ', '.join(phnum)
    return phnum

def writeIn(Item):
    file1 = open('files.csv','a+', newline='', encoding='utf-8')
    items=[]
    items.append(Item)
    keys = items[0].keys()
    dict_writer=csv.DictWriter(file1, keys)
    with file1:
        dict_writer.writerows(items)
    return 0


def process_mailbox(M):
    rv, data = M.search(None, "ALL")
    
    if rv != 'OK':
        print("No messages found!")
        return
    counter = 0
    data = data[0].split()
    data = [x for x in data if x not in scrapedemail]
    for num in data:
        scrapedemail.append(num)
        rv, data = M.fetch(num, '(RFC822)')
        if rv != 'OK':
            print("ERROR getting message", num)
            return
        
        counter = counter + 1
        print(counter)
        msg = email.message_from_bytes(data[0][1])
        subject = ""
        emailids = ""
        phonenumbers = ""
        if(msg['Subject']!=None):
            hdr = email.header.make_header(email.header.decode_header(msg['Subject']))
            subject = str(hdr)
        dc = {}
        for part in msg.walk():
            if part.get_content_type() == 'text/plain':
                body = part.get_payload(decode=True)
                body = str(body)
                emailids = getEmails(body)
                phonenumbers = phoneNum(body)
        dc['From'] = msg['From']
        dc['Subject'] = subject
        dc['E - mail'] = emailids
        dc['Phone'] = phonenumbers
        dc['Raw Date'] = msg['Date']
        writeIn(dc)
    print("Done")


def login():
    M = imaplib.IMAP4_SSL('imap.gmail.com', 993)

    try:
        rv, data = M.login(EMAIL_ACCOUNT, PASSWORD)
    except imaplib.IMAP4.error:
        print ("LOGIN FAILED!!! ")
        sys.exit(1)

    assert rv == 'OK', 'login failed'

    rv, mailboxes = M.list()
    if rv == 'OK':
        print("Login Succesful!")

    rv, data = M.select(EMAIL_FOLDER)

    if rv == 'OK':
        print("Processing mailbox...\n")
        try:
            process_mailbox(M)
        except Exception as e:
            print(e)
            login()
            
        M.close()
    else:
        print("ERROR: Unable to open mailbox ", rv)


    M.logout()

login()
