#!/usr/bin/env python

import imaplib
import email
import email.header
import re
from email.utils import parsedate_tz, mktime_tz, formatdate
import time, datetime

username = 'YOURGMAIL@GMAIL.COM'
password = 'YOURGMAILPASSWORD'
url = "imap.gmail.com"
email_account = imaplib.IMAP4_SSL(url)
email_account.login(username, password)

#List all your gmail folders, for debug use
for folder in email_account.list()[1]:
   print(folder.decode('utf-8'))
   
# This is the gmail trick...
email_account.select('"[Gmail]/Sent Mail"')

# Search mail by uid, in case user delete email in real time while fetching emails
# The search criteria also by date range
typ, data = email_account.uid("search",None, '(SINCE "1-Oct-2015" BEFORE "31-Oct-2015")')
uids = data[0].split()

with open('test.csv',mode='w') as f:
  f.write('')

# BODY.PEEK for not set mails as READ when parsing mails.
for uid in data[0].split():
  typ, data = email_account.uid("fetch",uid, '(BODY.PEEK[HEADER.FIELDS (DATE FROM TO SUBJECT)])')
  # Other way to retrieve emails body
  #typ, data = email_account.fetch(ids, '(RFC822)')
  raw_email = data[0][1]
  email_message = email.message_from_bytes(raw_email)
  
  # handling the newlines 
  ffrom = email_message['From'].replace('\n', ' ').replace('\r', ' ')
  fsubject = email_message['Subject'].replace('\n', ' ').replace('\r', ' ').replace(',',' ')

  # a way to take care all different encoding prefix
  try:
    foundf = re.search(r'(?<=\?).*?(?=\?)', ffrom).group(0)
    parts = email.header.decode_header(ffrom)
    rawfrom = str(parts[0][0].decode(foundf,'ignore'))
    decodedfrom = rawfrom
  except AttributeError:
    decodedfrom = ffrom

  try:
    foundq = re.search(r'(?<=\?).*?(?=\?)', fsubject).group(0)
    parts = email.header.decode_header(fsubject)
    rawsubject = str(parts[0][0].decode(foundq,'ignore'))
    decodedsubject = rawsubject
  except AttributeError:
    decodedsubject = fsubject

  # for debug use or check the output in realtime
  print ("\nDate:", email_message['Date'])
  print ("From:", ffrom)
  print ("Subject:", fsubject + "\n")
 
  # write to file and handling the timezone things also the time conversion...
  with open('test.csv',mode='a+') as f:
    maildate = email_message['Date']
    tt = parsedate_tz(maildate)
    timestamp = mktime_tz(tt)
    localtimestamp = time.strptime(formatdate(timestamp,True), '%a, %d %b %Y %H:%M:%S +0800')
    convertedtime = time.strftime("%d/%m/%Y",time.localtime(time.mktime(localtimestamp)))
    f.write(convertedtime + ',')
    f.write(repr(decodedfrom).replace('\'', '') + ',')
    f.write(repr(decodedsubject).replace('\'', '').replace('\\xad','') + "\n")
 
email_account.close()
email_account.logout()
print ("Done!")
