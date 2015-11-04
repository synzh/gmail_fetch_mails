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
email_account.select('"[Gmail]/All Mail"')

# Search mail by uid, in case user delete email in real time while fetching emails
# The search criteria also by date range
typ, data = email_account.uid("search",None, '(SINCE "01-Oct-2015" BEFORE "31-Oct-2015")')
uids = data[0].split()

with open('test.csv',mode='w') as f:
  f.write('')

trim_characters = ['\r','\n',',']
# BODY.PEEK for not set mails as READ when parsing mails.
for uid in data[0].split():
  typ, data = email_account.uid("fetch",uid, '(BODY.PEEK[HEADER.FIELDS (DATE FROM TO SUBJECT)])')
  # Other way to retrieve emails body
  #typ, data = email_account.fetch(ids, '(RFC822)')
  raw_email = data[0][1]
  email_message = email.message_from_bytes(raw_email)
  
  # handling the newlines and comma
  ffrom = email_message['From']
  fsubject = email_message['Subject']
  for ch in trim_characters:
    if ch in email_message['From']:
      ffrom = ffrom.replace(ch," ")
    if ch in email_message['Subject']:
      fsubject = fsubject.replace(ch," ")

  # a way to take care all different encoding prefix and change to bytes
  try:
    foundf = re.search(r'(?<=\?).*?(?=\?)', ffrom).group(0)
    print("found="+ foundf + "  Subject:"+ffrom)
    parts = email.header.decode_header(ffrom)
    rawfrom = parts[0][0].decode(foundf,'ignore').encode('utf-8','ignore')
    decodedfrom = rawfrom
  except AttributeError:
    decodedfrom = ffrom.encode('utf-8','ignore')

  # a way to take care all different encoding prefix
  try:
    foundq = re.search(r'(?<=\?).*?(?=\?)', fsubject).group(0)
    print("found="+ foundq + "  Subject:"+fsubject)
    parts = email.header.decode_header(fsubject)
    rawsubject = parts[0][0].decode(foundq,'ignore').encode('utf-8','ignore')
    decodedsubject = rawsubject.replace(b',',b'')
  except AttributeError:
    decodedsubject = fsubject.encode('utf-8','ignore').replace(b',',b'')

  # for debug use or check the output in realtime
  print ("\nDate:", email_message['Date'])
  print ("Decoded From:", decodedfrom)
  print ("Original From:", ffrom)
  print ("Decoded Subject:", decodedsubject)
  print ("Subject:", fsubject + "\n")

  # to update file content with the encoding we want in binary mode, in this case utf-8
  # write to file and handling the timezone things also the time conversion...
  with open('test.csv',mode='ab+') as f:
    maildate = email_message['Date']
    tt = parsedate_tz(maildate)
    timestamp = mktime_tz(tt)
    localtimestamp = time.strptime(formatdate(timestamp,True), '%a, %d %b %Y %H:%M:%S +0800')
    convertedtime = time.strftime("%d/%m/%Y",time.localtime(time.mktime(localtimestamp)))
    comma = str(",").encode('utf-8','ignore')
    newline = str("\n").encode('utf-8','ignore')
    f.write(str(convertedtime).encode('utf-8','ignore'))
    f.write(comma)
    f.write(decodedfrom)
    f.write(comma)
    f.write(decodedsubject)
    f.write(newline)

email_account.close()
email_account.logout()
print ("Done!")
