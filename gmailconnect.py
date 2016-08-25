import imaplib
import email
import datetime
import sys
import re
import email
from pprint import pprint
from password import *

######################################################### FUNCTIONS

def num_to_month(num):
    if num == 1:
        return "Jan"
    elif num == 2:
        return "Feb"
    elif num == 3:
        return "Mar"
    elif num == 4:
        return "Apr"
    elif num == 5:
        return "May"
    elif num == 6:
        return "Jun"
    elif num == 7:
        return "Jul"
    elif num == 8:
        return "Aug"
    elif num == 9:
        return "Sep"
    elif num == 10:
        return "Oct"
    elif num == 11:
        return "Nov"
    elif num == 12:
        return "Dec"    

def datetime_to_format(date):
    year = ''
    month = ''
    day = ''
    newdate = ''
    for i in range(len(date)):
        if i < 4:
            year += date[i]
        elif i > 4 and i < 7:
            month += date[i]
        elif i > 7:
            day += date[i]
    newdate += day
    newdate += "-"
    newdate += num_to_month(int(month))
    newdate += "-"
    newdate += year
    return newdate

def extract_body(payload):
    if isinstance(payload,str):
        return payload
    else:
        return '\n'.join([extract_body(part.get_payload()) for part in payload])

def parse_list_response(line):
    flags, delimiter, mailbox_name = list_response_pattern.match(line).groups()
    mailbox_name = mailbox_name.strip('"')
    return (flags, delimiter, mailbox_name)

######################################################### VARIABLES

######################################################### CONNECT

conn = imaplib.IMAP4_SSL("imap.gmail.com", 993)
conn.login(username, password)

list_response_pattern = re.compile(r'\((?P<flags>.*?)\) "(?P<delimiter>.*)" (?P<name>.*)')

######################################################### SEARCH

'''
typ, data = conn.search(None, 'UNSEEN')
try:
    for num in data[0].split():
        typ, msg_data = conn.fetch(num, '(RFC822)')
        for response_part in msg_data:
            if isinstance(response_part, tuple):
                msg = email.message_from_string(response_part[1])
                subject=msg['subject']                   
                print(subject)
                payload=msg.get_payload()
                body=extract_body(payload)
                print(body)
#        typ, response = conn.store(num, '+FLAGS', r'(\Seen)')

'''

today = datetime.date.today()
yesterday = today.replace(day=today.day-1)
print today
print yesterday

tostr = today.strftime('%Y/%m/%d')
yestr = yesterday.strftime('%Y/%m/%d')
search = '(SINCE \"' + datetime_to_format(yestr) + '\" BEFORE \"' + datetime_to_format(tostr) + '\")'

print search

try:
    #Connect to the All Mail inbox
    conn.select("[Gmail]/All Mail", readonly=True)
    #Return id's of all emails since midnight last night
    typ, msg_ids = conn.search(None, search)
    #print the id's
    print "All Mail ", typ, msg_ids

    if typ != 'OK': #exit if it doesn't 
        sys.exit("Bad search for emails in last day")

    ids = msg_ids[0].split(" ")
#    print ids
#    typ, msg_data = conn.fetch(id[0], '(BODY.PEEK[HEADER] FLAGS)')
#    pprint(msg_data)

    ######################################################### COUNT

    blacklist = ['khan']
    count = 0

    for num in range(len(ids)):
        flag = 0
        typ, msg_data = conn.fetch(ids[num], '(BODY.PEEK[HEADER])')
        for response_part in msg_data:
            if isinstance(response_part, tuple):
                if 'admission'.upper() in response_part[1].upper():
                    #print 'admission'.upper()
                    #print response_part[1].upper()
                    count += 1
                    flag = 1

        if flag == 0:
            typ, msg_data = conn.fetch(ids[num], '(BODY.PEEK[TEXT])')
            for response_part in msg_data:
                if isinstance(response_part, tuple):
                    if 'admission'.upper() in response_part[1].upper():
                        #print 'admission'.upper()
                        #print response_part[1].upper()
                        count += 1
                        flag = 1

        if flag == 1:
            typ, msg_data = conn.fetch(ids[num], '(RFC822)')
            for response_part in msg_data:
                if isinstance(response_part, tuple):
                    msg = email.message_from_string(response_part[1])
                    print count
                    for header in [ 'subject', 'from' ]:
                        print '%-8s: %s' % (header.upper(), msg[header])
                    print
                    #DSubtract one from the count if the email is from one on the blacklist
                    for name in blacklist:
                        if name.upper() in msg['from'].upper():
                            count -= 1

    print 
    print count
finally:
    try:
        conn.close()
    except:
        pass
    conn.logout()
