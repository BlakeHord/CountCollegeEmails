import schedule
import time
import openpyxl as xl
import imaplib
import email
import datetime
import sys
import re
from password import *
import urllib2

blacklist = ['khan']

######################################################### FUNCTIONS

def internet_on():
    try:
        response=urllib2.urlopen('http://google.com',timeout=5)
        print "connected to the internet!"
        return True
    except urllib2.URLError as err: pass
    print "not connected to the internet..."
    return False

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

def set_averages(sheet):
    email_avg = []
    letter_avg = []
    for start in range(3,10): #Make average formulas for emails
        local_avg = '=AVERAGE(B'+str(start)
        for num in range(start + 7,MAXROW + 1,7):
            local_avg += ",B"+str(num)
        local_avg += ")"
        email_avg.append(local_avg)

    for start in range(10,17): #Make average formulas for letters
        local_avg = '=AVERAGE(C'+str(start)
        for num in range(start + 7,MAXROW + 1,7):
            local_avg += ",C"+str(num)
        local_avg += ")"
        letter_avg.append(local_avg)

    ctr = 0
    for i in range(6,13): #Write email average to cells
        cell_str = 'I'+str(i)
        sheet[cell_str].value = email_avg[ctr]
        ctr += 1

    ctr = 0
    for i in range(6,13): #Write letter average to cells
        cell_str = 'J'+str(i)
        sheet[cell_str].value = letter_avg[ctr]
        ctr += 1

######################################################### JOB TO RUN EVERY DAY
def job():
    MAXROW = 358
    count = 0

    ######################################################### CONNECT

    while internet_on() == False:
        time.sleep(5)

    conn = imaplib.IMAP4_SSL("imap.gmail.com", 993)
    conn.login(username, password)

    list_response_pattern = re.compile(r'\((?P<flags>.*?)\) "(?P<delimiter>.*)" (?P<name>.*)')

    ######################################################### SEARCH

    print "How many days ago? 0 = yesterday.\n"
    add = int(raw_input(""))
    
    today = datetime.date.today() - datetime.timedelta(add)
    yesterday = today - datetime.timedelta(1)

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
    #    print msg_data

        ######################################################### COUNT

        for num in range(len(ids)):
            flag = 0
            typ, msg_data = conn.fetch(ids[num], '(BODY.PEEK[HEADER])')
            for response_part in msg_data:
                if isinstance(response_part, tuple):
                    if 'admission'.upper() in response_part[1].upper() or ('Recruitment'.upper() in response_part[1].upper() and 'University'.upper() in response_part[1].upper()):
                        #print 'admission'.upper()
                        #print response_part[1].upper()
                        count += 1
                        flag = 1
                        break

            if flag == 0:
                typ, msg_data = conn.fetch(ids[num], '(BODY.PEEK[TEXT])')
                for response_part in msg_data:
                    if isinstance(response_part, tuple):
                        if 'admission'.upper() in response_part[1].upper() or ('Recruitment'.upper() in response_part[1].upper() and 'University'.upper() in response_part[1].upper()):
                            #print 'admission'.upper()
                            #print response_part[1].upper()
                            count += 1
                            flag = 1
                            break

            if flag == 1:
                typ, msg_data = conn.fetch(ids[num], '(RFC822)')
                for response_part in msg_data:
                    if isinstance(response_part, tuple):
                        msg = email.message_from_string(response_part[1])
                        print count
                        for header in [ 'subject', 'from' ]:
                            print '%-8s: %s' % (header.upper(), msg[header])
                        #Subtract one from the count if the email is from one on the blacklist
                        for name in blacklist:
                            if name.upper() in msg['from'].upper():
                                count -= 1
        print count
    finally:
        try:
            conn.close()
        except:
            pass
        conn.logout()

    #number of spam emails now stored in count variable
    ######################################################### WRITE TO SPREADSHEET

   
job()

