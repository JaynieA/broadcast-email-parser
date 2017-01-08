#Imports
import imaplib
import sys
import re
import time
import email.header
import pyodbc
import datetime
from email.utils import parseaddr
from connection import *

def init():
    emailLogin(EMAIL_ACCOUNT, PASSWORD)
    selectMailbox(EMAIL_FOLDER)

def deleteAllEmails(folderName):
    #deletes all emails in the email folder passed in as a param
    print('Deleting emails in %s...' % folderName)
    M.select(folderName)
    typ, data = M.search(None, 'ALL')
    for num in data[0].split():
        M.store(num, '+FLAGS', '\\Deleted')
    M.expunge()
    M.close()

def determineSleepTime():
    #Get current Time
    now_time = datetime.datetime.now().time()
    #Set start time at 6:00 PM
    start_time = datetime.time(18, 0, 0)
    #Set end time at 7:00 AM
    end_time = datetime.time(7, 0, 0)
    #If it is after 6:00PM, or before 7:00AM
    if (now_time >= start_time or now_time <= end_time):
        print('Sleeping for 3 hours...')
        return 10800
    else:
        print('sleeping for 1 hour...')
        return 3600

def emailLogin(account, password):
    #Login to gmail account
    try:
        rv, data = M.login(account, password)
    #Display error message if login fails
    except imaplib.IMAP4.error:
        currentTime = datetime.datetime.now().strftime('%I:%M %p')
        print('Login Failed at %s' % currentTime)
        #Exit to deal with login failure
        sys.exit(1)
    #display success info on login
    print(rv, data)
    currentTime = datetime.datetime.now().strftime('%I:%M %p')
    print('Logged in at %s' % currentTime)

def getCompanyName(line):
    try:
        return line.split(':')[1].strip()
    except IndexError:
        return ''

def getCompanyShortName(company_name):
    #connect to the database
    connection = pyodbc.connect(connStr)
    try:
        #Get company short name from database using company name in the email
        with connection.cursor() as cursor:
            sql = """SELECT Company.CompanyNameShort FROM Company WHERE (((Company.BbinName) = '"""+ company_name +"""'));"""
            cursor.execute(sql)
            #Get results from the query for company_name_short
            company_name_short = [x[0] for x in cursor.fetchall()]
            #Set to misc_reseller as default if company_name not in database
            if company_name_short == None or company_name_short == []:
                company_name_short = 'Misc_Reseller'
            #If company_name is in the database, define value as company_name_short
            else:
                company_name_short = company_name_short[0]
    finally:
        #close the connection to the database
        connection.close()
    return company_name_short

def getDateTimeInfo(body):
    #Use regex to get the date/time the original message was sent
    regex = r"(?<=Sent: )(.*)(?=To:)"
    match = re.findall(regex, body, re.DOTALL)[0].strip()
    #convert to datetime object
    match = time.strptime(match, "%A, %B %d, %Y %H:%M %p")
    #parse into separate date and time variables
    sendDate = time.strftime('%m-%d-%Y', match)
    sendTime = time.strftime('%H:%M %p', match)
    return [sendDate, sendTime]

def getEmailBody(msg, data):
    return email.message_from_bytes(data[0][1]).get_payload()

def getIndividualSearchesList(partSearchString):
    #Get rid of blank lines in partSearch string
    partSearchString = partSearchString.strip()
    #Discard the first line in partSearch to end up with just the searches done
    searchString = partSearchString.replace(partSearchString.split('\r\n')[0], '').strip()
    #Split the searchString into a list and return it
    return searchString.split('\n')

def getMessages():
    rv, data = M.search(None, 'ALL')
    #If there are no emails in the folder
    if rv != 'OK':
        print('No messages found')
        return
    #If there are emails in the folder, process them
    #initialize email counter
    email_count = 1
    for num in data[0].split():
        rv, data = M.fetch(num, '(RFC822)')
        #If there is an error, display it and return
        if rv != 'OK':
            print('ERROR getting message', num)
            return
        #Initialize msg as raw email message
        msg = email.message_from_bytes(data[0][1])
        #Display current email count, get data from emails
        print('Email #%s' % email_count)
        parseMessage(msg, data)
        #Increment email counter
        email_count += 1
        #Print dividing line for clarity
        print('\n-----------------------Email End-----------------------\n')

def getPersonName(line):
    regex = r'^.*?(?= P:)'
    return re.findall(regex, line, re.DOTALL)[0].strip()

def getPart(search):
    #Part = everything before first comma, formatted without returns and newlines
    return search.split(',')[0].replace('\r\n', '')

def getPartSearched(line):
    try:
        regex = r'^.*?(?= Searched by:)'
        return re.findall(regex, line, re.DOTALL)[0]
    except IndexError:
        return ''

def getSearchInfo(body):
    #Use regex to get the main info block containing search info from the email body
    regex = r'(?<=and Their Contact Info.)(.*)(?=To turn this feature off)'
    match = re.findall(regex, body, re.DOTALL)[0].strip()
    #get rid of all lines that start and end with '---''
    regex = r'(\---(.*?)\---)'
    match = re.sub(regex, '', match)
    return match

def parseMessage(msg, data):
    #Parse data from the email message
    body = getEmailBody(msg, data)
    #Date and Time message was sent
    sendDateTime = getDateTimeInfo(body)
    sendDate = sendDateTime[0]
    sendTime = sendDateTime[1]
    print('Date Sent: %s\nTime Sent: %s' % (sendDate, sendTime))
    #Separate search info from the rest of the body
    searchInfo = getSearchInfo(body)
    #Get individual part searches to parse through
    partSearches = splitPartSearches(searchInfo)
    #For each part searched...
    for partSearch in partSearches:
        #Get the part number for the part searched
        part_num = getPart(partSearch)
        #Display the part number currently bein parsed
        print('\nPart: %s' % part_num)
        print('Searches:')
        #Get a split list of each individual search done
        eachSearch = getIndividualSearchesList(partSearch)
        #Parse through each search by line and return a list of lists of results
        #Format of results lists is: [part_searched, company_name, person_first_name, person_last_name, company_name_short]
        completeSearchInfoList = parseIndividualSearches(eachSearch)
        #Insert a record into the database
        createRecord(completeSearchInfoList, part_num, sendDate, sendTime)

def createRecord(infoList, part, date, time):
    for search in infoList:
        part_searched = search[0]
        company_name = search[1]
        person_first_name = search[2]
        person_last_name = search[3]
        company_name_short = search[4]
        print (part_searched, company_name, person_first_name, person_last_name, company_name_short)
        #connect to the database:
        connection = pyodbc.connect(connStr)
        try:
            #insert record
            with connection.cursor() as cursor:
                sql = 'INSERT INTO tbl_broadcast_data (part_number, part_searched, search_date, search_time, company_name, company_name_short, person_first_name, person_last_name) VALUES (?, ?, ?, ?, ?, ?, ?, ?)'
                cursor.execute(sql, (part, part_searched, date, time, company_name, company_name_short, person_first_name, person_last_name))
        except pyodbc.IntegrityError:
            print ("INTEGRITY ERROR")
        finally:
            connection.close()

def parseIndividualSearches(eachSearch):
    #Parse through eachSearch by line
    completeSearchInfo = []
    lineCounter = 1
    for line in eachSearch:
        #If the line is an odd number (the first line in an individual search),
        #get the part searched and company named
        if (lineCounter % 2 != 0):
            part_searched = getPartSearched(line)
            company_name = getCompanyName(line)
            #connect to the database to get company short name
            company_name_short = getCompanyShortName(company_name)
        #Else, if the number is even (second line of individual search),
        #get the name of the person
        elif (lineCounter % 2 == 0):
            person_name = getPersonName(line)
            #split into first and last names
            if (person_name != 'N/A'):
                person_first_name = person_name.split(' ')[0]
                person_last_name = person_name.split(' ')[1]
            #set first and last names to blank string if name is 'N/A'
            else:
                person_first_name = ''
                person_last_name = ''
        #Print all info after search has been parsed (every 2 lines)
        if (lineCounter % 2 == 0):
            completeSearchInfo.append([part_searched, company_name, person_first_name, person_last_name, company_name_short])
        #Increment the lineCounter
        lineCounter += 1
    return completeSearchInfo

def selectMailbox(folder):
    #Select mailbox folder, mark new messages as read
    rv, data = M.select(folder, readonly = False)
    if rv == 'OK':
        print('Processing Emails in %s...\n' % folder)
        #Get all emails in the folder
        getMessages()
        #Close currently selected mailbox
        M.close()
    else:
        #Display the error
        print('ERROR. Unable to open %s, %s' % (folder, rv))

def splitPartSearches(searchList):
    #split into individual part searches
    return searchList.split('\r\n\r\n')

#initialize app
while True:
    print('App Initialized... \n')
    #Create an IMAP4 instance pointing to gmail server
    M = imaplib.IMAP4_SSL('imap.gmail.com')
    #Initialize login and parsing
    init()
    #delete all emails in Inbox
    deleteAllEmails('Inbox')
    #delete all emails that were parsed in EMAIL_FOLDER mailbox
    deleteAllEmails(EMAIL_FOLDER)
    #Log out of email account
    M.logout()
    #Sleep for 1 hour or 3 hours, depending on current time
    time.sleep(determineSleepTime())
