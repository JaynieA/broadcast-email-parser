#Imports
import imaplib
import sys
import re
import time
import email.header
from email.utils import parseaddr
from connection import *

#Create an IMAP4 instance pointing to gmail server
M = imaplib.IMAP4_SSL('imap.gmail.com')

def init():
    emailLogin(EMAIL_ACCOUNT, PASSWORD)
    selectMailbox(EMAIL_FOLDER)

def emailLogin(account, password):
    #Login to gmail account
    try:
        rv, data = M.login(account, password)
    #Display error message if login fails
    except imaplib.IMAP4.error:
        print('Login Failed')
        #Exit to deal with login failure
        sys.exit(1)
    #display success info on login
    print(rv, data)

def parseMessage(msg, data):
    #Parse data from the email message
    body = getEmailBody(msg, data)
    #Date and Time message was sent
    sendDateTime = getDateTimeInfo(body)
    sendDate = sendDateTime[0]
    sendTime = sendDateTime[1]
    print('Date Sent: %s\nTime Sent: %s\n' % (sendDate, sendTime))

    #separate search info from the rest of the body
    searchInfo = getSearchInfo(body)

    #TODO: extract the following into it's own function
    #Get individual part searches to parse through
    partSearches = splitPartSearches(searchInfo)
    #For each part searched...
    for partSearch in partSearches:
        #print(partSearch,'~~~\n')

        #Get the part number for the part searched
        part_num = getPart(partSearch)
        #print('Part:', part_num)

        #Get rid of blank lines in partSearch string
        partSearch = partSearch.strip()

        #Discard the first line in partSearch to end up with just the searches
        searchesOnly = partSearch.replace(partSearch.split('\r\n')[0], '').strip()
        #print('\n',searchesOnly)

        #Split a list of the searchesOnly string
        soloSearchList = searchesOnly.split('\n')
        
        #Parse through soloSearchList by line
        print('\nPart Search:')
        counter = 1
        for line in soloSearchList:
            #if the line is an odd number (the first line in an individual search),
            #get the part searched and company named
            if (counter % 2 != 0):
                print('Odd:', line)
            #else, if the number is even (second line of individual search),
            #
            elif (counter % 2 == 0):
                print('Even:', line)
            #increment the counter
            counter += 1


    ##print(body)

def getPart(search):
    #Part = everything before first comma, formatted without returns and newlines
    return search.split(',')[0].replace('\r\n', '')

def getSearchInfo(body):
    #Use regex to get the main info block containing search info from the email body
    regex = r'(?<=and Their Contact Info.)(.*)(?=To turn this feature off)'
    match = re.findall(regex, body, re.DOTALL)[0].strip()
    #get rid of all lines that start and end with '---''
    regex = r'(\---(.*?)\---)'
    match = re.sub(regex, '', match)
    return match

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
        print('\nEmail #%s' % email_count)
        parseMessage(msg, data)
        #Increment email counter
        email_count += 1

def selectMailbox(folder):
    #Select mailbox folder, mark new messages as read
    rv, data = M.select(folder, readonly = False)
    if rv == 'OK':
        print('Processing Emails in %s \n' % folder)
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
init()
