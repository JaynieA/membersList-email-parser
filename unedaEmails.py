#Import Libraries
import sys
import imaplib
import getpass
import email
import email.header
import datetime
import re
from email.utils import parseaddr
from connection import *

#Create an IMAP4 instance (with SSL for security) that connects to the gmail server
M = imaplib.IMAP4_SSL('imap.gmail.com')

##FUNCTIONS
#initializes app
def init():
    loginToEmail(M, EMAIL_ACCOUNT, EMAIL_FOLDER, PASSWORD)

def formatString(string):
    #Transform to uppercase
    string = string.upper()
    #Remove all characters in list
    for ch in [',',':', ';', '(',')','x','X', '[NETWORK-EQUIPMENT]', '.']:
        if ch in string:
            string=string.replace(ch,'')
    return string

def getCondition(string):
    conditions = []
    #split string parameter into words
    for word in string.split():
        #find words that fit condition criteria, append them to conditions list
        for condition in ['NIB', 'NEW', 'NOB', 'REF', 'USED']:
            if word.startswith(condition):
                conditions.append(word)
    return conditions

def getDateTime(tupule):
    date_tuple = email.utils.parsedate_tz(tupule['Date'])
    if date_tuple:
        local_date = datetime.datetime.fromtimestamp(
        email.utils.mktime_tz(date_tuple))
        postDate = local_date.strftime("%m-%d-%Y")
        postTime = local_date.strftime("%H:%M %p")
        return postDate, postTime

def getEmailBody(data):
    body = email.message_from_bytes(data[0][1])
    return body;

def getParts(string):
    result = []
    #split string parameter into words
    for word in string.split():
        #find words that start with 'AIR-'
        if word.startswith('AIR-'):
            #append matches to result list
            result.append(word)
    return result;

def getSenderInfo(msg):
    #Gets the message sender's name and email address from raw message string
    sender = str(email.header.make_header(email.header.decode_header(msg['from'])))
    senderName = (parseaddr(sender)[0])
    senderEmail = (parseaddr(sender)[1])
    return senderName, senderEmail

def getStatus(string):
    statusFound = []
    #split string parameter into words
    for word in string.split():
        #find words that match status criteria
        for status in ['WTB', 'RFQ', 'WTS']:
            if word.startswith(status):
                #append matches to statusFound list
                statusFound.append(status)
                return statusFound

def getSubjectLine(msg):
    #Gets raw subject line from raw message string
    subject = str(email.header.make_header(email.header.decode_header(msg['Subject'])))
    return subject

def getQuantity(string):
    #Gets clear number not followed by "/" or a digit followed by "/"
    quantity = re.findall('\s+(\d+(?!/)(?!\d/))', string)
    return quantity

def loginToEmail(host, account, folder, password):
    #Log in to gmail account
    try:
        rv, data = host.login(account, password)
    #display error message if login fails
    except imaplib.IMAP4.error:
        print('Login Failed')
        #exit or deal with login failure
        sys.exit(1)
    #Displays info about what account it's signing into and if it's successful
    print(rv, data)
    #Selects UNEDA mailbox (Marks new messages as read)
    rv, data = host.select(folder, readonly=False)
    if rv == 'OK':
        print("Processing UNEDA emails..\n")
        #Retrieves all emails
        retrieveEmails(host)
        host.close()
    else:
        print("ERROR: Unable to open UNEDA mailbox ", rv)

def parseRawEmailMessages(msg, data):
    #Get Email Subject Line
    subjectLine = formatString(getSubjectLine(msg))
    print('Subject Line:' ,  subjectLine)

    #Get Email Sender's Info
    senderName = getSenderInfo(msg)[0]
    print('Sender Name:', senderName)
    senderEmail = getSenderInfo(msg)[1]
    print('Sender Email:', senderEmail)

    #Parse Subject Line for parts, condition, status
    partsInSubject = getParts(subjectLine)
    print('Parts:', partsInSubject)
    conditionsInSubject = getCondition(subjectLine)
    print('Conditions:', conditionsInSubject)
    statusInSubject = getStatus(subjectLine)
    print('Status:', statusInSubject)
    quantityInSubject = getQuantity(subjectLine)
    print('Quantity:', quantityInSubject)

    #Get and Print Message DATE & TIME
    date = getDateTime(msg)[0]
    print('Date:', date)
    time = getDateTime(msg)[1]
    print('Time:', time)

    '''
    #   Get the body of the email
    emailBody = getEmailBody(data)
    print(emailBody)
    '''
    #Print a dividing line between each email for clarity
    print('~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~')



#Retrieves emails and initializes parsing
def retrieveEmails(host):
    rv, data = host.search(None, "ALL")
    if rv != 'OK':
        print("No messages found!")
        return
    #Process each email in the folder
    for num in data[0].split():
        rv, data = host.fetch(num, '(RFC822)')
        #Return if there is an error
        if rv != 'OK':
            print("ERROR getting message", num)
            return
        #Define raw email message as msg
        msg = email.message_from_bytes(data[0][1])
        #Parse msg to get desired data
        parseRawEmailMessages(msg, data)

#Initializes the app
init();

#Log out of the Email Account
M.logout()
