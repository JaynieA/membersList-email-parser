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
    ##Turns into all caps and gets rid of commas and colons
    #   make string all caps
    string = string.upper()
    #   Replace commas with spaces
    string = string.replace(","," ")
    #   Replace colons with spaces
    string = string.replace(":"," ")
    #   Replace semicolons with spaces
    string = string.replace(";"," ")
    #   Replace parenthesis with spaces
    string = string.replace("("," ")
    string = string.replace(")"," ")
    #   Replace X's with spaces
    string = string.replace("x", " ")
    string = string.replace("X", " ")
    return string

def getEmailBody(data):
    body = email.message_from_bytes(data[0][1])
    return body;

def getSenderInfo(msg):
    #Gets the message sender's name and email address
    sender = str(email.header.make_header(email.header.decode_header(msg['from'])))
    senderName = (parseaddr(sender)[0])
    senderEmail = (parseaddr(sender)[1])
    return senderName, senderEmail

def getSubjectLine(msg):
    subject = str(email.header.make_header(email.header.decode_header(msg['Subject'])))
    return subject

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
    #   Get Email Subject Line
    subjectLine = formatString(getSubjectLine(msg))
    print(subjectLine)

    partsInSubject = getPartNumbersFromString(subjectLine)
    print('Parts in subject:',partsInSubject)

    #   Get Email Sender's Info
    senderName = getSenderInfo(msg)[0]
    senderEmail = getSenderInfo(msg)[1]
    print(senderName, senderEmail)


    '''
    #   Get the body of the email
    emailBody = getEmailBody(data)
    print(emailBody)
    #Print a dividing line between each email for clarity in idle
    print('~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~')
    '''
#FUNCTION: wordSplitter. Splits the string into words, and goes through each word in the string and finds the first word that starts with "AIR-"
def getPartNumbersFromString(string):
    result = []
    for word in string.split():
        if word.startswith('AIR-'):
            result.append(word)
    return result;

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


'''#FUNCTION: typeSplitter. Splits the string into words, and goes through each word in the string and finds the first word that fits type status criteria
def typeSplitter(subject):
    for word in subject.split():
        if word.startswith('WTB') or word.startswith("wtb") or word.startswith('RFQ') or word.startswith("rfq") or word.startswith("WTS") or word.startswith("wts"):
            return word

#FUNCTION: typeSplitter. Splits the string into words, and goes through each word in the string and finds the words that fits type condition criteria, appends them to list
def conditionSplitter(subject):
    hdrCdtn = []
    global hdrCdtn
    for word in subject.split():
        if word.startswith('NIB') or word.startswith("(NIB)") or word.startswith("nib") or word.startswith("(nib)")or word.startswith("(new)")or word.startswith("NEW") \
           or word.startswith('NOB') or word.startswith("(NOB)") or word.startswith("nob") or word.startswith("(nob)") \
           or word.startswith("REF") or word.startswith("(REF)") or word.startswith("ref") or word.startswith("(ref") \
           or word.startswith("USED") or word.startswith("(USED)") or word.startswith("used")or word.startswith("(used)"):
            hdrCdtn.append(word)
    #print ("Condition: ", hdrCdtn)
'''
