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

#CLASSES
#Create a class for the lines in the email body
class EmailBodyLine:
    "The information contained in a line of the email body"
    lineCount = 0
    #constructor
    def __init__(self, partList, conditionList, quantityList, statusList):
        self.parts = partList
        self.conditions = conditionList
        self.quantity = quantityList
        self.status = statusList
    #Method
    def displayLineInfo(self):
        print ('Parts:', self.parts, ', Conditions:', self.conditions, ', Quantity:', self.quantity, ', Status:', self.status)

##FUNCTIONS
#initializes app
def init():
    loginToEmail(M, EMAIL_ACCOUNT, EMAIL_FOLDER, PASSWORD)

def condenseList(listName):
    if len(listName) == 0:
        listName = None
    elif len(listName) == 1:
        listName = listName[0]
    return listName

def formatEmailBody(emailBody, senderName):
    #   Get rid of everything after 'Uneda Code of Conduct Policy'
    emailBody = emailBody.split('UNEDA Code of Conduct Policy')[0]
    #get rid of everything after the sender's signature (if it exists)
    if senderName != '':
        senderFirstName = senderName.split()[0]
        if senderFirstName in emailBody:
            emailBody = emailBody.split(senderName)[0]
    #get rid of everything after logo image if exists
    for word in ['[cid:', '[image:']:
        if word in emailBody:
            emailBody = emailBody.split(word)[0]
    #get rid of characters in list
    for ch in ['=3D', '=A0', '*', '.', '(', ')', '=']:
        if ch in emailBody:
            emailBody=emailBody.replace(ch,'')
    #Transform email body to all uppercase
    emailBody = emailBody.upper()
    return emailBody

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
        for condition in ['NIB', 'NEW', 'NOB', 'REF', 'USED', 'REFURB']:
            if word.startswith(condition):
                if word == 'REFURB':
                    word = 'REF'
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

def getEmailTextFromBody(data):
    #parses raw email body (email message type), and returns the text content of the email as a string
    body = email.message_from_bytes(data[0][1]).get_payload()
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
        for status in ['WTB', 'RFQ', 'WTS', 'SELL']:
            if word.startswith(status):
                #change 'SELL' to 'WTS'
                if status == 'SELL':
                    status = 'WTS'
                #append matches to statusFound list
                statusFound.append(status)
                return statusFound

def getSubjectLine(msg):
    #Gets raw subject line from raw message string
    subject = str(email.header.make_header(email.header.decode_header(msg['Subject'])))
    return subject

def getQuantity(string):
    #Use to parse the email subject ONLY
    #Gets clear number not followed by "/" or a digit followed by "/"
    quantity = re.findall('\s+(\d+(?!/)(?!\d/))', string)
    for number in quantity:
        #deletes numbers added to quantity list with lenghts over 4
        if len(number) > 4:
            quantity.remove(number)
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

def splitAndParseEmailBody(emailBody, emailNumber):
    #SPLIT EMAIL BODY INTO LINES TO PARSE
    emailBodyByLine = emailBody.split('\n')
    lines = []
    for line in emailBodyByLine:
        line = line.strip()
        line = line.replace('\t', ' ')
        if line != '':
            lines.append(line)
    #print('All Email Lines :',lines)

    #hold all objects created from body parsing in list
    allBodyObjects = []
    #For each line in the email, find the following:
    lineCounter = 1
    for line in lines:
        #Condition
        conditionInLine = getCondition(line)
        conditionInLine = condenseList(conditionInLine)
        #Part Number beginning with 'AIR-'
        partInLine = getParts(line)
        partInLine = condenseList(partInLine)
        #Quantity
        quantityInLine = [int(s) for s in line.split() if s.isdigit() and len(s) < 4]
        quantityInLine = condenseList(quantityInLine)
        #Status
        statusInLine = getStatus(line)
        if statusInLine == None:
            statusInLine = []
        statusInLine = condenseList(statusInLine)

        #Make an object of the findings from the line and display it
        # IF all lists are not empty, use constructor to make the object
        if partInLine == None and conditionInLine == None and quantityInLine == None and statusInLine == None:
            pass
        else:
            #   concatenate a name for the object
            objName = str(emailNumber) + '.' + str(lineCounter)
            #print(objName)
            objName = EmailBodyLine(partInLine, conditionInLine, quantityInLine, statusInLine)
            #   call it's method displayLineInfo to display the new object's info
            #objName.displayLineInfo()
            allBodyObjects.append(objName)
            #increment the line counter
            lineCounter += 1
    return allBodyObjects

def parseRawEmailMessages(msg, data, emailNumber):
    #Print Position of Current Email that is parsing
    print('Email #:', emailNumber)

    #Get Email Subject Line
    subjectLine = formatString(getSubjectLine(msg))
    #print('Subject Line:' ,  subjectLine)

    #Get Email Sender's Info
    senderName = getSenderInfo(msg)[0]

    #Get and Print Message DATE & TIME
    date = getDateTime(msg)[0]
    print('Date:', date)
    time = getDateTime(msg)[1]
    print('Time:', time)

    print('Sender Name:', senderName)
    senderEmail = getSenderInfo(msg)[1]
    print('Sender Email:', senderEmail)

    #Parse Subject Line for parts, condition, status
    print('\nHEADER')
    #Parts
    partsInSubject = getParts(subjectLine)
    partsInSubject = condenseList(partsInSubject)
    print('Parts:', partsInSubject)
    #Condition
    conditionsInSubject = getCondition(subjectLine)
    conditionsInSubject = condenseList(conditionsInSubject)
    print('Conditions:', conditionsInSubject)
    #Status
    statusInSubject = getStatus(subjectLine)
    statusInSubject = condenseList(statusInSubject)
    print('Status:', statusInSubject)
    #Quantity
    quantityInSubject = getQuantity(subjectLine)
    quantityInSubject = condenseList(quantityInSubject)
    print('Quantity:', quantityInSubject)

    #Get the body of the email
    emailBody = getEmailTextFromBody(data)
    #Format the text of the email body
    emailBody = formatEmailBody(emailBody, senderName)

    #print(emailBody)

    allBodyObjects = splitAndParseEmailBody(emailBody, emailNumber)
    #Save all info from parsing the header into a list
    completeHeaderInfo = [partsInSubject, conditionsInSubject, statusInSubject, quantityInSubject]
    #Send all body objects and header info to be organized
    organizeInfoToInsert(allBodyObjects, completeHeaderInfo)

    #Print a dividing line between each email for clarity
    print('~~~~~~~~~~~~~~~~~~~~~~EMAIL END~~~~~~~~~~~~~~~~~~~~~~')

def organizeInfoToInsert(allBodyObjects, completeHeaderInfo):
    #Loop through the objects returned
    bodyResultsLength = len(allBodyObjects)
    #If no info objects returned from email body
    if bodyResultsLength == 0:
        print('HEADER INFO ONLY ')

        #If singular values only (no list values) in header
        if all(type(i) != list for i in completeHeaderInfo):
            print('all singular stuff in header')

    #If info objects have been returned from email body
    elif bodyResultsLength > 0:
        print('\nBODY')
        for item in allBodyObjects:
            item.displayLineInfo()


#Retrieves emails and initializes parsing
def retrieveEmails(host):
    rv, data = host.search(None, "ALL")
    if rv != 'OK':
        print("No messages found!")
        return
    #Process each email in the folder
    emailCounter = 1
    for num in data[0].split():
        rv, data = host.fetch(num, '(RFC822)')
        #Return if there is an error
        if rv != 'OK':
            print("ERROR getting message", num)
            return
        #Define raw email message as msg
        msg = email.message_from_bytes(data[0][1])
        #Parse msg to get desired data
        parseRawEmailMessages(msg, data, emailCounter)
        emailCounter += 1

#Initializes the app
init();

#Log out of the Email Account
M.logout()
