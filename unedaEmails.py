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
        EmailBodyLine.lineCount += 1
    #Method
    def displayLineCount(self):
        print('Line # %d' % EmailBodyLine.lineCount)
    #Method
    def displayLineInfo(self):
        print ('Parts:', self.parts, ', Conditions:', self.conditions, ', Quantity:', self.quantity, ', Status:', self.status)

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


def parseRawEmailMessages(msg, data, emailNumber):
    #Print Position of Current Email that is parsing
    print('Email #:', emailNumber)

    #Get Email Subject Line
    subjectLine = formatString(getSubjectLine(msg))
    print('Subject Line:' ,  subjectLine)

    #Get Email Sender's Info
    senderName = getSenderInfo(msg)[0]
    '''
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

    #Get the body of the email
    emailBody = getEmailTextFromBody(data)
    #Format the text of the email body
    emailBody = formatEmailBody(emailBody, senderName)

    print(emailBody)

    #SPLIT EMAIL BODY INTO LINES TO PARSE
    emailBodyByLine = emailBody.split('\n')
    lines = []
    for line in emailBodyByLine:
        line = line.strip()
        line = line.replace('\t', ' ')
        if line != '':
            lines.append(line)

    print('All Email Lines :',lines)

    #For each line in the email, find the following:
    lineCounter = 1
    for line in lines:
        #Condition
        conditionInLine = getCondition(line)
        if len(conditionInLine) == 1:
            conditionInLine = conditionInLine[0]
        if len(conditionInLine) == 0:
            conditionInLine = None
        #Part Number beginning with 'AIR-'
        partInLine = getParts(line)
        if len(partInLine) == 1:
            partInLine = partInLine[0]
        if len(partInLine) == 0:
            partInLine = None
        #Quantity
        quantityInLine = getQuantity(line)
        if len(quantityInLine) == 1:
            quantityInLine = quantityInLine[0]
        if len(quantityInLine) == 0:
            quantityInLine = None
        #status
        statusInLine = getStatus(line)
        if statusInLine == None:
            statusInLine = []
        if len(statusInLine) == 1:
            statusInLine = statusInLine[0]
        if len(statusInLine) == 0:
            statusInLine = None

        #Make an object of the findings from the line and display it
        #   concatenate a name for the object
        objName = str(emailNumber) + '.' + str(lineCounter)
        print(objName)
        # IF all lists are not empty, use constructor to make the object
        if partInLine == None and conditionInLine == None and quantityInLine == None and statusInLine == None:
            print('skip this line- it contains no info')
        else:
            objName = EmailBodyLine(partInLine, conditionInLine, quantityInLine, statusInLine)
            #   call it's method displayLineInfo to display the new object's info
            objName.displayLineInfo()

            #increment the line counter
            lineCounter += 1

    #Print a dividing line between each email for clarity
    print('~~~~~~~~~~~~~~~~~~~~~~END~~~~~~~~~~~~~~~~~~~~~~')



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

def getEmailTextFromBody(data):
    #parses raw email body (email message type), and returns the text content of the email as a string
    body = email.message_from_bytes(data[0][1]).get_payload()
    return body;

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
