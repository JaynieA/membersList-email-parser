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

def combineHeaderAndBodyInfo(headerInfo, bodyInfo):
    totalParts = []
    #If item not already in totalParts list, add it
    #Do this for info from the header
    if headerInfo != None and headerInfo != []:
        for part in headerInfo:
            if part not in totalParts:
                totalParts.append(part)
    #And info from the body
    if bodyInfo != None and bodyInfo != []:
        for part in bodyInfo:
            if part not in totalParts:
                totalParts.append(part)
    return totalParts

def combineCompleteHeadAndBodyInfo(completeHeaderInfo, completeBodyInfo):
    #define header variables
    partsInSubject = completeHeaderInfo[0]
    conditionsInSubject = completeHeaderInfo[1]
    quantityInSubject = completeHeaderInfo[2]
    statusInSubject = completeHeaderInfo[3]
    #define body variables
    partsFromBody = completeBodyInfo[0]
    conditionsFromBody = completeBodyInfo[1]
    quantityFromBody = completeBodyInfo[2]
    statusFromBody = completeBodyInfo[3]
    #Parts
    totalParts = combineHeaderAndBodyInfo(partsInSubject, partsFromBody)
    totalParts = setDefaultIfNoneType(totalParts, 'part')
    #Conditions
    totalCondition = combineHeaderAndBodyInfo(conditionsInSubject, conditionsFromBody)
    totalCondition = setDefaultIfNoneType(totalCondition, 'condition')
    #Quantity
    totalQuantity = combineHeaderAndBodyInfo(quantityInSubject, quantityFromBody)
    totalQuantity = setDefaultIfNoneType(totalQuantity, 'quantity')
    #Status
    totalStatus = combineHeaderAndBodyInfo(statusInSubject, statusFromBody)
    totalStatus = setDefaultIfNoneType(totalStatus, 'status')
    #print('Parts:', totalParts, '\nConditions:', totalCondition, '\nQuantity:', totalQuantity, '\nStatus:', totalStatus)
    return [totalParts, totalCondition, totalQuantity, totalStatus]

def formatDetailInsertStatements(totalCombinedInfoList):
    #separate out variables from totalCombinedInfoList
    totalParts = totalCombinedInfoList[0]
    totalCondition = totalCombinedInfoList[1]
    totalQuantity = totalCombinedInfoList[2]
    totalStatus = totalCombinedInfoList[3]
    #if totalParts is valid, continue
    if totalParts != 'ERROR':
        #Make all list lenghts the same length as parts list for an accurate number of inserts
        secondaryInfo = [totalCondition, totalStatus, totalQuantity]
        for infoType in secondaryInfo:
            #If it's not as long, make it longer by duplicating last in list
            while len(totalParts) > len(infoType):
                infoType.append(infoType[-1])
        #initialize counter variable
        counter = 0
        #For as long as there are parts, continue to make inserts
        for part in totalParts:
            print('Line #',(counter+1))
            print('DETAIL INSERT:',totalQuantity[counter], totalCondition[counter], part)
            #increment the counter
            counter += 1

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
    for ch in ['=3D', '=A0', '*', '.', '(', ')', '=', ':']:
        if ch in emailBody:
            emailBody=emailBody.replace(ch,'')
    #Transform email body to all uppercase
    emailBody = emailBody.upper()
    return emailBody

def formatMainInsertStatement(mainInsertInfo):
    postDate = mainInsertInfo[0]
    postTime = mainInsertInfo[1]
    senderEmail = mainInsertInfo[2]
    senderName = mainInsertInfo[3]
    status = mainInsertInfo[4]
    print('MAIN INSERT:', postDate, postTime, senderEmail, senderName, status)

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
                if word == 'NEW':
                    word = 'NIB'
                if word.startswith('USED'):
                    word = 'USED'
                if word.startswith('REF'):
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

def getInfoFromHeader(subjectLine):
    #Parts
    partsInSubject = getParts(subjectLine)
    #Condition
    conditionsInSubject = getCondition(subjectLine)
    #Status
    statusInSubject = getStatus(subjectLine)
    #Quantity
    quantityInSubject = getQuantity(subjectLine)
    #print('Parts:',partsInSubject, '\nConditions:', conditionsInSubject, '\nQuantity:', quantityInSubject, '\nStatus:', statusInSubject)
    #Save all info from parsing the header into a list and return it
    return [partsInSubject, conditionsInSubject, quantityInSubject, statusInSubject]

def getInfoFromBody(emailBody):
    #Parts
    partsFromBody = getParts(emailBody)
    #Condition
    conditionsFromBody = getCondition(emailBody)
    #Quantity
    quantityFromBody = getQuantity(emailBody)
    #Status
    statusFromBody = getStatus(emailBody)
    #print('Parts:', partsFromBody, '\nConditions:', conditionsFromBody, '\nQuantity:', quantityFromBody, '\nStatus:', statusFromBody)
    return [partsFromBody, conditionsFromBody, quantityFromBody, statusFromBody]

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
    allNumbers = re.findall('\s+(\d+(?!/)(?!\d/))', string)
    quantity = []
    #append all numbers under 100 to quantity list
    for number in allNumbers:
        if len(number) < 3:
            quantity.append(number)
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

def setDefaultIfNoneType(infoList, type):
    #possibilities for type: part, condition, quantity, status
    if infoList is None or infoList == []:
        if type is 'part':
            #TODO: make sure there is logic for this error before db inserts
            infoList = 'ERROR'
        if type is 'condition':
            infoList = ['REF']
        if type is 'quantity':
            infoList = ['1']
        if type is 'status':
            infoList = ['RFQ']
    return infoList

def parseRawEmailMessages(msg, data, emailNumber):
    #Print Position of Current Email that is parsing
    print('Email #:', emailNumber)
    #Get Email Subject Line
    subjectLine = formatString(getSubjectLine(msg))
    #Get Email Sender's Info
    senderName = getSenderInfo(msg)[0]
    senderEmail = getSenderInfo(msg)[1]
    #Get and Print Message DATE & TIME
    date = getDateTime(msg)[0]
    time = getDateTime(msg)[1]
    #Parse SUBJECT LINE for parts, condition, status, quantity
    completeHeaderInfo = getInfoFromHeader(subjectLine)
    #Get the body of the email
    emailBody = getEmailTextFromBody(data)
    #Format the text of the email body
    emailBody = formatEmailBody(emailBody, senderName)
    #print(emailBody)
    #Parse EMAIL BODY for parts, condition, status, quantity
    completeBodyInfo = getInfoFromBody(emailBody)
    #COMBINE parsed info from head and body
    totalCombinedInfo = combineCompleteHeadAndBodyInfo(completeHeaderInfo, completeBodyInfo)
    #TODO: get company name from person email

    #TODO: either get row number from db or figure out how to generate and pass in
    #insert with: postDate, postTime, senderEmail, senderName, companyName, contactId, status
    mainStatus = totalCombinedInfo[3][0]
    mainInsertInfo = [date, time, senderEmail, senderName, mainStatus]
    formatMainInsertStatement(mainInsertInfo)

    #Format insert statements for the database
    #TODO: send along into below function: emailNumber, lineCounter, part, qty, condition
    formatDetailInsertStatements(totalCombinedInfo)

    #Print a dividing line between each email for clarity
    print('~~~~~~~~~~~~~~~~~~~~~~EMAIL END~~~~~~~~~~~~~~~~~~~~~~')

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
