import win32com.client
import datetime
from datetime import date, timedelta
import time, re
import uuid
import smtplib
import json
import requests
import schedule



def main_module():
    mailServer = 'D01ML253/01/M/IBM'
    mailPath = "mail1\spssls"

    mailPassword = creds['SPSS_NOTES_PASS']
    notesSession = win32com.client.Dispatch('Lotus.NotesSession')

    notesSession.Initialize(mailPassword)
    notesDatabase = notesSession.GetDatabase(mailServer, mailPath)

    [currDate, yesterday] = GetTime()
    # currDate = '2017-09-22'
    # yesterday = '2017-09-21'
    print(currDate)

    # Iterate over Emails
    ct = 0
    yesct = 0
    for document in makeDocumentGenerator('($Inbox)', notesDatabase):
        dateMail = datetime.datetime.fromtimestamp(int(document.GetItemValue('PostedDate')[0]))
        dateMail = str(dateMail)
        SPSSLog = {}
        SPSSLog['timestamp'] = dateMail
        if str(dateMail).startswith(currDate):
            ct += 1
            print(ct)
            # if (ct <= 62):
            #    continue
            [body_text, sendto, found, SPSSLog] = ParseEmail(document, SPSSLog, notesDatabase)
            if (len(sendto)) > 0:  # Valid Email ID
                # print sendto
                SendEmail(body_text, sendto, found, notesDatabase, notesSession)
                CloudantLogging(SPSSLog)
            # print body_text
            '''''
            try:
                [body_text, sendto, found, SPSSLog] = ParseEmail(document, SPSSLog)
                #SendEmail(body_text, sendto, found,notesDatabase,notesSession)
                #CloudantLogging(SPSSLog)
            except:
                pass

                server = smtplib.SMTP('smtp.gmail.com', 587)
                server.starttls()
                server.login(creds['GMAIL_USER'], creds['GMAIL_PASS'])
                fromWhom = document.GetItemValue('From')[0].strip()
                sendto = fromWhom.partition('<')[-1].rpartition('>')[0]
                msg = "Date: " + dateMail + " Sender: " + sendto
                server.sendmail(creds['GMAIL_USER'], creds['GMAIL_SENDER'], msg)
                server.quit()
            '''''

        else:
            if str(dateMail).startswith(yesterday):
                yesct = yesct + 1
                if yesct > 20:
                    exit(0)


def makeDocumentGenerator(folderName, notesDatabase):
    """
        INPUT: Lotus Notes folder name, in this case Inbox
        Iterate over all emails in the folder
    """

    # Get folder
    folder = notesDatabase.GetView(folderName)
    if not folder:
        raise Exception('Folder "%s" not found' % folderName)
    # Get the first document
    document = folder.GetFirstDocument()
    # If the document exists,
    while document:
        # Yield it
        yield document
        # Get the next document
        document = folder.GetNextDocument(document)



def SendEmail(Resp, sendto, found, notesDatabase, notesSession):
    """
        INPUT
        :param LiceneseAuthcode: License Authorization key extracted from the bot.
        :param sendto: Sender Email
        :param found: found = True if authCode, lockCode, and LicenseActivation key found
        Send email with appropriate response text to the user. Response is binary.  If License Activation key is present, send the key with appropriate other body message
        Else send error message
    """

    doc = notesDatabase.CreateDocument()
    uid = str(uuid.uuid4().hex)
    doc.ReplaceItemValue('UNIVERSALID', uid)
    doc.ReplaceItemValue("SendTo", sendto)
    doc.ReplaceItemValue("Form", "Memo")
    doc.ReplaceItemValue("Subject", "Activation Code")

    if not found:
        # body
        body = doc.CreateRichTextItem("Body")
        body.AppendText(Resp)

        # save in `Sent` view; default is False
        doc.SaveMessageOnSend = True
        doc.Send(False)

    else:
        body = doc.CreateRichTextItem("Body")
        body.AppendText(Resp)
        doc.SaveMessageOnSend = True
        doc.Send(False)


def GetTime():
    """
       INPUT:
       OUTPUT: Get yesterday's date
       Run the job once a day, and handle the emails from yesterday
    """

    # Get yesterday's time ...........
    currDate = date.today() - timedelta(1)
    currDate = str(currDate.strftime("%Y-%m-%d"))
    yesterday = date.today() - timedelta(2)
    yesterday = str(yesterday.strftime("%Y-%m-%d"))
    return currDate, yesterday


def ParseEmail(document, SPSSLog, notesDatabase):
    """
        INPUT: current email
        Extract authCode and lockCode from the email body. Extract sender email address.
        If authoCode and lockCode present:
            get LicenseAuthCode
            If LicenseAuthCode found, found = True else found = False
    """

    body = document.GetItemValue('Body')[0].encode('utf-8').strip()
    authCode = re.findall(r'([a-fA-F0-9]{20})', body)
    lockCode = re.findall(r'[10|4]{1,2}-[a-zA-Z0-9]{5}', body)

    ErrorEmailResp = 'You have reached an automated inbox.  We were unable to create your IBM SPSS license string.  There was a problem with your request.  The most common reasons for failure to create a license string include: \n\n \
        1. Your request did not include a lock code.  http://www-01.ibm.com/support/docview.wss?uid=swg21980079 \n\n \
        2. Your request did not include an authorization code.  http://www-01.ibm.com/support/docview.wss?uid=swg21478930 \n\n \
        3. Your request included an authorization code that has no remaining installations. http://www-01.ibm.com/support/docview.wss?uid=swg21480566 \n\n \
        4. Your request included an authorization code that does not exist in our records.  Please check the code and try again. \n\n \
        5. Your request included an expired authorization code.  Please contact Sales to renew your contract for current codes.  https://www.ibm.com/us-en/marketplace/analytics#category-headline \n\n \
        For more information on licensing your SPSS products and to review your support options on the Predictive Analytics Community Get Help page: https://developer.ibm.com/predictiveanalytics/get-help/ \n\n \
        Thank you for contacting IBM Support.'

    # Check that lock code does not start with 0
    if len(lockCode) > 0:
        if lockCode[0].startswith('0'):
            lockCode = []

    fromWhom = document.GetItemValue('From')[0].strip()
    sendto = fromWhom.partition('<')[-1].rpartition('>')[0]
    # sendto = "prateeti@gmail.com"
    print(sendto)
    # print fromWhom, sendto

    SPSSLog['body'] = body
    SPSSLog['AuthCode'] = authCode
    SPSSLog['LockCode'] = lockCode
    SPSSLog['sender'] = sendto

    found = False

    if ((len(authCode) > 0) and (len(lockCode) > 0)):
        found = True
        try:
            LicenseAuthCode = getLicense(str(authCode[0]), str(lockCode[0]))
            print
            LicenseAuthCode
        except:
            # Send Error Email to prateeti@gmail.com
            sendto = 'prateeti@gmail.com'
            doc = notesDatabase.CreateDocument()
            uid = str(uuid.uuid4().hex)
            doc.ReplaceItemValue('UNIVERSALID', uid)
            doc.ReplaceItemValue("SendTo", sendto)
            doc.ReplaceItemValue("Form", "Memo")
            doc.ReplaceItemValue("Subject", "BOT ERROR!!")
            body = doc.CreateRichTextItem("Body")
            body.AppendText("Error Bot failed")
            doc.SaveMessageOnSend = False
            doc.Send(False)
            driver.quit()
            exit()

        SPSSLog['LicenseKey'] = LicenseAuthCode
        if LicenseAuthCode == "LicenseKey Error":
            body_text = ErrorEmailResp
            SPSSLog['LicenseKey'] = 'LicenseKey Error'


        else:
            body_text = "You have reached an automated inbox.  Your IBM SPSS license string is: \n\n"
            body_text = body_text + str(LicenseAuthCode) + "\n"
            body_text = body_text + "To apply your code: \n\n 1.	Go to Programs->All Programs->IBM SPSS Statistics->IBM SPSS Statistics [version] License Authorization Wizard. \n 2.	Right click on the program and select 'Run as Administrator'. \n 3.	Follow the screens and when prompted for your authorization code, copy & paste in the license string. \n 4.	Complete the screens until the end & start your product! \n\n Note that the product and version of your license string must match those of the installed product. \n \n Mismatch will cause an error on start of SPSS products: http://www-01.ibm.com/support/docview.wss?uid=swg21486143 \n \n To resolve, please review your support options on the Predictive Analytics Community Get Help page: \n https://developer.ibm.com/predictiveanalytics/get-help/ \n \n Thank you for contacting IBM Support."

    else:
        SPSSLog['LicenseKey'] = 'NA'
        # Negative Email Response
        body_text = ErrorEmailResp

    SPSSLog['EmailResp'] = body_text

    return body_text, sendto, found, SPSSLog


def CloudantLogging(SPSSLog):
    """
        INPUT: Logs in json
        Store the logs in cloudant
    """

    json_data = json.dumps(SPSSLog)
    # print SPSSLog

    try:
        headers = {'Content-Type': 'application/json'}
        response = requests.post('https://pratemoh.cloudant.com/spsslogs', auth=('pratemoh', 'babamama721302'),
                                 data=json_data, headers=headers)
    except:
        print
        "FAILED", SPSSLog


if __name__ == "__main__":
    main_module()
    # schedule.every().day.at("13:16").do(main_module)