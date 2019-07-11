import win32com.client
import datetime
from datetime import date, timedelta
import pywintypes
import getpass
import os
import readingDocs,questionParser,readExcel
import Email
import itertools as it
from itertools import tee
import mailBody
import couchdb
from cloudant.client import Cloudant

#client = Cloudant("e06b4c3d-78c6-4f21-8bb3-297e658dc8b1-bluemix", "4299f4e82a5b36181a52abd82d9d74bd5bf3f77a350d5db2daf9b30882df8cb8",
                  #url="https://e06b4c3d-78c6-4f21-8bb3-297e658dc8b1-bluemix:4299f4e82a5b36181a52abd82d9d74bd5bf3f77a350d5db2daf9b30882df8cb8@e06b4c3d-78c6-4f21-8bb3-297e658dc8b1-bluemix.cloudantnosqldb.appdomain.cloud", connect='true')

#db = client['cdsdata']
server = couchdb.Server("http://%s:%s@9.199.145.193:5984/" % ("admin", "admin123"))
#self.db = client['cdsdata']
db =server['cdsproject']

#user = "admin"
#password = "admin"
#couch = couchdb.Server("http://%s:%s@9.199.145.49:5984/" % (user, password))
#db = couch['cdsdata']


def GetTime():
    currDate = date.today() #- timedelta(1)
    currDate = str(currDate.strftime("%Y-%m-%d"))
    yesterday = date.today() #- timedelta(2)
    yesterday = str(yesterday.strftime("%Y-%m-%d"))
    return currDate, yesterday


def makeDocumentGenerator(folderName,notesDatabase):
    folder = notesDatabase.GetView(folderName)
    folder.Refresh
    #NotesViewNavigator = folder.CreateViewNavFromAllUnread()
    if not folder:
        raise Exception('Folder "%s" not found' % folderName)
    # Get the first document
    document = folder.GetFirstDocument()
    print(document)
    # If the document exists,
    number_of_documents = 0
    while document:
        # Yield it
        yield document
        # Get the next document
        number_of_documents +=1
        print(number_of_documents)
        document = folder.GetNextDocument(document)
        if number_of_documents < 80:
            continue
        else:
            folder.Refresh
            document = folder.GetFirstDocument()
            number_of_documents = 0


def main():
    # Get credentials
    mailServer = 'CN=NALLN234/OU=40/OU=LLN/O=IBM'
    mailPath = r'data3\126\1000836090.nsf'

    notesSession = win32com.client.Dispatch(r'Lotus.NotesSession')


    notesSession.Initialize('6377rajn@')
    notesDatabase = notesSession.GetDatabase(mailServer, mailPath)
    username = notesSession.UserName
    print(username)
    print(notesDatabase.Title)
    print(notesDatabase.Type)
    [currDate, yesterday] = GetTime()
    print(currDate)
    iteration = 0
    for document in makeDocumentGenerator('($Inbox)', notesDatabase):
        record = document.GetItemValue('Body')
        #dateMail = datetime.datetime.fromtimestamp(document.GetItemValue('PostedDate')[0])
        dateMail = str(document.GetItemValue('PostedDate')[0])
        recived_from = str(document.GetItemValue('From')[0])
        subject = document.GetItemValue('Subject')[0]
        print(recived_from)
        if '[IBM EMEA CDS]' in subject:
            subjectParts = subject.split("-")
            customername = subjectParts[1]
            tpName = subjectParts[-1]
            print(dateMail) #2018-12-20 00:38:35+00:00
            print(document.Created) #2018-12-20 00:38:34+00:00
            print(subject)
            #print(document.Items[1])
            analysed_questions = {}
            surveyAttachment = []
            mailType = ""
            vanConn = ''
            try:
                for whichItem in range(len(document.Items)):
                    item = document.Items[whichItem]
                    if item.Name == '$FILE':
                        #print(item.Values[0])
                        attachment = document.GetAttachment(item.Values[0])
                        file_path = "C:\\Users\\RajnishKumarVENDORRo\\PycharmProjects\\CDS\\attachments\\"+item.Values[0]
                        attachment.ExtractFile("" + file_path)
                        if 'Question' in item.Values[0]:
                            mailType = 'Question'
                            l = [x for x in questionParser.questcheck(file_path)]
                            if len(l) == 0:
                                doc = {
                                    'status' : "Accepted",
                                    'file_path' : file_path,
                                    'unfilled_content' : []
                                }
                                analysed_questions[item.Values[0]] = doc
                            else:
                                doc = {
                                    'status': "Rejected",
                                    'file_path': file_path,
                                    'unfilled_content': l
                                }
                                analysed_questions[item.Values[0]] = doc
                        elif 'Survey' in item.Values[0]:
                            mailType = 'Survey'
                            surveyAttachment,vanConn = readExcel.surveyAnalysis(file_path)
                if mailType == "Question":
                    accepted = []
                    rejected = []
                    body = ""
                    attachments_to_be_sent = None
                    for key in analysed_questions:
                        if analysed_questions[key]['status'] == "Accepted":
                            accepted.append(analysed_questions[key]['file_path'])
                        elif analysed_questions[key]['status'] == "Rejected":
                            rejected.append(analysed_questions[key]['file_path'])
                    if len(rejected) == 0:
                        body = mailBody.FinishedQuestion
                        customer = db[customername]
                        for id in customer['TPlist']:
                            if customer['TPlist'][id]['TP name'] == tpName:
                                customer['TPlist'][id]["Status"] = "On-boarding Complete. Start Testing."
                                db.save(customer)
                    elif len(rejected) > 0 and len(accepted) == 0:
                        body+= "Please fill manditory fields in file/s:\n "
                        for file in rejected:
                            body+= file.split("\\")[-1]
                        attachments_to_be_sent = it.takewhile(lambda x: os.path.exists(x), rejected)
                        customer = db[customername]
                        for id in customer['TPlist']:
                            if customer['TPlist'][id]['TP name'] == tpName:
                                customer['TPlist'][id]["Status"] = "Some Files Rejected."
                                customer['TPlist'][id]["Rejected Files"] = [x.split("\\")[-1] for x in rejected]
                                db.save(customer)
                    elif len(rejected) > 0 and len(accepted) > 0:
                        body+= "These files has been accepted:\n"
                        for file in accepted:
                            body+= file.split("\\")[-1]+"\n"
                        body+= "Please fill these manditory fields in file/s:\n "
                        for file in rejected:
                            body+= file.split("\\")[-1]+"\n" + str(analysed_questions[file.split("\\")[-1]]['unfilled_content'])
                        attachments_to_be_sent = it.takewhile(lambda x: os.path.exists(x), rejected)
                        customer = db[customername]
                        for id in customer['TPlist']:
                            if customer['TPlist'][id]['TP name'] == tpName:
                                customer['TPlist'][id]["Status"] = "Some Files Rejected."
                                customer['TPlist'][id]["Rejected Files"] = [x.split("\\")[-1] for x in rejected]
                                db.save(customer)

                    print(key,analysed_questions[key]['status'],analysed_questions[key]['unfilled_content'])
                    sendto = [recived_from, ]
                    Email.send_mail(subject, body, sendto, attach=attachments_to_be_sent)
                    document.PutInFolder(r"CDS\Questionnarie", True)
                    document.RemoveFromFolder("($Inbox)")
                    #os.remove(file_path)
                elif mailType == 'Survey':
                    #print(any(True for x in surveyAttachment))
                    print("Its a survey")
                    survey_to_be_attached,surveyAttachment,fordb = tee(surveyAttachment,3)
                    if any(True for x in surveyAttachment):
                        body = mailBody.survey_reply
                        customer = db[customername]
                        for id in customer['TPlist']:
                            if customer['TPlist'][id]['TP name'] == tpName:
                                customer['TPlist'][id]['Questionnaire Sent'] = "YES"
                                customer['TPlist'][id]['Questionnaire Send Date'] = datetime.datetime.now().strftime("%Y-%m-%d %H:%M")
                                customer['TPlist'][id]['File names'] = [x.split("\\")[-1] for x in fordb]
                                customer['TPlist'][id]["Status"] = "Questionnaire Sent,"+vanConn
                                db.save(customer)
                    else:
                        body = mailBody.SurveyCompleted
                        customer = db[customername]
                        for id in customer['TPlist']:
                            if customer['TPlist'][id]['TP name'] == tpName:
                                customer['TPlist'][id]['Questionnaire Sent'] = "NO"
                                customer['TPlist'][id]["Status"] = vanConn+" Survey Completed. "
                                db.save(customer)
                    sendto = [recived_from, ]
                    print(body)
                    Email.send_mail(subject, body, sendto, attach=survey_to_be_attached)
                    document.PutInFolder(r"CDS\Testing", True)
                    document.RemoveFromFolder("($Inbox)")
                else:
                    customer = db[customername]
                    for id in customer['TPlist']:
                        if customer['TPlist'][id]['TP name'] == tpName:
                            customer['TPlist'][id]["Status"] = "Manual Intervention required !!!"
                            db.save(customer)
                Attachment_folder = 'C:\\Users\\RajnishKumarVENDORRo\\PycharmProjects\\CDS\\attachments\\'
                for the_file in os.listdir(Attachment_folder):
                    file_path = os.path.join(Attachment_folder, the_file)
                    os.remove(file_path)
                print('##########################################################')
            except:
                customer = db[customername]
                for id in customer['TPlist']:
                    if customer['TPlist'][id]['TP name'] == tpName:
                        customer['TPlist'][id]["Status"] = "Manual Intervention required !!!"
                        db.save(customer)



if __name__ == "__main__":
    main()