import win32com.client


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


mailServer = 'CN=NALLN234/OU=40/OU=LLN/O=IBM'
mailPath = r'data3\126\1000836090.nsf'
notesSession = win32com.client.Dispatch(r'Lotus.NotesSession')
notesSession.Initialize('6377rajn@')
notesDatabase = notesSession.GetDatabase(mailServer, mailPath)
username = notesSession.UserName
print(username)
print(notesDatabase.Title)
print(notesDatabase.Type)
for document in makeDocumentGenerator('($Inbox)', notesDatabase):
    record = document.GetItemValue('Body')
    # dateMail = datetime.datetime.fromtimestamp(document.GetItemValue('PostedDate')[0])
    dateMail = str(document.GetItemValue('PostedDate')[0])
    recived_from = str(document.GetItemValue('From')[0])
    subject = document.GetItemValue('Subject')[0]
    print(recived_from)
    print(subject)