import os, uuid
import win32com.client
import pywintypes # for exception


print('started')

def send_mail(subject,body_text,sendto,copyto=None,blindcopyto=None,
              attach=None):
    session = win32com.client.Dispatch(r'Lotus.NotesSession')
    session.Initialize('6377rajn@')
    mailServer = 'CN=NALLN234/OU=40/OU=LLN/O=IBM'
    mailPath = r'data3\126\1000836090.nsf'

    db = session.getDatabase(mailServer, mailPath)
    if not db.IsOpen:
        try:
            db.Open()
        except pywintypes.com_error:
            print( 'could not open database: {}'.format(mailPath) )

    doc = db.CreateDocument()
    doc.ReplaceItemValue("Form","Memo")
    doc.ReplaceItemValue("Subject",subject)

    # assign random uid because sometimes Lotus Notes tries to reuse the same one
    uid = str(uuid.uuid4().hex)
    doc.ReplaceItemValue('UNIVERSALID',uid)

    # "SendTo" MUST be populated otherwise you get this error:
    # 'No recipient list for Send operation'
    doc.ReplaceItemValue("SendTo", sendto)

    if copyto is not None:
        doc.ReplaceItemValue("CopyTo", copyto)
    if blindcopyto is not None:
        doc.ReplaceItemValue("BlindCopyTo", blindcopyto)

    # body
    body = doc.CreateRichTextItem("Body")
    body.AppendText(body_text)

    # attachment
    if attach is not None:
        attachment = doc.CreateRichTextItem("Attachment")
        for att in attach:
            attachment.EmbedObject(1454, "", att, "Attachment")

    # save in `Sent` view; default is False
    doc.SaveMessageOnSend = True
    print(body)
    doc.Send(False)

if __name__ == '__main__':
    subject = "test subject"
    body = "test body"
    sendto = ['Rajnish Kumar Robin <rajnish.k.robin@st.niituniversity.in>',]
    #files = ['/path/to/a/file.txt','/path/to/another/file.txt']
    #attachment = it.takewhile(lambda x: os.path.exists(x), files)

    send_mail(subject, body, sendto)
