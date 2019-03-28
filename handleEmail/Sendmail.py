import os, uuid
import win32com.client
import config
import pythoncom
import itertools as it



class Sendmail:

    def __init__(self):
        pythoncom.CoInitialize()
        sess = win32com.client.Dispatch(r'Lotus.NotesSession')
        sess.Initialize(config.password)
        server_name = config.mailServer
        self.db_name = config.mailPath
        self.db = sess.getDatabase(server_name, self.db_name)
        print(server_name,self.db_name)

    def send_mail(self, subject, body_text, sendto, copyto=None, blindcopyto=None, attach=None):
        db = self.db
        if not db.IsOpen:
            try:
                db.Open()
            except:
                print('could not open database: {}'.format(self.db_name))

        doc = db.CreateDocument()
        doc.ReplaceItemValue("Form", "Memo")
        doc.ReplaceItemValue("From", "rajrobin@in.ibm.com")
        doc.ReplaceItemValue("Subject", subject)

        # assign random uid because sometimes Lotus Notes tries to reuse the same one
        uid = str(uuid.uuid4().hex)
        doc.ReplaceItemValue('UNIVERSALID', uid)

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
        doc.Send(True)


if __name__ == '__main__':
    subject = "test subject"
    body = "test body new body"
    sendto = ['rajnish.robin@gmail.com', ]
    #print(__name__)
    files = ['C:\\Users\\RajnishKumarVENDORRo\\PycharmProjects\\CDSApplication\\AttachmentList\\key.txt']
    attachment = it.takewhile(lambda x: os.path.exists(x), files)
    k = Sendmail()
    k.send_mail(subject, body, sendto, attach=attachment)
