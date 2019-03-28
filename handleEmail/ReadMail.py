import win32com.client
import pywintypes
import getpass
import config

mailServer = config.mailServer
mailPath = config.mailPath

notesSession = win32com.client.Dispatch(r'Lotus.NotesSession')

notesSession.Initialize(config.password)
notesDatabase = notesSession.GetDatabase(mailServer, mailPath)
print(notesDatabase.Title)
print(notesDatabase.Type)
folder = notesDatabase.getView('MiniView - Trash2')
