

def makeDocumentGenerator(folderName,notesDatabase):
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