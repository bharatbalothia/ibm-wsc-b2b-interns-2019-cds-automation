import docx
from docx.enum.style import WD_STYLE_TYPE
import difflib
import re



fields = ['Company Name','Contact Name','Email address','Technical Contact Name','Email address','Production EDI','Qualif','Test EDI']


def getText(filename):
    doc = docx.Document(filename)
    for para in doc.paragraphs:
        if re.search('[a-zA-Z]', para.text):
            yield para.text


def parse(filename):
    info = {}
    for lines in getText(filename):
        if any(s in lines for s in fields):
            resp = lines.split(":")
            if len(resp) > 1:
                info[resp[0]] = re.sub('_', '', resp[1])
    return info