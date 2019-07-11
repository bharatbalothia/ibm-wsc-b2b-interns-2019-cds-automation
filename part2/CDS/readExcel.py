import docx
from docx.enum.style import WD_STYLE_TYPE
import difflib
import re
import Email
import os
import itertools as it
import pandas as pd
import couchdb

server = couchdb.Server("http://%s:%s@9.199.145.193:5984/" % ("admin", "admin123"))
db =server['cdstpdata']

quest_map = {
        "AS2" : ["IBM_AS2_Questionnaire_EuroGIODE(1).docx"],
        "OFTP2" : ["IBM_OFTP2_Questionnaire_EuroGIODE.docx"],
        "FTP(S) / SFTP" : ["IBM_SFTP_Questionnaire.docx","IBM_FTP_Questionnaire_EuroGIODE.docx"],
        "HTTP(S)" : ["IBM_HTTPS_Questionnaire.docx"]
    }


def surveyAnalysis(SurveyFile='TP_Survey_ABC_Corp_Universal.xlsm',customer_name = "ABC Corp."):
    df = pd.read_excel(SurveyFile,sheet_name='Sheet1')
    df = df.fillna('')
    TPData = df.to_dict(orient='records')[0]
    if customer_name not in db:
        tpdetailData = {
            "_id" : customer_name,
            "TPDetails" : [TPData]
        }
        db.save(tpdetailData)
    else:
        doc = db[customer_name]
        doc['TPDetails'].append(TPData)
        db.save(doc)
    #print(TPData)
    remark = "TP opted for: "
    if len(df['Method other than VAN'][0]) > 2 and "SMTP" not in df['Method other than VAN'][0]:
        connections = df['Method other than VAN'][0].split(',')
    elif len(df['Method other than VAN'][0]) > 2 and "SMTP" in df['Method other than VAN'][0]:
        connections = df['Method other than VAN'][0].split(',')
        remark += "SMTP, "
        for x in connections:
            if "SMTP" in x:
                connections.pop(x)
    else:
        connections =[]
    vanConn = df['Value Added Network Provider (VAN) '][0]
    if vanConn:
        remark+="(VAN) "+vanConn+","
    files = []
    for data in connections:
        names = quest_map[data]
        remark += data+", "
        for name in names:
            file_path = "C:\\Users\\RajnishKumarVENDORRo\\PycharmProjects\\CDS\Test files\\Questionaire\\" + name
            files.append(file_path)
    print(files)
    attachment = it.takewhile(lambda x: os.path.exists(x), files)
    return attachment,remark


print(surveyAnalysis())