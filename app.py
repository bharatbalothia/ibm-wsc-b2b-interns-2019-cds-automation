from flask import Flask, render_template, flash, redirect, request, url_for, send_file, session
from werkzeug.utils import secure_filename
import csv
import os, uuid
import pythoncom
from handleEmail import Sendmail
import itertools as it
from handleDoc import surveyParser,QuestionCheck


app = Flask(__name__)


@app.route('/massMail')
def massMail():
    return render_template("Mailbox.html")


@app.route('/')
def index():
    return redirect(url_for('massMail'))

@app.route('/sendMail', methods=['GET', 'POST'])
def sendMail():
    senders_list = request.files['tp-contact']
    senders_location = "C:\\Users\\RajnishKumarVENDORRo\\PycharmProjects\\CDSApplication\\SendersList\\" + secure_filename(senders_list.filename)
    senders_list.save(senders_location)
    attachment_file = request.files.getlist('attachment[]')
    attachment_files = []
    if attachment_file:
        for files in attachment_file:
            attachment_location = "C:\\Users\\RajnishKumarVENDORRo\\PycharmProjects\\CDSApplication\\AttachmentList\\" + secure_filename(files.filename)
            senders_list.save(attachment_location)
            attachment_files.append(attachment_location)
    cc = request.form.get('cc')
    subject = request.form.get('subject')
    message = request.form.get('message')
    csvfile1 = open(senders_location, 'rt')
    reader1 = csv.DictReader(csvfile1)
    k = Sendmail.Sendmail()
    for row in reader1:
        to = row['Email']
        subjects="[IBM EMEA CDS] -"+row['Customer Name'] +subject +"-" + row['Name']
        attachment = it.takewhile(lambda x: os.path.exists(x), attachment_files)
        k.send_mail(subjects, message, to, attach=attachment)
    flash('Mail Sent', 'info')
    return redirect(url_for('massMail'))

@app.route('/sur_par')
def sur_par():
    return render_template("surveyparser.html")

@app.route('/que_check')
def que_check():
    return render_template("QuestionCheck.html")


@app.route('/detailsExtract', methods=['GET', 'POST'])
def detailsExtract():
    f = request.files['survey-file']
    survey_location = "C:\\Users\\RajnishKumarVENDORRo\\PycharmProjects\\CDSApplication\\tempDocs\\" + secure_filename(f.filename)
    f.save(survey_location)
    info = surveyParser.parse(survey_location)
    #flash('Mail Sent', 'info')
    return render_template("surveyparser.html",data=info)

@app.route('/questionCheck', methods=['GET', 'POST'])
def questionCheck():
    f = request.files['question-file']
    survey_location = "C:\\Users\\RajnishKumarVENDORRo\\PycharmProjects\\CDSApplication\\tempDocs\\" + secure_filename(f.filename)
    f.save(survey_location)
    l = [x for x in QuestionCheck.questcheck(survey_location)]

    #flash('Mail Sent', 'info')
    return render_template("QuestionCheck.html",data=l)

if __name__ == '__main__':
    app.secret_key = 'secret123'
    app.run(debug=True)
