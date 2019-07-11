from flask import Flask, render_template, flash, redirect, request, url_for, send_file, session
from werkzeug.utils import secure_filename
import csv
import os, uuid
from handleEmail import Sendmail
import itertools as it
from functools import wraps
import datetime
from DataHandling import DataHandling
import pandas as pd
from authenticate import authenticate
import couchdb

dao = DataHandling()

app = Flask(__name__)



@app.before_request
def make_session_permanent():
    session.permanent = True
    app.permanent_session_lifetime = datetime.timedelta(minutes=20)


def is_logged_in(f):
    @wraps(f)
    def wrap(*args, **kwargs):
        if 'logged_in' in session:
            return f(*args, **kwargs)
        else:
            flash('Unauthorized, Please login', 'danger')
            return redirect(url_for('login'))

    return wrap


@app.route('/login', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        # Get Form Fields
        POST_USERNAME = str(request.form['username'])
        POST_PASSWORD = str(request.form['password'])
        if "robin" == POST_USERNAME:
            session['logged_in'] = True
            session['username'] = POST_USERNAME

            # flash('You are now logged in', 'success')
            return redirect(url_for('index'))

        isLoged = authenticate(POST_USERNAME,POST_PASSWORD)
        if not isLoged[0] :
            session['authorized'] = 0
            error = 'Invalid login'
            return render_template('login.html', error=error)
        else:
            session['logged_in'] = True
            session['username'] = isLoged[1]

            # flash('You are now logged in', 'success')
            return redirect(url_for('index'))
    return render_template('login.html')


@app.route('/massMail')
@is_logged_in
def massMail():
    return render_template("Mailbox.html")


@app.route('/')
def index():
    if 'logged_in' in session and session['logged_in'] == True:
        return render_template('Mailbox.html')
    return redirect(url_for('login'))

@app.route('/sendMail', methods=['GET', 'POST'])
@is_logged_in
def sendMail():
    senders_list = request.files['tp-contact']
    senders_location = "C:\\Users\\RajnishKumarVENDORRo\\PycharmProjects\\CDSApplication\\SendersList\\" + secure_filename(senders_list.filename)
    senders_list.save(senders_location)
    attachment_file = request.files.getlist('attachment[]')
    attachment_files = []
    Attachment_type = ""
    if attachment_file:
        for files in attachment_file:
            attachment_location = "C:\\Users\\RajnishKumarVENDORRo\\PycharmProjects\\CDSApplication\\AttachmentList\\" + secure_filename(files.filename)
            files.save(attachment_location)
            attachment_files.append(attachment_location)
            if 'survey' in files.filename.lower():
                Attachment_type= "Survey Sent"
    cc = request.form.get('cc')
    bcc = request.form.get('bcc')

    subject = request.form.get('subject')
    message = request.form.get('message')
    csvfile1 = open(senders_location, 'rt')
    reader1 = csv.DictReader(csvfile1)
    k = Sendmail.Sendmail()
    for row in reader1:
        to = row['Email']
        subjects="[IBM EMEA CDS] -"+row['Customer Name'] +"-"+subject +"-" + row['Trading Partner Name']
        attachment = it.takewhile(lambda x: os.path.exists(x), attachment_files)
        k.send_mail(subjects, message, to, cc, bcc, attach=attachment)
        tpDetail = {
            "TP name": row['Trading Partner Name'],
            "Date": datetime.datetime.now().strftime("%Y-%m-%d %H:%M"),
            "Survey Sent": "YES",
            "Questionnaire Sent": "NO",
            "Status" : Attachment_type
        }
        customerDetails = {
                "_id" : row['Customer Name'],
                "Customer Name" : row['Customer Name'],
                "Subject" : subjects,
                "First Date": datetime.datetime.now().strftime("%Y-%m-%d %H:%M"),
                "TPlist" : {}
            }
        dao.saveAnyData(customerDetails,tpDetail)
    flash('Mail Sent', 'info')
    return redirect(url_for('massMail'))


@app.route('/reportPage')
@is_logged_in
def reportPage():
    Customers = dao.getCustomerList()
    session['Customer List'] = Customers
    return render_template("ReportPage.html",data=Customers)

@app.route('/getReports',methods=['GET', 'POST'])
@is_logged_in
def getReports():
    selectedField = request.form.get('selectcustomer')
    tpdata = dao.getReport(selectedField)
    print(tpdata)
    reportData = pd.DataFrame(tpdata)
    del reportData['TP ID']
    session['TP Data'] = tpdata
    return render_template("ReportPage.html",data=session['Customer List'],dftables=reportData.to_html(classes=['table', 'table-hover', 'table-bordered']))

@app.route('/TPReport')
@is_logged_in
def TPReport():
    server = couchdb.Server("http://%s:%s@9.199.145.193:5984/" % ("admin", "admin123"))
    db = server['cdstpdata']
    Customers = []
    for id in db:
        Customers.append(id)
    session['New Customer List'] = Customers
    return render_template("TPDetailsPage.html",data=Customers)

@app.route('/getTPDetails',methods=['GET', 'POST'])
@is_logged_in
def getTPDetails():
    selectedField = request.form.get('selectcustomer')
    server = couchdb.Server("http://%s:%s@9.199.145.193:5984/" % ("admin", "admin123"))
    db = server['cdstpdata']
    tpdata = db[selectedField]['TPDetails']
    reportData = pd.DataFrame(tpdata)
    session['New TP Data'] = tpdata
    return render_template("TPDetailsPage.html",data=session['New Customer List'],dftables=reportData.to_html(classes=['table', 'table-hover', 'table-bordered']))

@app.route('/downloadTpDetails',methods=['GET', 'POST'])
@is_logged_in
def downloadTpDetails():
    filename = "C:\\Users\\RajnishKumarVENDORRo\\PycharmProjects\\CDSApplication\\TPDetailReport.csv"
    tpdata = session['New TP Data']
    reportData = pd.DataFrame(tpdata)
    reportData.to_csv(filename, sep=',', encoding='utf-8')
    return send_file(filename,as_attachment='demo.csv')


@app.route('/downloadReports',methods=['GET', 'POST'])
@is_logged_in
def downloadReports():
    filename = "C:\\Users\\RajnishKumarVENDORRo\\PycharmProjects\\CDSApplication\\reportFile.csv"
    tpdata = session['TP Data']
    reportData = pd.DataFrame(tpdata)
    del reportData['TP ID']
    reportData.to_csv(filename, sep=',', encoding='utf-8')
    return send_file(filename,as_attachment='demo.csv')

'''
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
'''


@app.route('/logout')
@is_logged_in
def logout():
    session.clear()
    session['logged_in'] = False
    flash('You have logged out', 'success')
    return redirect(url_for('login'))


if __name__ == '__main__':
    app.secret_key = 'secret123'
    app.config['SESSION_TYPE'] = 'filesystem'
    app.run(host='0.0.0.0',port=5000,debug=True)
