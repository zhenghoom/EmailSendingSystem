from flask import Flask, render_template, request, url_for, redirect
from time import sleep
import win32com.client as client
import pythoncom
import csv

app = Flask(__name__, static_url_path="/static")

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/mass_send', methods=['POST'])
def saveValue():
    file = request.form['file']
    subject = request.form['subject']
    content = request.form['content']
    chunks = request.form['num-chunks']
    time = request.form['time-interval']
    attachment = request.form['attachment']
    attachments = attachment.split(", ")

    pythoncom.CoInitialize()
    with open(file, newline='') as lines:
        reader = csv.reader(lines)
        individual = [row for row in reader]
    # Emails to send out every cycle
    chunks = [individual[x:x + int(chunks)] for x in range(0, len(individual), int(chunks))]

    template = content
    outlook = client.Dispatch("Outlook.Application")
    if request.form['submitType'] == 'Submit':
        if attachment == '':
            for chunk in chunks:
                for name, email in chunk:
                    message = outlook.CreateItem(0)
                    message.Display()
                    message.To = email
                    message.Subject = subject
                    message.HTMLBody = template.format(name)
                # sleep(int(time))
        else:
            for chunk in chunks:
                for name, email in chunk:
                    message = outlook.CreateItem(0)
                    message.Display()
                    message.To = email
                    message.Subject = subject
                    message.HTMLBody = template.format(name)
                    for i in range(len(attachments)):
                        message.Attachments.Add(attachments[i])
                # sleep(int(time))
        return render_template('loading.html'), {"Refresh": "4; url=mass_send"}

    elif request.form['submitType'] == 'Save':
        if attachment == '':
            for chunk in chunks:
                for name, email in chunk:
                    message = outlook.CreateItem(0)
                    message.To = email
                    message.Subject = subject
                    message.HTMLBody = template.format(name)
                    message.Save()
                # sleep(int(time))
        else:
            for chunk in chunks:
                for name, email in chunk:
                    message = outlook.CreateItem(0)
                    message.To = email
                    message.Subject = subject
                    message.HTMLBody = template.format(name)
                    for i in range(len(attachments)):
                        message.Attachments.Add(attachments[i])
                    message.Save()
                # sleep(int(time))
        return render_template('loading.html'), {"Refresh": "4; url=mass_send"}

@app.route('/mass_send')
def massSend():
    return render_template('massSend.html')

@app.route('/excel')
def excelTemp():
    return render_template('excelTemp.html')

@app.route('/loading')
def submitLoading():
    return render_template('submitLoading.html'),{"Refresh": "1; url=/"}

if __name__ == '__main__':
    app.run(debug=True)
