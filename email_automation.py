from exchangelib import *
import csv
import speech_recognition as sr
import os
import smtplib
from email.mime.text import MIMEText
from email.header import Header
from datetime import *

def loadSysParameter(email_sysFile):
    # Load csv file into sysParameter
    with open(email_sysFile, newline="") as csvfile:
        sysFileReader = csv.reader(csvfile, delimiter=',')
        sysParameter = {}
        for row in sysFileReader:
            if sysFileReader.line_num == 2:
                receivers = [ele for ele in row[1:] if ele != ""]
                sysParameter[row[0]] = receivers
            else:
                for ele in row:
                    if ele != "":
                        sysParameter[row[0]] = ele
    return sysParameter



def ObtainAudio(username,password,server,primary_smtp_address):
    creds = Credentials(
        username=username,  # Or myusername@example.com for O365
        password=password
    )
    config = Configuration(server=server, credentials=creds)

    account = Account(
        primary_smtp_address=primary_smtp_address,
        autodiscover=False,
        config=config,
        access_type=DELEGATE
    )

    filtered_items = account.inbox.all().filter(subject__startswith='New Voicemail')
    for i, item in enumerate(filtered_items):
        if i == 0:
            targetemail = item
            print(targetemail.subject)
            break

    for attachment in targetemail.attachments:
        with open(attachment.name, 'wb') as f:
            f.write(attachment.content)
        return attachment.name

def SendEmail(text, sender, receivers, sender_name, receiver_name, subject, ip_address, port):
    smtpIPaddr = ip_address
    smtpPort = port
    # first for content, second for plain, third for setting code
    message = MIMEText(text+".\nverified", 'plain','utf-8')
    message['From'] = Header(sender_name, 'utf-8')
    message['To'] = Header(receiver_name, 'utf-8')
    message['Subject'] = Header(subject, 'utf-8')

    try:
        smtpObj = smtplib.SMTP(smtpIPaddr, smtpPort)
        smtpObj.sendmail(sender, receivers, message.as_string())
        print("mail complete")
        smtpObj.close()
    except smtplib.SMTPException:
        print("email sent error")

def RecongizeSound(audio):
    r = sr.Recognizer()
    with sr.AudioFile(audio) as source:
        audio_data = r.record(source)
        # recognize (convert from speech to text)
        text = r.recognize_google(audio_data)
    return text

def IsNormal(text):
    print(date.today().day)
    if str(date.today().day) in text:
        print('Correct date!')
        return True
    else:
        return False

def RemoveFile(filename):
    if os.path.isfile(filename):
        os.remove(filename)

if __name__ == '__main__':
    file = 'email_sysFile.csv'
    sysParameter = loadSysParameter(file)
    sender, receivers, sender_name, receiver_name, subject, username, password, server, primary_smtp_address, ip_address, port = sysParameter.values()

    attachment_name = ObtainAudio(username, password, server, primary_smtp_address) #download audio file from email and return attachment name
    text = RecongizeSound(attachment_name)
    if IsNormal(text):
        print(text)
        SendEmail(text, sender, receivers, sender_name, receiver_name, subject, ip_address, port)
    else:
        SendEmail('system operation error', sender, receivers, sender_name, receiver_name, subject, ip_address, port)
    RemoveFile(attachment_name)


