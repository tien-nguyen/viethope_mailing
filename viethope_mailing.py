import getpass
import smtplib
from email.MIMEMultipart import MIMEMultipart
from email.MIMEText import MIMEText
from email.MIMEBase import MIMEBase
from email import encoders
import argparse

import gspread

from oauth2client.service_account import ServiceAccountCredentials
import time
import datetime


class GoogleDoc(object):

    def __init__(self):
        # use creds to create a client to interact with the Google Drive API
        self.scope = ['https://spreadsheets.google.com/feeds']
        self.creds = ServiceAccountCredentials.from_json_keyfile_name(
            'client_secret.json', self.scope)
        self.client = gspread.authorize(self.creds)
        self.sheet = None

    def retrieveEmails(self):
        self.sheet = self.client.open("Viethope Email Address").sheet1

    def getNumberOfDonors(self):
        return len(self.sheet.get_all_records())

    def getSurNameAndEmail(self, i):
        return [self.sheet.cell(i, 1).value,
                self.sheet.cell(i, 4).value]

    def updateTimeStamp(self, i):
        now = datetime.datetime.fromtimestamp(
            time.time()).strftime('%Y-%m-%d %H:%M:%S')
        self.sheet.update_cell(i, 5, now)

    def close(self):
        self.client.close()


class Email(object):

    def __init__(self, from_email, subject):
        self.msg = MIMEMultipart()
        self.msg['From'] = from_email
        self.msg['Subject'] = subject

    def addRecipient(self, recipient_email):
        self.msg['To'] = recipient_email

    def addEmailContent(self, body):
        self.msg.attach(MIMEText(body, 'plain'))

    def addAttachment(self, filename, file_location):
        attachment = open(file_location, "rb")

        part = MIMEBase('application', 'octet-stream')
        part.set_payload((attachment).read())
        encoders.encode_base64(part)
        part.add_header('Content-Disposition',
                        "attachment; filename= %s" % filename)

        self.msg.attach(part)


def email(from_email, password, subject, body, attachment = None,
          attachment_location = None, test=True):

    google_doc = None
    if not test:
        google_doc = GoogleDoc()
        google_doc.retrieveEmails()

    server = smtplib.SMTP('smtp.gmail.com', 587)
    server.starttls()
    server.login(from_email, password)

    cc = from_email
    if not test:
        n = google_doc.getNumberOfDonors()
        for i in xrange(2, n + 2):
            list_ = google_doc.getSurNameAndEmail(i)
            email = Email(from_email, subject)
            if attachment is not None:
                email.addAttachment(attachment, attachment_location)

            email_body = body.format(list_[0])
            email.addEmailContent(email_body)
            server.sendmail(from_email, [list_[1], cc], email.msg.as_string())
            google_doc.updateTimeStamp(i)
            print("Sending email to {} @ {}".format(list_[0], list_[1]))
            time.sleep(5)

    else:
        body = body.format(" buddy, ")
        email.addEmailContent(body)
        server.sendmail(from_email, [from_email, cc],
                        email.msg.as_string())

    server.quit()


def message():
    message = """ Dear {}, \n
               Happy Holidays to all the friends of VNHelp! \n
               This year marks VNHelp's 25th Anniversary.
               The organization has been in operation since 1991 and has implemented many projects in Vietnam that have benefited many thousands of needy people. As you are celebrating the 2016 holidays, we are happy to share with you the wonderful results of some of our programs. To date, VNHelp has provided scholarships to 6,429 college students, built 49 schools, sponsored 283 orphans, connected 66,000 people to clean water, performed free cataract surgeries on 5,800 poor patients, and provided loans to 400 women. \n
               During this Season of Giving, we hope you remember the many ways VNHelp works in Vietnam to make lives better.  At this special time of the year, you can send a Christmas gift to your loved one by making a one-time donation to honor the person, or you can dedicate your contribution to a favorite VNHelp project. However you choose to donate to our work, we greatly appreciate your help, and all the beneficiaries in Vietnam are very grateful. Your financial support will enable VNHelp to continue our humanitarian mission in 2017 and beyond. We hope to hear from you soon. Thank you for your support! \n
               May this season bring you good health, happiness, and prosperity.

               Merry Christmas and Happy New Year!

               Sincerely,

               Thu Nguyen
               Viethope"""
    return message


if __name__ == "__main__":

    parser = argparse.ArgumentParser()

    parser.add_argument('-e', action='store',
                        help="Please type your email!", dest="email")
    # parser.add_argument('-p', action='store',
    #                    help="Please type your password!", dest="password")

    args = parser.parse_args()

    password = getpass.getpass("Password: ")

    SUBJECT = "VNHelp's 25th Season's Greetings [TESTING]"
    MESSAGE = message()

    email(args.email, password, SUBJECT,
          MESSAGE, attachment="Thiep moi.jpg",
          attachment_location=
          "/Users/tien/Dropbox/VietHope/mailing_script/cat.jpg",
          test=False)
