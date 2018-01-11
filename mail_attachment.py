import imaplib
import poplib
from django.core.management import BaseCommand
import logging
import email
import tempfile
import os
import re
import xlrd
import shutil
import traceback
import collections
from django.core.mail import get_connection, send_mail
from django.template.loader import get_template
from django.template import Context
import datetime
from django.core.mail import send_mail
from django.core.mail import EmailMessage

PROJECT_ROOT = os.path.abspath(os.path.dirname(__name__))
root_path = PROJECT_ROOT.split('root_directory')[0]
upload_path = os.path.join(root_path,'static/uploads/file_directory/')


class MailAttachement(BaseCommand):
    def __init__(self):
        super(MailAttachement, self).__init__()
        # Connect to imap server
        self.username = 'abc@outlook.com'
        self.password = 'xyz'
        self.sender_email_id = []

    def handle(self, *args, **options):
        mail = imaplib.IMAP4_SSL('outlook.office365.com')
        mail.login(self.username, self.password)
        status ,data = mail.select()
        if status == 'OK':
            logging.info(" ************************START************************\n")
            logging.info(" Processing Inbox mailbox... [%s] : \n" % datetime.datetime.now())
            r = self.process_inbox_mailbox(mail)
            mail.close()
            return r
        else:
            logging.info(" ERROR: Unable to open Inbox mailbox, status %s time %s" % (status, datetime.datetime.now()))
            return 0,'false'

        mail.logout()

    def process_inbox_mailbox(self, mail):
        err_dict = {}
        status, data = mail.search(None,'(UNSEEN)')
        if status != 'OK':
            logging.info(" No messages found!!! [%s] : \n" % datetime.datetime.now())
            return 0,'false'                                                                                                                                                                                                                                                                        print 'dataaa',data[0]
        if data[0]:
            for msg_id in data[0].split():
                try:
                    r = self.process_mail(mail, msg_id)
                except Exception, e:
                    logging.info("Unable to process mail %s %s" % (e, traceback.format_exc()))
                    return 0,'false'
            return r
        else:
            logging.info(" No messages found!!! [%s] : \n" % datetime.datetime.now())
            return 0,'false'


    def process_mail(self, mail, msg_id):
        err_dict =dict()
        status, data = mail.fetch(msg_id, '(RFC822)')
        if status != 'OK':
            logging.info("ERROR getting message id %s time %s" % (msg_id, datetime.datetime.now()))
            return 0,'false'
        msg = email.message_from_string(data[0][1])
        sender_name = msg['from'].split(' ')[0]
        sender_email = msg['from'].split('<')[-1].split('>')[0]
        email_list = ['abc@outlook.com']
        if str(sender_email) in email_list:
            self.sender_email_id.append(sender_email)
            attachment_files = [part.get_filename() for part in msg.walk()]
            for part in attachment_files:
                if part is not None and part.split('.')[-1] in ['CSV','csv']:
                    flag = 1
                    break
                else:
                    flag = 0
            if not flag:
                err_list = "No csv Attachment Found. Please attach file in .csv or .CSV format and send it."
                logging.info(" No CSV Attachment Found. Mail sent to %s. [%s] \n"%(sender_email,datetime.datetime.now()))
                self.email_processing(err_list, err_dict, sender_email, sender_name)
                return 0,'false'
            else:
                for part in msg.walk():
                    filename = part.get_filename()
                    if filename:
                        if filename.split('.')[-1] not in ['CSV','csv']:
                            continue
                        else:
                            file_path = os.path.join(upload_path, filename)
                            if not os.path.isfile(file_path):
                                fp = open(file_path, 'wb')
                                fp.write(part.get_payload(decode=True))
                            fp = open(file_path, 'r')
                            attachments = [(filename, fp.read())]
                            fp.close()
                return 1,sender_email
        else:
            logging.info(" Not processing mail of %s [%s] : \n" %(sender_email,datetime.datetime.now()))
            return 0,'false'


    def email_processing(self, err_list, err_dict, sender_email_id, sender_name, attachment, success_list=None):
        #print sender_email_id
        email_list = ['abc@outlook.com']
        if sender_email_id in email_list:
            sender_email_id = [sender_email_id]
            context = err_list
            subject = 'Response: file'
            EMAIL_HOST='outlook.office365.com'
            EMAIL_PORT=587
            SMTPSecure = 'tls'
            EMAIL_HOST_USER = "abc@outlook.com"
            EMAIL_HOST_PASSWORD = "xyz"
            EMAIL_USE_TLS = True
            connection = get_connection(host=EMAIL_HOST,port=EMAIL_PORT,username=EMAIL_HOST_USER,password=EMAIL_HOST_PASSWORD,use_tls=EMAIL_USE_TLS) 
            email = EmailMessage(subject, context,'abc@outlook.com',sender_email_id,connection=connection)
            email.attach_file(attachment)
            email.send(fail_silently=False)
            return 'true'

