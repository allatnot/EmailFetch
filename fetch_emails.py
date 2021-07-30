import sys
import imaplib, email
import os
from datetime import date
import io
class MailConnection:
    def __init__(self, host, username, password,folder = 'Inbox',
                 subject=None,sub_folder = None):
        '''
        This script is intended to download attachments from the email and save into the prferred directory
        host = email server
        username = email address
        password = email password
        folder = email folder
        delete = 'y' if delete email after read or any other letter in no
        subject = email subject header if any
        sub_folder = child folder if any
        '''
        self.host = host
        self.username = username
        self.password = password
        self.folder = folder
        self.sub_folder = sub_folder
        self.subject = subject
        self.connections = imaplib.IMAP4_SSL(host)
        self.connections.login(username, password)
        if sub_folder:
            if subject:
                self.connections.select('{}/{}'.format(folder,sub_folder))
                self.rks, self.email_items = self.connections.search(None,'(SUBJECT "{}")'.format(subject))
            if subject is None:
                self.rks, email_items = self.connections.select('{}/{}'.format(folder,sub_folder))
                lst = []
                for i in range(1,int(email_items[0])+1):
                    lst.append(i)
                sd = ' '.join(map(str, lst))
                self.email_items = sd.encode()
        if sub_folder is None:
            if subject:
                self.connections.select('{}'.format(folder))
                self.rks, self.email_items = self.connections.search(None,'(SUBJECT "{}")'.format(subject))
            if subject is None:
                self.rks, email_items = self.connections.select('{}'.format(folder))
                lst = []
                for i in range(1,int(email_items[0])):
                    lst.append(i)
                sd = ' '.join(map(str, lst))
                self.email_items = sd.encode()
    def get_email_attachment(self, directory, delete = 'y'):
        '''
        directory = name of the directory to save the emails
        '''
        if not os.path.exists(directory):
            #create a directory if not exists
            cons = input('The specified directory does not exists! Do you want to create one? Enter "y" for "yes" or any letter for "no" and hit Enter: ')
            cons = cons.upper()
            if cons == 'Y':
                os.makedirs(directory, exist_ok = True)
            else:
                print('Good bye!!!')
                sys.exit()
        connection = self.connections
        rks = self.rks
        email_items = self.email_items
        if type(email_items) is list:
            items = email_items[0].split()
        else:
            items = email_items.split()
        mails = []
        for il in items:
            resp, data = connection.fetch(il, "(RFC822)")
            email_body = data[0][1]
            mail = email.message_from_bytes(email_body)
            mails.append(mail)
        for ml in range(len(mails)):
            if mails[ml].get_content_maintype() != 'multipart':
                continue
            for part in mails[ml].walk():
                if part.get_content_maintype() == 'multipart':
                    continue
                if part.get('Content-Disposition') is None:
                    continue
                fil_name = part.get_filename()
                
                filename = 'Downloaded_on_{}_{}_{}'.format(date.today(),ml,fil_name)
                att_path = os.path.join(directory, filename)
                try:
                    fp = open(r'{}'.format(att_path), 'wb')
                except OSError:
                    fl_name = fil_name.split('.')
                    #print(fl_name)
                    try:
                        fl_name = fl_name[1].split('?')
                        ext = fl_name[0]
                    except IndexError:
                        ext = ' '
                    filename = 'downloaded_on_{}_{}_.{}'.format(date.today(), ml,ext)
                    att_path = os.path.join(directory, filename)
                    fp = open(r'{}'.format(att_path), 'wb')
                fp.write(part.get_payload(decode='utf-8'))
                fp.close()
        delete = delete.upper()
        if delete == 'Y':
            for il in items:
                conn.store(il, "+FLAGS", "\\Deleted")
            conn.expunge()
        connection.close()
        connection.logout()
        return
    def read_emails(self, directory, delete ='y'):
        if not os.path.exists(directory):
            cons = input('The specified directory does not exists! Do you want to create one? Enter "y" for "yes" or any letter for "no" and hit Enter: ')
            cons = cons.upper()
            if cons == 'Y':
                os.makedirs(directory, exist_ok = True)
            else:
                print('Good bye!!!')
                sys.exit()
        connection = self.connections
        rks = self.rks
        email_items = self.email_items
        if type(email_items) is list:
            items = email_items[0].split()
        else:
            items = email_items.split()
        mails = []
        for il in items:
            resp, data = connection.fetch(il, "(RFC822)")
            email_body = data[0][1]
            mail = email.message_from_bytes(email_body)
            mails.append(mail)
        for ml in range(len(mails)):
            if mails[ml].get_content_maintype() != 'multipart':
                for part in mails[ml].walk():
                    try:
                        file = open(directory+'/Email_{}_{}.txt'.format(mails[ml]['Subject'], ml), 'w')
                    except OSError:
                        file = open(directory+'/{}_{}.txt'.format('Email', ml), 'w')
                    subject = mails[ml]['Subject']
                    file.write(subject)
                    line = part.get_payload(decode='utf-8')
                    try:
                        line = line.decode()
                    except:
                        line = str(line)
                    file.write(line)
                    file.close()
            else:
                for part in mails[ml].walk():
                    if part.get_content_maintype() == 'multipart':
                        continue
                    if part.get('Content-Disposition') is None:
                        continue
                    fil_name = part.get_filename()
                    subject = mails[ml]['Subject']
                    try:
                        fp = io.open(directory+'/Attachment_{}_{}.txt'.format(subject,ml),'w', encoding='utf-8')
                    except OSError:
                        fp = io.open(directory+'/Attachment_{}.txt'.format(ml), 'w', encoding='utf-8')
                    fp.write('Subject   =====   {}\n\n'.format(subject))
                    fp.write('Name of the attachment  ===   {}'.format(fil_name))
                    fp.close()
        delete = delete.upper()
        if delete == 'Y':
            for il in items:
                conn.store(il, "+FLAGS", "\\Deleted")
            conn.expunge()
        connection.close()
        connection.logout()
        return
