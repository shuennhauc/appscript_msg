# https://stackoverflow.com/questions/61529817/automate-outlook-on-mac-with-python
# https://appscript.sourceforge.io/py-appscript/index.html


from appscript import app, k #pip install
from mactypes import Alias
from pathlib import Path
import os
import pandas as pd #pip install

class Message:

    def __init__(self, subject, body, to_recipients=[], cc_recipients=[], attachments=None, send_type='Show'):
        self.outlook = app('Microsoft Outlook')

        self.create_msg(subject, body)

        self.add_recipient_list(to_recipients, type_ = 'to')
        self.add_recipient_list(cc_recipients, type_ = 'cc')
        if attachments != None:
            self.add_attachments(attachments)

        if send_type=='show':
            self.msg.open()
        elif send_type=='send':
            self.msg.send()
        else:
            raise ValueError('send_type only accepts "show" or "send"')

    
    def create_msg(self, subject, body):
        self.msg = self.outlook.make(
            new=k.outgoing_message,
            with_properties={
                k.subject: subject,
                k.content: body,
                })

    def add_recipient_list(self, email_list, type_ = 'to'):
        if not isinstance(email_list, list): 
            email_list = [email_list]

        for email in email_list:
            self.add_recipient(email, type_)
        

    def add_recipient(self,email, type_ = 'to'):
        msg = self.msg

        if type_ == 'to':
            new_recipient = k.to_recipient
        elif type_ == 'cc':
            new_recipient = k.cc_recipient

        msg.make(
                new=new_recipient,
                with_properties={
                        k.email_address: {k.address: email}
                        }
            )

    def add_attachments(self, attachments):
        msg = self.msg

        if not isinstance(attachments, list):
                attachments = [attachments]

        for attachment_path in attachments:

            msg.make(new=k.attachment,
                with_properties={
                    k.file: Alias(str(attachment_path))})
