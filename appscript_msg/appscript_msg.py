# https://stackoverflow.com/questions/61529817/automate-outlook-on-mac-with-python
# https://appscript.sourceforge.io/py-appscript/index.html


from appscript import app, k 
from mactypes import Alias
from time import sleep

class Message:

    def __init__(self):
        self.outlook = app('Microsoft Outlook')
    
    def create_email(self, subject, body, to_recipients=[], cc_recipients=[], attachments=None, send_type='show', pause_confirm=True, send_delay=1):      
        """
        Args:
            subject (str): the email subject.
            body (str): the email body.
            to_recipient (list): list of email address to send
            cc_recipient (list): list of email address to cc
            attachments (list): list absolute path to attachments
            send_type (str): Defaults to 'show'. 'show' will open the email message for preview; 'send' will send without preview
            
        """
        self.create_msg(subject, body)

        self.add_recipient_list(to_recipients, type_ = 'to')
        self.add_recipient_list(cc_recipients, type_ = 'cc')
        if attachments != None:
            self.add_attachments(attachments)

        if send_type=='show':
            self.msg.open()
            
            if not isinstance(pause_confirm, bool):
                raise ValueError("pause_confirm only accepts booleans")
            
            if pause_confirm:
                input('Press any key to continue')
                
            
        elif send_type=='send':
            self.msg.send()
                   
            if not isinstance(send_delay, float):
                raise ValueError("Assign send_delay as a float")
            
            sleep(send_delay)
            
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
