# Getting started
```
# import
from appscript_msg import Message

# instantiate Message
msg = Message()
```

# Creating email
## Arguments
* subject
* body
* to_recipients
* cc_recipients
* attachments
* send_type

# Simple Example
```
# create message and preview
msg.create_email(subject = 'Example Email', 
                 body = 'Hello,<br><br>This is an example email!<br><br>Thanks!',
                 to_recipients = 'chang@purdue.edu')
```

