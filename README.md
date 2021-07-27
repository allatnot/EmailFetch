# EmailFetch
This script is designed to download attachments with or with out given email subject header or sub folder or preferred email folder.

## Descriptions
### The script first creates the MailConnection class object and get_email_attachment method with the following values
### MailConnection class
##### *rsd = MailConnection(host, username, password, folder = "Inbox", subject = None, sub_folder = None)*
#### host = name of the mail server
#### username = email address
#### password = email password
#### folder = email folder name default is Inbox
#### subject = email header subject default is None
#### sub_folder = name of the sub folder if any default is None
## Method get_email_attachment
##### *rsd.get_email_attachement(directory, delete = 'y')*
### The method needs a directory parameter and delete as optional parameter
#### directory = name of the directory where the downloaded attachment should be saved. If the directory does not exists the directory will be created if permitted by the user.
##### If the directory does not exist will be asked if the user wants to create it
#### delete = 'y' if delete email after read or any other letter in no
## Method to fetch emails and details about the attachments
##### *rsd.read_emails(directory, delete='y')*
