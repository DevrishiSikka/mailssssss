import smtplib
from smtplib import SMTPDataError, SMTPAuthenticationError, SMTPException
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email.mime.multipart import MIMEMultipart
from email import encoders
import openpyxl


#### HELPER FUNCTIONS #####

class style:
    BOLD = '\033[1m'
    END = '\033[0m'


##########################


sponsorEmail = []
sponsorFile = '150_emails.xlsx'
workbookObj = openpyxl.load_workbook(sponsorFile)
sheetObj = workbookObj.active
for row in sheetObj.iter_rows():
    sponsorEmail.append(row[1].value)
del sponsorEmail[0]

senderEmail = 'fyivitc@gmail.com'
senderPass = 'aykfxtqmkgexkyap'

message = MIMEMultipart('alternative')
message['Subject'] = " FYI’s event InFYInite |A Pre-Vibrance extravaganza"
message['From'] = senderEmail

messageHTML = open("infyinite_marketing.html", "r").read()
HTMLpart = MIMEText(messageHTML, 'html')
message.attach(HTMLpart)

# SET ATTACHMENT
fileName = 'infyinite_poster.png'
fileAtt = open(fileName, 'rb')
messageAttachment = MIMEBase('application', 'octet-stream')
messageAttachment.set_payload(fileAtt.read())
encoders.encode_base64(messageAttachment)
messageAttachment.add_header('Content-Disposition', f"attachment; filename= {fileName}")
message.attach(messageAttachment)
counter = 1
for emailId in sponsorEmail:
    try:
        s = smtplib.SMTP('smtp.gmail.com', 587)
        s.starttls()
        s.login(senderEmail, senderPass)
        s.sendmail(senderEmail, emailId, message.as_string())
        print(f"{counter} -> Mail sent succesfully to {emailId}.")
        counter = counter + 1
        s.quit()
    except SMTPDataError as e:
        print(f"The SMTP server refused to accept the message data for {style.BOLD + emailId + style.END}")
    except SMTPAuthenticationError as e:
        print("Incorrect Username and/or Password")
    except SMTPException as e:
        print("Some unknown error occured")
    except Exception as e:
        print(f"Some error occured at {emailId}")
