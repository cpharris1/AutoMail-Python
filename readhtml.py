import smtplib
import ssl
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

subj = "Test"
from_name = "Susan Lawson <susanlawson@lawsonmyside.com>"
receiver_email = "chloeh796@gmail.com"
email_str = "This is a test"
smtp_server = "mail.lawsonmyside.com"
port = 465
sender_email = "susanlawson@lawsonmyside.com"
pwd = "lC@2sQ^9r&C6Eg"


message = MIMEMultipart("alternative")
message["Subject"] = subj
message["From"] = from_name
message["To"] = receiver_email

html = email_str

part1 = MIMEText(html, "html")

message.attach(part1)

# create a secure SSL context
context = ssl.create_default_context()

try:
    with smtplib.SMTP_SSL(smtp_server, port, context=context) as server:
        server.login(sender_email, pwd)
        server.sendmail(
            sender_email, receiver_email, message.as_string()
        )
    print("Successfully sent email to " + receiver_email)

except smtplib.SMTPException as error:
    print("\nCould not send email to " + receiver_email)
    print("Error: " + str(error))