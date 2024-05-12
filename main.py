"""Script I wrote to test how to send HTML email messages."""
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from smtplib import SMTP
import json

# extract smtp client credentials
client_credentials = json.loads(open("./2024-bootcamp/data/credentials.json").read())
server_address = client_credentials["server_address"]
port = client_credentials["port"]
username = client_credentials["username"]
password = client_credentials["password"]

# create smtp client
smtp = SMTP(server_address, port)
smtp.ehlo()
smtp.starttls()
smtp.login(username, password)

# add send delay
counter = 0  # no. of emails sent
delay_s = 60  # delay in seconds

# construct html email message
message = MIMEMultipart()
message["From"] = "ckodontech@gmail.com"
message["To"] = "franciskohara@gmail.com"
message["Subject"] = "Welcome to Ckodon"
text = """\
<html>
<head>
</head>
<body>
<div align="center">
<img src="https://github.com/francisohara24/ckodon-decision-emailer/blob/master/2024-bootcamp/data/logo.jpg?raw=true">
</div>
<div>
<p style="font-size:15px">
Hello Mentee,<br><br>

This is a test email.<br><br>

Best,<br>
The Ckodon Foundation.
</p>
</div>
</body>
</html>
"""
content = MIMEText(text, "html")
message.attach(content)

# send html email message
try:
    smtp.send_message(message)
    print("Success!")
except:
    print("Failed!")

