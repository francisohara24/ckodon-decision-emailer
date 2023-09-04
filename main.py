import pandas as pd
from smtplib import SMTP
from email.message import EmailMessage
import json
from email_validator import validate


# function for sending graduate decisions
def grad_emailer():
    smtp = start_smtp()  # instantiate smtp client
    graduate_students = pd.read_excel("./data/Graduate Pairing.xlsx")  # import graduate student data
    for row in graduate_students.index:
        name = graduate_students.loc[row]["Full Name"]
        email_address = graduate_students.loc[row]["Email"]

        # create email message
        msg = EmailMessage()
        msg["From"] = "ckodontech@gmail.com"
        msg["To"] = email_address
        msg["Subject"] = "Ckodon Mentorship Program Update"
        body = f"""Dear {name}, 

Congratulations!
You've been selected to join the Ckodon 2023 Graduate Application Mentorship Program.
Join the following WhatsApp group to receive further instructions about your next steps and to confirm your mentor:

https://chat.whatsapp.com/DAeyCTPMirl65UKc54rhvc

DO NOT SHARE this link with anyone.

Best,
The Ckodon Foundation Team."""
        msg.set_content(body)

        # send email message
        try:
            smtp.send_message(msg)
            print("SENT")
        except:
            print(row + 2, name, email_address)

# function for sending undergraduate decisions
def undergrad_emailer():
    pass


# instantiate SMTP client
def start_smtp():
    client_credentials = json.loads(open("./data/credentials.json").read())
    hostname = client_credentials["hostname"]
    port = client_credentials["port"]
    username = client_credentials["username"]
    password = client_credentials["password"]
    smtp = SMTP(hostname, port)
    smtp.starttls()
    smtp.login(username, password)
    return smtp


grad_emailer()

# for email_address in graduate_student_emails:
#     # validate email addresses
#     status = validate(email_address.strip())
#     if not status:
#         print(status, email_address)  # print all invalid email addresses
    # grad_emailer(email_address, smtp)
