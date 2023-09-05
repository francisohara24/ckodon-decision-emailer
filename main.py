import pandas as pd
from smtplib import SMTP
from email.message import EmailMessage
import json
from time import sleep
from email_validator import validate


# function for sending graduate decisions
def grad_emailer():
    # instantiate smtp client
    smtp = start_smtp()

    # import graduate student data
    graduate_students = pd.read_excel("./data/Graduate Pairing.xlsx")

    # validate email addresses
    graduate_emails = graduate_students["Email"]
    for email in graduate_emails:
        status = validate(email.strip())
        if not status:
            print("Invalid:", email)

    wait_counter = 0  # count no. of emails till waiting time.

    # obtain whatsapp link from credentials.json
    credentials = json.loads(open("./data/credentials.json").read())
    grad_whatsapp_link = credentials["grad_whatsapp_link"]

    # create email message
    for row in graduate_students.index:
        name = graduate_students.loc[row]["Full Name"]
        email_address = graduate_students.loc[row]["Email"]

        msg = EmailMessage()
        msg["From"] = "ckodontech@gmail.com"
        msg["To"] = email_address
        msg["Subject"] = "Ckodon Mentorship Program Update"
        body = f"""Dear {name}, 

Congratulations!
You've been selected to join the Ckodon 2023 Graduate Application Mentorship Program.
Join the following WhatsApp group to receive further instructions about your next steps and to confirm your mentor:

{grad_whatsapp_link}

DO NOT SHARE this link with anyone.

Best,
The Ckodon Foundation Team."""
        msg.set_content(body)

        # send email message
        try:
            smtp.send_message(msg)
            print("SENT")
            wait_counter += 1
        except Exception as exception:
            print(exception)
            print(row + 2, name, email_address)
            break

        # wait 1 minute every 50 emails
        if wait_counter == 50:
            wait_counter = 0
            wait()

# function for sending undergraduate decisions
def undergrad_emailer():
    smtp = start_smtp()  # instantiate smtp client

    # import student data
    undergraduate_students = pd.read_excel("./data/Undergraduate Pairing.xlsx")

    # validate email addresses
    undergrad_emails = undergraduate_students["Email"]
    for email in undergrad_emails:
        status = validate(email.strip())
        if not status:
            print("Invalid:", email)

    # obtain whatsapp link from credentials.json
    credentials = json.loads(open("./data/credentials.json").read())
    undergrad_whatsapp_link = credentials["undergrad_whatsapp_link"]

    wait_counter = 0  # count no. of emails sent till waiting time.

    # create email message
    for row in undergraduate_students.index:
        name = undergraduate_students.loc[row]["Full Name"]
        email_address = undergraduate_students.loc[row]["Email"]
        msg = EmailMessage()
        msg["From"] = "ckodontech@gmail.com"
        msg["To"] = email_address
        msg["Subject"] = "Ckodon Mentorship Program Update"
        body = f"""Dear {name},

Congratulations!
You've been selected to join the Ckodon 2023 Graduate Application Mentorship Program.
Your mentorship group will be focused on building your undergraduate profile in order to prepare you to put forth a competitive application during the next application season.
Join the following WhatsApp group to learn more about your next steps and connect with your mentor:

{undergrad_whatsapp_link}

DO NOT SHARE this link with anyone.

Best,
The Ckodon Foundation Team.
"""
        msg.set_content(body)

        # send email message
        try:
            smtp.send_message(msg)
            print("SENT")
            wait_counter += 1
        except Exception as exception:
            print(exception)
            print(row + 2, name, email_address)
            break

        # wait 1 minute for every 50 emails
        if wait_counter == 50:
            wait_counter = 0
            wait()

# function for instantiating SMTP client
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


# function for waiting 1 minute
def wait():
    duration_s = 60  # wait duration in seconds
    print(f"-----Waiting for {duration_s} seconds-----")
    sleep(duration_s)


# grad_emailer()
# undergrad_emailer()
