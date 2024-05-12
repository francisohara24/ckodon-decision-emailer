# import required modules
import pandas as pd
from functions import create_doc, create_email, create_smtp, send_email
from time import sleep
import json

# read google form data
students = pd.read_excel("data/Ckodon Bio Submission Form (Responses).xlsx")
students = students.tail(len(students) - 1265)  # select current batch of recipients

# extract smtp client credentials
client_credentials = json.loads(open("./data/credentials.json").read())
server_address = client_credentials["server_address"]
port = client_credentials["port"]
username = client_credentials["username"]
password = client_credentials["password"]

# create smtp client
smtp = create_smtp(server_address, port, username, password)

# add send delay
counter = 0  # no. of emails sent
delay_s = 60  # delay in seconds

# loop through all students
for row in students.index:
    # extract name, email, bio of student
    student = students.loc[row]
    student_name = student["Full Name"]
    student_email = student["Email"]
    student_major = student["Intended Major"]

    # create word document containing bio
    doc_content = f"""Name:   {student_name}
Email:   {student_email}
Major:   {student_major}
Bio:
        {student["Bio"]}
"""

    doc_name = f"Ckodon Bio [{student_name}].docx"
    doc_path = create_doc(doc_name, doc_content)

    # create email message
    msg_subject = "Ckodon Bio Update"
    msg_recipient = student_email
    msg_attachment = doc_path
    msg_content = f"""Dear {student_name},

Your Ckodon Bio Submission has been reviewed.
Attached to this email is a Word document containing your updated Bio.

Please do not reply to this email.

Best,
The Ckodon Foundation."""

    msg = create_email(msg_recipient, msg_subject, msg_content, msg_attachment)

    # send email message
    try:
        send_email(msg, smtp)
        counter += 1

        # Wait 1 minute for every 50 emails sent.
        if counter == 50:
            counter = 0
            print("---Waiting 1 minute---")
            sleep(delay_s)

    except Exception as Except:
        print(Except)
        print("Error sending mail at-")
        print("Row No:", row + 2)
        print("Student Name:", student_name)
        print("Student Email:", student_email)
        break
