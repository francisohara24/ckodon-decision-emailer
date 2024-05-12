"""Script for validating email addresses"""
import re
import pandas as pd

# import student data
PG800_UG700A = pd.read_excel("./data/PG800+UG700A.xlsx")
PG884_UG700B = pd.read_excel("./data/PG884+UG700B.xlsx")
UG675 = pd.read_excel("./data/UG675.xlsx")

# verify imported data
print("PG800+UG700A:\n", PG800_UG700A.head())
print("PG884_UG700B:\n", PG884_UG700B.head())
print("UG675:\n", UG675.head())


def validate(email: str)-> bool:
    """Returns a boolean indicating whether the given email is a valid email address.

    Args:
        email: The email address to be validated.

    Returns:
        `True` if the given email is a valid email address. Otherwise, returns `False`.
    """
    regex = "[a-zA-Z0-9.\-+_]+@[a-zA-Z0-9.\-+_]+[a-zA-Z]{2,}"  # regex pattern for validating address
    email = email.strip()  # remove leading and trailing whitespace
    if re.search(regex, email):
        return True
    return False


for email in students['Email']:
    print(validate(email), email)
