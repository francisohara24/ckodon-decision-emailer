"""Script for validating applicant email addresses."""
from re import search


def validate(email_address):
    pattern = r"^[A-Za-z0-9_\-.+]+@[a-zA-Z0-9.\-_+]+[A-Za-z]{2,}$"
    return bool(search(pattern, email_address))
