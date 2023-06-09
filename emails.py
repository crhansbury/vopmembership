from email.message import EmailMessage
import smtplib
import os
import sys
from querymember import query_member_attr


def send_email(reciever_email, subject, body):
    """Sends an email to the email defined in the 'reciever_email' parameter.
    Sends from vopmembershiptest@gmail.com. Password is defined in a local
    environment variable and cannot be retrieved outside of the test environment
    at this time."""
    sender_email = "vopmembershiptest@gmail.com"
    password = os.environ.get("VOP_PASSWORD")

    # Create an instance of the EmailMessage class to generate the email
    em = EmailMessage()
    em["From"] = sender_email
    em["To"] = reciever_email
    em["Subject"] = subject
    em.set_content(body)

    # Log in to sender_email and define SMTP server
    smtp_server = "smtp.gmail.com"
    smtp_port = 587  # SMTP over TLS

    # Connect to email server using TLS and log in to email
    server = smtplib.SMTP(smtp_server, smtp_port)
    server.ehlo()
    server.starttls()
    server.login(sender_email, password)

    # Send the email
    try:
        server.sendmail(sender_email, reciever_email, em.as_string())
    except:
        print(f"Email to {reciever_email} unsuccessful.")
    server.quit()
    print(f"✅ Email successfully sent to {reciever_email} with the subject: {subject}.")


def generate_email(template, receiver_email, spreadsheet):
    """Returns two values: (subject, email_body). Fills in an existing template
    for an email with all of the memeber's attributes, their Section Leader, and
    the Section Leader's email.

    The subject is defined in the template as the first line followed by an
    empty line. The body includes the third line of the template to the end.

    The variables in the template include all attributes of the member
    associated with the 'receiver_email' parameter passed into the function as
    defined in the spreadsheet file passed into the 'spreadsheet' parameter.
    The variables must be written in the template as: {first_name}, {last_name},
    {pronouns}, {section}, {id}, {email}, {phone}, {address}, {city}, {state},
    {zipcode}. Other variables inclue the section, name, and email of the
    Section Leader associated with the member's section. These variables must
    be written in the template as: {sec_first}, {sec_last}, and {sec_email}."""

    # Retrieving the attributes from the spreadsheet for the member
    member_attributes = query_member_attr(
        spreadsheet,
        "email",
        receiver_email,
        "first_name",
        "last_name",
        "pronouns",
        "section",
        "id",
        "phone",
        "address",
        "city",
        "state",
        "zip",
        "role",
    )
    # Turning the items in the list of attributes into strings
    mattr_string = [str(i) for i in member_attributes[0]]
    # Retrieving attributes for all Section Leaders from the spreadshet
    sec_leader_attributes = query_member_attr(
        spreadsheet,
        "role",
        "Section Leader",
        "section",
        "first_name",
        "last_name",
        "email",
    )
    # Finding the Section Leader for the member's section
    section_leader = []
    for list in sec_leader_attributes:
        new_section = mattr_string[3]
        if list[0].startswith(new_section[0]):
            section_leader.extend(list)
    # Prevent an error if modifying the section leader's role attribute
    if len(section_leader) < 1:
        section_leader = ["None", "NO SECTION LEADER FOUND", " ", "NO EMAIL AVAILABLE"]
    # Open the email template and replace all variables with the correct info
    with open(template) as f:
        subject = f.readline().strip().format(first_name=mattr_string[0])
        # Testing to ensure the template is formatted correctly to return an
        # appropriate subject line
        if len(subject) > 70:
            print(
                f"{template} is not formatted correctly. Please ensure that "
                "the email subject is the first line, followed by an empty "
                "line, followed by the email body. The subject can be no "
                "greater than 70 characters long."
            )
            sys.exit(3)
        elif subject == "":
            print(
                f"{template} is not formatted correctly. Please ensure that "
                "the email subject is the first line, followed by an empty "
                "line, followed by the email body. The subject can be no "
                "greater than 70 characters long."
            )
            sys.exit(3)
        next(f)
        email_template = f.read()
    # Appending the template text to fill in the values
    email_body = email_template.format(
        first_name=mattr_string[0],
        last_name=mattr_string[1],
        pronouns=mattr_string[2],
        section=mattr_string[3],
        id=mattr_string[4],
        email=receiver_email,
        phone=mattr_string[5],
        address=mattr_string[6],
        city=mattr_string[7],
        state=mattr_string[8],
        zipcode=mattr_string[9],
        role=mattr_string[10],
        sec_first=section_leader[1],
        sec_last=section_leader[2],
        sec_email=section_leader[3],
    )
    return subject, email_body
