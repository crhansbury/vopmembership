from querymember import query_active_member, query_inactive_member
# from openpyxl import load_workbook
# from datetime import datetime
# from emails import new_member_email, send_email

# def add_member(file):
#     """Adds a member to the spreadsheet. Takes criteria for:
#     name, pronouns, voice part, role, email address, phone number, and 
#     mailing address. Appends spreadsheet with this information. Updates
#     a field that indicates the date of the latest update."""
#     # Prompt the user for info
#     first_name = input("First Name: ")
#     last_name = input("Last Name: ")
#     pronouns = input("Pronouns: ")
#     section = input("Section: ")
#     role = input("Role: ")
#     email = input("Email: ")
#     phone = input("Phone Number: ")
#     address = input("Street Address: ")
#     city = input("City: ")
#     state = input("State: ")
#     zipcode = input("Zip Code: ")
#     # Generate new member ID
#     member_id = create_member_id(file)
#     # Open the workbook and load the "Active Member" sheet and append info
#     workbook = load_workbook(filename=file)
#     sheet = workbook["Active Members"]
#     rows = (
#         last_name,first_name,pronouns,section,member_id,role,address,city, \
#         state,zipcode,phone,email
#     )
#     sheet.append(rows)
#     # Adding the date modified
#     date = datetime.now()
#     sheet["M1"] = "Date Modified: {} at {}".format(date.strftime("%m/%d/%Y"), \
#                                                    date.strftime("%H:%M"))
#     workbook.save(file)
#     print("{} {} successfully added to {}.".format(first_name, last_name, file))
#     # Generating and sending welcome email once new member has been successfully
#     # appended to spreadsheet
#     subject, body = new_member_email("email-templates/new_member_template.txt", email, file)
#     send_email(email, subject, body)


def create_member_id(file):
    """Creates a unique Member ID. Checks over every existing 
    Member ID in spreadsheet (both active and inactive) and generates a new, 
    sequential ID that does not already exist."""
    # Creating and concatenating the lists of member IDs from both active
    # and inactive member sheets
    active_member_id_list = query_active_member(file, "id")
    inactive_member_id_list = query_inactive_member(file, "id")
    id_list = [int(i) for i in active_member_id_list + inactive_member_id_list]
    # Assigning the new id number as one greater than the largest id in the
    # spreadsheet
    max_number = max(id_list)
    new_id = max_number + 1
    return new_id
