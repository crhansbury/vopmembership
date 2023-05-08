from openpyxl import load_workbook
from openpyxl.styles import Border, Side, Alignment
from datetime import datetime
from emails import generate_email, send_email
from newmember import create_member_id
from querymember import query_active_member, query_member_object

def add_member(file):
    """Adds a member to the spreadsheet. Takes criteria for:
    name, pronouns, voice part, role, email address, phone number, and 
    mailing address. Appends spreadsheet with this information. Updates
    a field that indicates the date of the latest update."""
    # Prompt the user for info
    print("Please enter the information for the new member.")
    first_name = input("First Name: ")
    last_name = input("Last Name: ")
    pronouns = input("Pronouns: ")
    section = input("Section: ")
    role = input("Role: ")
    email = input("Email: ")
    phone = input("Phone Number: ")
    address = input("Street Address: ")
    city = input("City: ")
    state = input("State: ")
    zipcode = input("Zip Code: ")
    # Generate new member ID
    try:
        member_id = create_member_id(file)
    except:
        print("{} not found, or is an invalid format. Please enter a valid filename.".format(file))
        exit()
    # Open the workbook and load the "Active Member" sheet and append info
    workbook = load_workbook(filename=file)
    sheet = workbook["Active Members"]
    row = (
        last_name,first_name,pronouns,section,member_id,role,address,city, \
        state,zipcode,phone,email
    )
    sheet.append(row)
    # Formatting the new row to match the rest of the spreadsheet
    thin_border = Border(left=Side(style='thin'),
                         right=Side(style='thin'),
                         top=Side(style='thin'),
                         bottom=Side(style='thin'))
    alignment = Alignment(horizontal='left', vertical='bottom', wrap_text=True)
    new_row = sheet.max_row
    for cell in sheet[new_row]:
        cell.border = thin_border
        cell.alignment = alignment
    # Adding the date modified
    date = datetime.now()
    sheet["M1"] = "Date Modified: {} at {}".format(date.strftime("%m/%d/%Y"),
                                                   date.strftime("%H:%M"))
    workbook.save(file)
    print("{} {} successfully added to {}.".format(first_name, last_name, file))
    # Generating and sending welcome email once new member has been successfully
    # appended to spreadsheet
    try:
        subject, body = generate_email("email-templates/new_member_template.txt", 
                                       email, file)
        send_email(email, subject, body)
    except:
        print("Email to {} was unsuccessful.".format(email))

def all_active_email(template, spreadsheet):
    """Sends an email to every member in the Active Members sheet of the 
    spreadsheet passed into the 'spreadsheet' parameter. Uses the email
    template passed into the 'template' parameter."""
    # Make a lsit of all emails from the Active Members sheet
    active_emails= query_active_member(spreadsheet, "email")
    # Send an email to each email in list
    for email in active_emails:
        subject, body = generate_email(template, email, spreadsheet)
        send_email(email, subject, body)

def inactive_member(file, attribute, attr_value):
    """Moves a member from the Active Member sheet to the Inactive Member 
    sheet."""
    member = query_member_object(file, attribute, attr_value)[0]
    # # Open the inactive sheet and append the member to the inactive sheet
    workbook = load_workbook(file)
    inactive_sheet = workbook["Inactive Members"]
    active_sheet = workbook["Active Members"]
    # Find the row that matches the email attribute
    active_sheet_row = 0
    for row_num, row in enumerate(active_sheet.iter_rows(min_row=2, min_col=12, 
                                                         max_col=12, 
                                                         values_only=True), 
                                                         start=2):
        if member.email == row[0]:
            active_sheet_row = row_num
    # Find the last row of inactive sheet
    inactive_sheet_row = inactive_sheet.max_row + 1
    # Copy/pasting all values from active sheet to inactive sheet
    for col_num, cell in enumerate(active_sheet[active_sheet_row], 1):
        value = cell.value
        inactive_sheet.cell(row=inactive_sheet_row, column=col_num, value=value)
    # Formatting the cells
    thin_border = Border(left=Side(style='thin'),
                         right=Side(style='thin'),
                         top=Side(style='thin'),
                         bottom=Side(style='thin'))
    alignment = Alignment(horizontal='left', vertical='bottom', wrap_text=True)
    for cell in inactive_sheet[inactive_sheet_row]:
        cell.border = thin_border
        cell.alignment = alignment
    # Delete the member from active member sheet
    active_sheet.delete_rows(idx=active_sheet_row)
    # Add date modified
    date = datetime.now()
    active_sheet["M1"] = "Date Modified: {} at {}".format(date.strftime("%m/%d/%Y"),
                                                   date.strftime("%H:%M"))
    workbook.save(file)
    print("{} successfully moved to Inactive Members.".format(member.first_name))
    # Send an automatic email when this completes
    subject, body = generate_email("email-templates/inactive_member_template.txt",
                                   member.email, file)
    send_email(member.email, subject, body)

def active_member(file, attribute, attr_value):
    """Moves a member from the Inactive Member sheet to the Active Member 
    sheet. Takes the argument of a member attribute and value, and looks
    up that member. Copies the member over from the Inactive sheet to the 
    Active sheet."""
    member = query_member_object(file, attribute, attr_value)[0]
    # Open the active sheet and append the member to the active sheet
    workbook = load_workbook(file)
    inactive_sheet = workbook["Inactive Members"]
    active_sheet = workbook["Active Members"]
    # Find the row that matches the email attribute
    inactive_sheet_row = 0
    for row_num, row in enumerate(inactive_sheet.iter_rows(min_row=2, min_col=12, 
                                                         max_col=12, 
                                                         values_only=True), 
                                                         start=2):
        if member.email == row[0]:
            inactive_sheet_row = row_num
    # Find the last row of inactive sheet
    active_sheet_row = active_sheet.max_row + 1
    # Copy/pasting all values from inactive sheet to active sheet
    for col_num, cell in enumerate(inactive_sheet[inactive_sheet_row], 1):
        value = cell.value
        active_sheet.cell(row=active_sheet_row, column=col_num, value=value)
    # Formatting the cells
    thin_border = Border(left=Side(style='thin'),
                         right=Side(style='thin'),
                         top=Side(style='thin'),
                         bottom=Side(style='thin'))
    alignment = Alignment(horizontal='left', vertical='bottom', wrap_text=True)
    for cell in active_sheet[active_sheet_row]:
        cell.border = thin_border
        cell.alignment = alignment
    # Delete member from inactive member sheet
    inactive_sheet.delete_rows(idx=inactive_sheet_row)
    # Add date modified
    date = datetime.now()
    active_sheet["M1"] = "Date Modified: {} at {}".format(date.strftime("%m/%d/%Y"),
                                                   date.strftime("%H:%M"))
    workbook.save(file)
    print("{} successfully moved to Active Members.".format(member.first_name))
    # Send an automatic email when this completes
    subject, body = generate_email("email-templates/active_member_template.txt",
                                   member.email, file)
    send_email(member.email, subject, body)

def search_member(file, attribute, attr_value):
    """Searches the spreadsheet for members that match the search criteria.
    Prints all the information for the members who match the search.
    Numbers all of the results if there are multiple members returned. Uses
    the query_member_object() function to search the file."""
    query_list = query_member_object(file, attribute, attr_value)
    member_count = 1    
    if len(query_list) > 1:
        print("Multiple members meet your search criteria.")
    for member in query_list:
        print("Result number {}:".format(member_count))
        print("Name: {} {}\n".format(member.first_name, member.last_name),
              "Pronouns: {}\n".format(member.pronouns),
              "Section: {}\n".format(member.section),
              "Member ID: {}\n".format(member.id),
              "Email: {}\n".format(member.email),
              "Phone number: {}\n".format(member.phone),
              "Address: {address}, {city} {state} {zip}".format(address=member.address,
                                                                city=member.city,
                                                                state=member.state,
                                                                zip=member.zip))
        member_count += 1


def main(file):
    """Prompts user for input. Prints out a main menu, asking
    if user would like to enter a new member into the spreadsheet, 
    change member information, 
    make a member inactive, or query the spreadsheet for information.
    Takes a command line argument of the file from which to read and edit."""
    while True:
        print("------------------------------------------")
        print("Welcome to the VOP Membership Portal!\n" \
              "What would you like to do?")
        print("[1] Add new members\n" \
            "[2] Remove members from Active Members\n" \
            "[3] Reinstate members to Active Members\n" \
            "[4] Update existing members\n" \
            "[5] Search for existing members\n" \
            "[6] Create a nametag for an existing member\n" \
            "[7] Create a label sheet for an existing member\n" \
            "[8] Send an email to all active members\n" \
            "[9] Exit")
        print("------------------------------------------")
        response = input("Please enter a number from the main menu:\n")
        if response == "9":
            break
        elif response == "1":
            # Double check that the user made the right choice
            check = input("You have selected 'Add a new member'. Proceed?\n" \
                          "(enter 'y' to continue, or 'n' to go back to main menu):\n")
            if check.lower() == "y":
                while True:
                    add_member(file)
                    # Check if the user wishes to add another member, or go back
                    # to main menu
                    again = input("Would you like to add another member? [y/n]\n")
                    if again.lower() == "y":
                        continue
                    else:
                        break
            else:
                continue
        elif response == "2":
            check = input("You have selected 'Remove a member from Active " \
                          "Members'. Proceed? [y/n]\n")
            if check.lower().strip() == "y":
                while True:
                    while True:
                        print("Please choose one of the following types of "\
                            "information for the member you wish you remove:\n"\
                            "[1] Email\n"\
                            "[2] Member ID")
                        response = input("")
                        attribute = ""
                        if response == "1":
                            attribute = "email"
                            break
                        elif response == "2":
                            attribute = "id"
                            break
                        else:
                            print("Please select a valid item from the menu.")
                            continue
                    attr_value = input("Please enter the member's {}:\n".format(attribute))
                    try:
                        inactive_member(file, attribute, attr_value)
                    except: 
                        print("{} is not a valid {}.".format(attr_value, attribute))
                        continue
                    again = input("Would you like to remove another member? [y/n]\n")
                    if again.lower().strip() == "y":
                        continue
                    else:
                        break
        elif response == "3":
            check = input("You have selected 'Reinstate a member to Active "\
                          "Members'. Proceed? [y/n]\n")
            if check.lower().strip() == "y":
                while True:
                    while True:
                        print("Please choose one of the following types of " 
                              "information for the member you wish to " \
                              "reinstate:\n" \
                              "[1] Email\n"\
                              "[2] Member ID")
                        response = input("")
                        attribute = ""
                        if response == "1":
                            attribute = "email"
                            break
                        elif response == "2":
                            attribute = "id"
                            break
                        else:
                            print("Please select a valid item from the menu.")
                            continue
                    attr_value = input("Please enter the member's {}:\n".format(attribute))
                    try:
                        active_member(file, attribute, attr_value)
                    except: 
                        print("{} is not a valid {}.".format(attr_value, attribute))
                        continue
                    again = input("Would you like to reinstate another member? [y/n]\n")
                    if again.lower().strip() == "y":
                        continue
                    else:
                        break
        elif response == "4":
            continue
        elif response == "5":
            check = input("You have selected 'Search for an existing member'. "\
                          "Proceed? [y/n]\n")
            if check.lower().strip() == "y":
                while True:
                    print("Please choose the search criteria:\n",
                          "[1] First Name\n",
                          "[2] Last Name\n",
                          "[3] Pronouns\n",
                          "[4] Section\n",
                          "[5] Member ID\n",
                          "[6] Email\n",
                          "[7] Phone number\n",
                          "[8] Street Address\n",
                          "[9] City\n",
                          "[10] State\n",
                          "[11] Zip Code")
                    reply = input("(choose a number from the menu above)\n")
                    attribute = ""
                    if reply == "1":
                        attribute = "first_name"
                    elif reply == "2":
                        attribute = "last_name"
                    elif reply == "3":
                        attribute = "pronouns"
                    elif reply == "4":
                        attribute = "section"
                    elif reply == "5":
                        attribute = "id"
                    elif reply == "6":
                        attribute = "email"
                    elif reply == "7":
                        attribute = "phone"
                    elif reply == "8":
                        attribute = "address"
                    elif reply == "9":
                        attribute = "city"
                    elif reply == "10":
                        attribute = "state"
                    elif reply == "11":
                        attribute = "zip"
                    else:
                        print("Please enter a valid number from the menu.")
                        continue
                    attr_value = input("Please enter the member's {}:\n".format(attribute))
                    search_member(file, attribute, attr_value)
                    again = input("Would you like to perform another search? [y/n]\n")
                    if again.lower().strip() == "y":
                        continue
                    else:
                        break
        elif response == "6":
            continue
        elif response == "7":
            continue
        elif response == "8":
            check = input("You have selected 'Send an email to all active " \
                          "members'. Proceed? [y/n]\n")
            if check == "y":
                template = input("Please enter the email template file: \n")
                try:
                    all_active_email(template, file)
                    continue
                except:
                    print("invalid file")
                    continue
            else:
                continue       
        else:
            print("Not a valid response. Please enter a number from the main menu.")
            continue
            
if __name__ == '__main__':
    main("vopmembership_data 4.xlsx")