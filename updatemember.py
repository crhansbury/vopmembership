from openpyxl import load_workbook
from openpyxl.styles import Border, Side, Alignment
from datetime import datetime
from querymember import query_member_object
from emails import send_email, generate_email

def update_member(file, email):
    """Updates the member information in the spreadsheet. Prompts user for:
     name, pronouns, voice part, role, email address, phone number, or 
      mailing address."""
    # Open the file and iterate over the emails field to find the desired row
    workbook = load_workbook(file)
    sheet = workbook["Active Members"]
    edit_row = 0
    best_email = email
    for row_num, row in enumerate(sheet.iter_rows(min_row=2, min_col=1, 
                                                  max_col=12, values_only=True), 
                                                  start=2):
        if email == row[11]:
            edit_row = row_num
            member_first = row[1]
            member_last = row[0]
    if edit_row == 0:
        print("{} not found in {}.".format(email, file))
    else:
        check = input("Edit information for {} {}? [y/n]\n".format(member_first, member_last))
        if check.lower().strip() == "y":
            while True:
                print("Please select which field you would like to edit\n",
                      "[1] First Name\n",
                      "[2] Last Name\n",
                      "[3] Pronouns\n",
                      "[4] Section\n",
                      "[5] Email\n",
                      "[6] Phone number\n",
                      "[7] Address\n",
                      "[8] Role")
                response = input("(enter a number from the menu above)\n")
                if response == "1":
                    first_name = input("Please enter new first name: ").title()
                    sheet.cell(row=edit_row, column=2, value=first_name)
                    member_first = first_name
                    # Prompt user for another edit
                    again = input("Edit another field? [y/n]\n").lower().strip()
                    if again == "y":
                        continue
                    else:
                        break
                if response == "2":
                    last_name = input("Please enter new last name: ").title()
                    sheet.cell(row=edit_row, column=1, value=last_name)
                    member_last = last_name
                    # Prompt user for another edit
                    again = input("Edit another field? [y/n]\n").lower().strip()
                    if again == "y":
                        continue
                    else:
                        break
                if response == "3":
                    pronouns = input("Please enter new pronouns: ").lower()
                    sheet.cell(row=edit_row, column=3, value=pronouns)
                    # Prompt user for another edit
                    again = input("Edit another field? [y/n]\n").lower().strip()
                    if again == "y":
                        continue
                    else:
                        break
                if response == "4":
                    section = input("Please enter new section: ").upper()
                    sheet.cell(row=edit_row, column=4, value=section)
                    # Prompt user for another edit
                    again = input("Edit another field? [y/n]\n").lower().strip()
                    if again == "y":
                        continue
                    else:
                        break
                if response == "5":
                    new_email = input("Please enter new email: ")
                    sheet.cell(row=edit_row, column=12, value=new_email)
                    best_email = new_email
                    # Prompt user for another edit
                    again = input("Edit another field? [y/n]\n").lower().strip()
                    if again == "y":
                        continue
                    else:
                        break
                if response == "6":
                    phone = input("Please enter new phone number: ")
                    sheet.cell(row=edit_row, column=11, value=phone)
                    # Prompt user for another edit
                    again = input("Edit another field? [y/n]\n").lower().strip()
                    if again == "y":
                        continue
                    else:
                        break
                if response == "7":
                    street_address = input("Please enter new street address: ").title()
                    sheet.cell(row=edit_row, column=7, value=street_address)
                    city = input("Please enter new city: ").title()
                    sheet.cell(row=edit_row, column=8, value=city)
                    state = input("Please enter new state: ").upper()
                    sheet.cell(row=edit_row, column=9, value=state)
                    zipcode = input("Please enter new zip code: ")
                    sheet.cell(row=edit_row, column=10, value=zipcode)
                    # Prompt user for another edit
                    again = input("Edit another field? [y/n]\n").lower().strip()
                    if again == "y":
                        continue
                    else:
                        break
                if response == "8":
                    role = input("Please enter new role: ").title()
                    sheet.cell(row=edit_row, column=6, value=role)
                    # Prompt user for another edit
                    again = input("Edit another field? [y/n]\n").lower().strip()
                    if again == "y":
                        continue
                    else:
                        break
                else:
                    print("Please enter a valid response from the menu.")
                    continue
            # Adding the date modified
            date = datetime.now()
            sheet["M1"] = "Date Modified: {} at {}".format(date.strftime("%m/%d/%Y"),
                                                   date.strftime("%H:%M"))
            workbook.save(file)
            print("{} {} successfully edited.".format(member_first, member_last))
            subject, body = generate_email("email-templates/updated_member_template.txt",
                                           best_email, file)
            send_email(best_email, subject, body)

        
update_member("vopmembership_data 5.xlsx", "vopmembershiptest+chanes@gmail.com")