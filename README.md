VOP Membership Management Application
===

## Overview

VOP is a local choir whose membership has grown rapidly in the last few years, from 30 members in 2019 to over 80 in 2023. The choir is in need of a membership management application that has the ability to scale with the rapid growth of the choir, while also working within a tightly constrained budget. The goal of this application is to provide an interface for membership registration and upkeep. This includes keeping track of important member information, suspending members on leave, and generating a unique member ID. The application should also make it easy to change any of that information. Finally, the application should automate the process of generating important onboarding documents or emails whenever a member is added, changed, or removed, as well as sending out an information change request each semester to ensure that all current member information is up-to-date.

## Case Scenario

Every semester there are ~20 new members who need to be entered into the system, assigned a member ID, and sent all the documents necessary for orientation. Additionally, there are ~10 current members who go on leave who need to be suspended from the database. Finally, every current member not on leave must review and update their member information as necessary at the start of each semester.

This information will be used to populate important choir documents. For example: nametags for each member with name, section, and pronouns; email lists based on section and role; concert programs with name and section; and statistics collection for grant applications.

### Spreadsheet Structure

The data will be kept in an excel spreadsheet. This spreadsheet will have two sheets, one for active members and one for inactive members. Both sheets will have 8 columns:
* Name (ex: First, Last)
* Member ID (ex: 1)
* Pronouns (ex: they/them)
* Voice Part (ex: B1)
* Role (ex: Section Leader)
* Email (ex: user@example.com)
* Phone (ex: (215) 555-1234)
* Address (ex: 123 Main St, Philadelphia PA 00000)

The program will use `openpyxl` to collect and edit the information in this spreadsheet. Using the `Data Class`, the program will parse the information in the xl sheet and use it to populate a dictionary, which will allow the program to collect the information from the excel spreadsheet.

The spreadsheet will also include a "date modified" field.

### Member Addition

When a new member joins VOP, there is information that must be collected about this new member to be stored in the membership database:
* Name (ex: First, Last)
* Pronouns (ex: they/them)
* Voice Part (ex: B1)
* Role (ex: Section Leader)
* Email (ex: user@example.com)
* Phone (ex: (215) 555-1234)
* Address (ex: 123 Main St, Philadelphia PA 00000)

The goal of this application is to make entering the information for each new member a quick and easy process by prompting the user for each of these fields and checking for the validity of each entry. A stretch goal for this application is to have this registration form sent out to each new member for them to fill out on their own.

### Member ID

Under the current conditions, in order to keep track of music, each member is currently hand-assigned a "folder number", which is then hand-written on each folder and distributed one-by-one to each member during the first few rehearsals. This process takes a long time, and is very prone to user error. Everyone's numbers are liable to change every semester, and the membership leaders are keeping the score of which numbers are in use or not in their heads. 

The goal of this application is to generate a unique member ID to each new member when they are entered into the database. Once assigned, the member ID will never be used again, even after a member is suspended. This member ID will be stored in the database and used to print out a label sheet every semester when new music is distributed. This label sheet will have the member ID as well as the name and section of each member on it. This will make it easy for each member of the choir to label their own music when it is recieved, and will allow for easy tracking of music when it is returned at the end of the semester.

### Member Suspension

Each semester a certain number of members leave the choir temporarily or permanently. Many members return after one or several semesters on leave and wish to be reinstated.

The goal of this application is to make the suspension of members going on leave simple by issuing one or two commands. This will be reflected in the spreadsheet by moving the Inactive Member from the Active Members sheet to the Inactive Members sheet, keeping the application from populating choir documents with suspended member information. This should make it easy to keep track of unique member IDs, as well as ease the transition back into active membership after going on leave for one or more semesters.

### Member Information Change

Personal information can change at any time, including address, name, pronouns, roles, and phone numbers. 

The goal of this application is to make changing membership information easy with user prompting. The member should be automatically sent an email when information changes, as well as new documentation depending on the change (ex: new nametag if there is a name change).

Additionally, an email should be sent to every current (not suspended) member at the beginning of each semester prompting each member to review their membership information and take action if there is out-of-date information. 

## Security

There are many security concerns around having permission to add or modify information in the membership database. To start, the application should be locked behind a username and password to be held only by the memberhsip administrator of the choir. This basic security measure will help keep confidential information from being accessed by unauthorized users, and will keep anyone from changing the information in the database without permission.

A future goal of this application is to create a user account for each new member of VOP, to allow them to add, modify, or suspend their own user accounts in the database to lighten the load of the membership administrator, while preventing individual users from accessing the confidential information inside the databse.

## Goal Outline

### Basic goals:
* get input on new member name, pronoun, section, role, email, phone, address
* generate a member ID
* output to a master spreadsheet of all members
    * 2 sheets: active members, inactive members
* be able to edit members in existing db
* print a nametag with name, section, and pronouns
* print a label sheet with member ID
* email new member welcome packet
* send out an email each semester prompting existing members for information updates
* lock application behind a username and password


### Future goals:
* GUI front-end for user input
* integrate with google api to output to google drive/sheets   
* accept user input with google forms to allow for registration individually
* allow each member to have their own account to access the database
    * each member will be able to add or modify their own information 
    * ensure that everyone can change their own info but no one else's, nor has access to other's confidential information

## Functions

* project
    * `main()`
        * Prompts user for input. Asks if user would like to enter a new member into the spreadsheet, change member information, make a member inactive, or query the spreadsheet for information.
* classes
    * `create_active_members()`
        * Creates a list using the information from the spreadsheet of all active members. Returns a list of each member on the spreadsheet as an instance of the Member dataclass, which includes all of the information from the columns in the spreadsheet.
    * `create_inactive_members()`
        * Creates a list using the information from the spreadsheet of all inactive members. Returns a list of each member on the spreadsheet as an instance of the Member dataclass, which includes all of the information from the columns in the spreadsheet.
* querymember
    * `query_active_member()`
        * Searches the list of active Member objects for a certain attribute, returning a list of all the attribute values for that attribute.
    * `query_inactive_member()`
        * Searches the list of inactive Member objects for a certain attribute, returning a list of all the attribute values for that attribute.
    * `query_member_attr()`
        * Looks for a specific attribute of a member, so one can search the file for someone's last name and return the other desired attribute. Example - query_member_attr('vopmembership_data.xlsx', first_name, 'Johnson', email) will return jwhite@domain.com. Takes four+ arguments - file name, the name of the attribute for the value being used to query the function, the value of the search, and as many return values as desired. Outputs a list of all the results that match the search, and each item of the list is a list of all the requested attributes for the individual Member instances.
* newmember
    * `add_member()`
        * Adds a member to the spreadsheet. Takes criteria for: name, pronouns, voice part, role, email address, phone number, and  mailing address. Appends spreadsheet with this information. Updates a field that indicates the date of the latest update.
    * `create_member_id()`
        * Creates a unique Member ID. Checks over every existing Member ID in spreadsheet (both active and inactive) and generates a new, sequential ID that does not already exist.
* emails
    * `send_email()`
        * Sends an email to a member in the spreadsheet. Planning to use for update confirmations, semesterly prompts for updates, new member welcome email.
    * `new_member_email_template()`
        * Fills in an existing template for a new member email with the new member's name, section, and section leader, creating a new file with the body of the email.
* updatemember
    * `update_member()`
        * Updates the member information in the spreadsheet. Prompts user for: name, pronouns, voice part, role, email address, phone number, or mailing address.
    * `inactive_member()`
        * Moves a member from the Active Member sheet to the Inactive Member sheet.
    * `active_member()`
        * Moves a member from the Inactive Member sheet to the Active Member sheet.
* material
    * `label_sheet()`
        * Creates a label sheet with the name and Member ID of a member.
    * `name_tag()`
        * Creates a nametag with the name, pronouns, and section of a member.

### Modules and Libraries

* `openpyxl`
* `dataclass`
* `datetime`

## Sources

Pregueiro, Pedro. "A Guide to Excel Spreadsheets in Python with openpyxl". https://realpython.com/openpyxl-excel-spreadsheets-python/#importing-data-from-a-spreadsheet. 2019, Aug 26.

The PyCoach. "How to Easily Automate Emails with Python." https://towardsdatascience.com/how-to-easily-automate-emails-with-python-8b476045c151. 2022, Jun 8.

The PyCoach. "How to Hide Passwords and Secret Keys in Your Python Scripts." https://medium.com/geekculture/how-to-hide-passwords-and-secret-keys-in-your-python-scripts-a8904d5560ec. 2022, May 26.

Unkonwn. "5 Examples of Write/Append Data to Excel Using openpyxl". https://www.excel-learn.com/append-data-openpyxl/#:~:text=For%20appending%20at%20the%20end,with%20our%20sample%20excel%20file.

OpenAI. ChatGPT. https://openai.com. 2023.

Grupetta, Stephen. "Using Python Optional Arguments When Defining Functions. https://realpython.com/python-optional-arguments/. 2021, Aug 30.

"How to get current date and time in Python?" https://www.programiz.com/python-programming/datetime/current-datetime#:~:text=If%20we%20need%20to%20get,class%20of%20the%20datetime%20module.&text=Here%2C%20we%20have%20used%20datetime,and%20time%20in%20another%20format. 