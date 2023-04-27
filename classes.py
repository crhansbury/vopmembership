from dataclasses import dataclass

# Defines a Member data class. Each member in the spreadsheet will be
# an instance of the Member class, with each column of information
# mapping to an attribute of the class.

@dataclass
class Member:
    id: str
    first_name: str
    last_name: str
    pronouns: str
    section: str
    role: str
    email: str
    address: str
    city: str
    state: str
    zip: str
    phone: str
