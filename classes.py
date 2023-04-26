from dataclasses import dataclass

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
