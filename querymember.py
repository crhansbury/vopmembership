from vopmembership import create_active_members

def query_member(file, attribute):
    """Searches the list of Member objects for a certain attribute, returning
    a list of the attributes."""
    active_members = create_active_members(file)
    member_list = []
    for member in active_members:
        attr_value = getattr(member, attribute)
        member_list.append(attr_value)
    return member_list
    
value = query_member("vopmembership_data.xlsx", "first_name")
print(value)


    

# Could use this to look up the members in the dictionary with the other
# functions as well, such as updatemember and move member.