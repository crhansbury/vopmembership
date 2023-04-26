from vopmembership import create_active_members, create_inactive_members

def query_active_member(file, attribute):
    """Searches the list of active Member objects for a certain attribute, 
    returning a list of the attributes."""
    active_members = create_active_members(file)
    member_list = []
    for member in active_members:
        attr_value = getattr(member, attribute)
        member_list.append(attr_value)
    return member_list

def query_inactive_member(file, attribute):
    """Searches the list of inactive Member objects for a certain attribute, 
    returning a list of the attributes."""
    inactive_members = create_inactive_members(file)
    member_list = []
    for member in inactive_members:
        attr_value = getattr(member, attribute)
        member_list.append(attr_value)
    return member_list

    
value = query_inactive_member("vopmembership_data.xlsx", "pronouns")
print(value)


    

# Could use this to look up the members in the object list with the other
# functions as well, such as updatemember and move member.
# Could also use to iterate over the list of member ID numbers to make a new
# member ID that is unique and in sequence.