from classes import create_active_members, create_inactive_members
import sys 

def query_active_member(file, attribute):
    """Searches the list of active Member objects for a certain attribute, 
    returning a list of all the attribute values for that attribute."""
    active_members = create_active_members(file)
    member_list = []
    for member in active_members:
        attr_value = getattr(member, attribute)
        if attr_value != None:
            member_list.append(attr_value)
    return member_list

def query_inactive_member(file, attribute):
    """Searches the list of inactive Member objects for a certain attribute, 
    returning a list of all the attribute values for that attribute."""
    inactive_members = create_inactive_members(file)
    member_list = []
    for member in inactive_members:
        attr_value = getattr(member, attribute)
        if attr_value != None:
            member_list.append(attr_value)
    return member_list

def query_member_attr(file, attribute, attr_value, *return_attributes):
    """Looks for a specific attribute of a member, so one can search the file
    for someone's last name and return the other desired attribute. Example - 
    query_member_attr('vopmembership_data.xlsx', first_name, 'Johnson', email)
    will return jwhite@domain.com. Takes four+ arguments - file name, the 
    name of the attribute for the value being used to query the function, the
    value of the search, and as many return values as desired. Outputs a list
    of all the results that match the search, and each item of the list is
    a list of all the requested attributes for the individual Member instances."""
    member_list = create_active_members(file) + create_inactive_members(file)
    result = []
    # Iterate over each instance of Member class
    for member in member_list:
        # Get the queried attribute for each Member
        mem_attr_value = getattr(member, attribute)
        # Find the Member whose attr matches the search
        if mem_attr_value == attr_value:
            attr_list = []
            # Iterate over each desired return attribute parameter and append 
            # to the list of attributes for the matched Member
            for r_attr in return_attributes:
                attr_list.append((getattr(member, r_attr)))
            # Append the matched member's attrs to the result list
            result.append(attr_list)
    return result

def query_member_object(file, attribute, attr_value):
    """Queries the spreadsheet for a member object matching the specified 
    attribute and returns a list of the matching objects."""
    members = create_active_members(file) + create_inactive_members(file)
    member_list = []
    for member in members:
        try:
            matching_attr = getattr(member, attribute)
            if matching_attr == attr_value:
                member_list.append(member)
        except AttributeError:
            print(f"'{attribute}' is not a valid attribute.")
            sys.exit(15)
    if len(member_list) < 1:
        print(f"No {attribute} found with the value {attr_value}.")
    else:
        return member_list
