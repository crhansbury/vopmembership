from querymember import query_active_member, query_inactive_member

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
    new_id = str(max_number + 1)
    return new_id
