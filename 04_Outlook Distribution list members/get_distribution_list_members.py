import win32com.client

DL_NAME = 'Test12'

# Outlook
outApp = win32com.client.gencache.EnsureDispatch("Outlook.Application")

# Get contact folder
contact_folder = outApp.Session.GetDefaultFolder(win32com.client.constants.olFolderContacts)

# Iterate through contacts
for contact_item in contact_folder.Items:
    # check if item is a distribution list
    if contact_item.Class == win32com.client.constants.olDistributionList:
        # check if distribution list's name is equal to the constant
        if contact_item.DLName == DL_NAME:
            # loop through distribution list members and get their email address
            for i in range(1, contact_item.MemberCount + 1):
                member_in_dl = contact_item.GetMember(i)
                print('{} | {}'.format(contact_item.DLName, member_in_dl.Address))
