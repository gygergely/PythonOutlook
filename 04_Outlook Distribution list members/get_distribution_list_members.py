import win32com.client


# Outlook
outApp = win32com.client.gencache.EnsureDispatch("Outlook.Application")

# Get contact folder
contact_folder = outApp.Session.GetDefaultFolder(win32com.client.constants.olFolderContacts)

# Iterate through contacts
for contact in contact_folder.Items:
    # print(str(type(contact))[-15:-2])

    if str(type(contact))[-15:-2] == '_DistListItem':

        for i in range(1, contact.MemberCount + 1):
            member_in_dl = contact.GetMember(i)
            print('{} | {}'.format(contact.DLName, member_in_dl.Name))
    else:
        print('Not a distribution list')
