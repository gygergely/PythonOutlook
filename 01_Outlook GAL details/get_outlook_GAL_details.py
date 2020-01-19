import win32com.client
import csv
from datetime import datetime

# Outlook
outApp = win32com.client.gencache.EnsureDispatch("Outlook.Application")
outGAL = outApp.Session.GetGlobalAddressList()
entries = outGAL.AddressEntries

# Create a dateID
date_id = (datetime.today()).strftime('%Y%m%d')

# Create empty list to store results
data_set = list()

# Iterate through Outlook address entries
for entry in entries:
    if entry.Type == "EX":
        user = entry.GetExchangeUser()
        if user is not None:
            if len(user.FirstName) > 0 and len(user.LastName) > 0:
                row = list()
                row.append(date_id)
                row.append(user.Name)
                row.append(user.FirstName)
                row.append(user.LastName)
                row.append(user.JobTitle)
                row.append(user.City)
                row.append(user.PrimarySmtpAddress)
                try:
                    row.append(
                        entry.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/proptag/0x3a26001e"))
                except:
                    row.append('None')

                # Store the user details in data_set
                data_set.append(row)

# Print out the result to a csv with headers
with open(date_id + 'outlookGALresults.csv', 'w', newline='', encoding='utf-8') as csv_file:
    headers = ['DateID', 'DisplayName', 'FirstName', 'LastName', 'JobTitle', 'City', 'PrimarySmtp', 'Country']
    wr = csv.writer(csv_file, delimiter=',')
    wr.writerow(headers)
    for line in data_set:
        wr.writerow(line)