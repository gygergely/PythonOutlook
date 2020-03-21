
import win32com.client

EMAIL_ACCOUNT = 'Enter your email address'  # e.g. 'good.employee@importantcompany.com'
ITER_FOLDER = 'Enter the Outlook folder name which emails you would like to iterate through'  # e.g. 'IterationFolder'
MOVE_TO_FOLDER = 'Enter the Outlook folder name where you move the processed emails'  # e.g 'ProcessedFolder'
SAVE_AS_PATH = 'Enter the path where to dowload attachments'  # e.g.r'C:\DownloadedCSV'
EMAIL_SUBJ_SEARCH_STRING = 'Enter the sub-string to search in the email subject'  # e.g. 'Email to download'


def find_download_csv_in_outlook():
    out_app = win32com.client.gencache.EnsureDispatch("Outlook.Application")
    out_namespace = out_app.GetNamespace("MAPI")

    out_iter_folder = out_namespace.Folders[EMAIL_ACCOUNT].Folders[ITER_FOLDER]
    out_move_to_folder = out_namespace.Folders[EMAIL_ACCOUNT].Folders[MOVE_TO_FOLDER]

    char_length_of_search_substring = len(EMAIL_SUBJ_SEARCH_STRING)

    # Count all items in the sub-folder
    item_count = out_iter_folder.Items.Count

    if out_iter_folder.Items.Count > 0:
        for i in range(item_count, 0, -1):
            message = out_iter_folder.Items[i]

            # Find only mail items and report, note, meeting etc items
            if '_MailItem' in str(type(message)):
                print(type(message))
                if message.Subject[0:char_length_of_search_substring] == EMAIL_SUBJ_SEARCH_STRING \
                        and message.Attachments.Count > 0:
                    for attachment in message.Attachments:
                        if attachment.FileName[-3:] == 'csv':
                            attachment.SaveAsFile(SAVE_AS_PATH + '\\' + attachment.FileName)
                            message.Move(out_move_to_folder)
    else:
        print("No items found in: {}".format(ITER_FOLDER))


if __name__ == '__main__':
    find_download_csv_in_outlook()
