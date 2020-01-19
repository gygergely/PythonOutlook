import win32com.client


def send_outlook_mail(recipients, subject='No Subject', body='Blank', send_or_display='Display', copies=None):

    outlook = win32com.client.Dispatch("Outlook.Application")

    ol_msg = outlook.CreateItem(0)

    str_to = ""
    for recipient in recipients:
        str_to += recipient + ";"

    ol_msg.To = str_to

    if copies is not None:
        str_cc = ""
        for cc in copies:
            str_cc += cc + ";"

        ol_msg.CC = str_cc

    ol_msg.Subject = subject
    ol_msg.Body = body

    if send_or_display.upper() == 'SEND':
        ol_msg.Send()
    else:
        ol_msg.Display()


if __name__ == '__main__':

    mail_subject = "AUTOMATED Text Python Email without attachments"

    mail_body = 'Dear Recipient\n\n'\
                'This is an automatically generated email by Python.\n\n'\
                'It is so amazing and fantastic\n\n'\
                'Wish you all the best\n\n'

    recipient_list = ['gyetvaigergely@gmail.com', 'gygergely1981@gmail.com', 'anita.vereb228@gmail.com']

    send_outlook_mail(recipients=recipient_list, subject=mail_subject)
