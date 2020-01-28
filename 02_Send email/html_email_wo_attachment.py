import win32com.client

# Hard coded email subject
MAIL_SUBJECT = 'AUTOMATED Text Python Email without attachments'

# Hard coded email HTML text
MAIL_BODY =\
    '<html> ' \
    '   <body>' \
    '       <p><b>Dear</b> Receipient,<br><br>'\
    '       This is an automatically generated email by <font size="5" color="blue">Python.</font><br>' \
    '       It is so <del>amazing and</del> fantastic<br>' \
    '       <strong>Wish you</strong> all the <font size="5" color="green">best</font><br>'\
    '   </body>' \
    '</html>'


def send_outlook_html_mail(recipients, subject='No Subject', body='Blank', send_or_display='Display', copies=None):
    """
    Send an Outlook HTMLText email
    :param recipients: list of recipients' email addresses (list object)
    :param subject: subject of the email
    :param body: HTML body of the email
    :param send_or_display: Send - send email automatically | Display - email gets created user have to click Send
    :param copies: list of CCs' email addresses
    :return: None
    """
    if len(recipients) > 0 and isinstance(recipient_list, list):
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
        ol_msg.HTMLBody = body

        if send_or_display.upper() == 'SEND':
            ol_msg.Send()
        else:
            ol_msg.Display()
    else:
        print('Recipient email address - NOT FOUND')


if __name__ == '__main__':

    recipient_list = ['recipient1@someemaildomain.com',
                      'recipient2@someemaildomain.com',
                      'recipient3@someemaildomain.com']

    copies_list = ['cc1@someemaildomain.com']

    send_outlook_html_mail(recipients=recipient_list, subject=MAIL_SUBJECT, body=MAIL_BODY, send_or_display='Display',
                           copies=copies_list)
