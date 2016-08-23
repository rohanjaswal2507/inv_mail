import smtplib, getpass
import email, openpyxl, sys
import email.mime.application

pb = 'Placement Brochure.pdf'
jnf = 'Job Notification Form - NIT Hamirpur.doc'
inv_letter = 'Invitation Letter.doc'

sender = 'rohanjaswal2507@gmail.com'

## E-mail content

msg = email.mime.Multipart.MIMEMultipart()
msg['Subject'] = 'Invitation for Campus Placement recruitment of B.Tech./B.Arch./M.Tech./M.Arch. 2017 batch students at NIT Hamirpur (HP)'

html_content = open('try.html', 'rb')
msg_text = html_content.read()


#attach inv_letter
filename=inv_letter
fp=open(filename,'rb')
att = email.mime.application.MIMEApplication(fp.read(),_subtype="doc")
fp.close()
att.add_header('Content-Disposition','attachment',filename=filename)
msg.attach(att)

#attach JNF
filename=jnf
fp=open(filename,'rb')
att = email.mime.application.MIMEApplication(fp.read(),_subtype="doc")
fp.close()
att.add_header('Content-Disposition','attachment',filename=filename)
msg.attach(att)

#attach placement brochure
filename=pb
fp=open(filename,'rb')
att = email.mime.application.MIMEApplication(fp.read(),_subtype="pdf")
fp.close()
att.add_header('Content-Disposition','attachment',filename=filename)
msg.attach(att)


## Get all the emails from text file or excelsheet
if len(sys.argv) < 2:
    print('Please provide the name of the excelsheet containing the e-mail IDs')
    exit()
else:
    wb_name = sys.argv[1]
    wb = openpyxl.load_workbook(wb_name)
    sheet = wb.active


n = int(raw_input('Enter the total number of e-mail IDs in the list:'))
print('The email IDs in your list are:')
for i in range(2, n+2):
    print(sheet['B' + str(i)].value)

confirmation = raw_input('Is that okay? (y/n)')
if confirmation == 'n':
    print('Please run the program with correct number of email addresses')
    exit()

print(' ')
## SMTP action
mailer = smtplib.SMTP('smtp.gmail.com', 587)
mailer.ehlo()
mailer.starttls()
password = getpass.getpass('Enter Password for ' + sender + ':')
not_logged_in = True

while not_logged_in:
    try:
        mailer.login(sender, password)
        not_logged_in = False
        print('Successfully logged in.')
        print('Now, sit back and relax. Your internet connection is slow. I will send all the e-mails')
    except Exception:
        password = getpass.getpass('Bro, please enter the password again: ')


## send e-mails to all in the excel sheet
for i in range (2, n+2):

    receiver = sheet['B' + str(i)].value
    msg['From'] = sender
    msg['To'] = receiver
    body = email.mime.Text.MIMEText(msg_text, 'html')
    msg.attach(body)

    try:
        mailer.sendmail(sender, receiver, msg.as_string())
        print('Mail successfully sent to ' + receiver)

    except Exception:
        print("Could not send mail to " + receiver)


mailer.quit()
