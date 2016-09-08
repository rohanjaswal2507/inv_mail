import smtplib, getpass
import email, sys, datetime
import email.mime.application

pb = 'Placement Brochure.pdf'
jnf = 'Job Notification Form - NIT Hamirpur.doc'
inv_letter = 'Invitation Letter.doc'
sender = 'rohanjaswal2507@gmail.com'

#Open log file to append the log
log_file = open('emails.log', 'a')

email_id = raw_input('Please enter the email ID:\n')

print('The email id entered by you is ' + email_id)
confirmation = raw_input('Is that okay? (y/n)')

if confirmation == 'n':
    print('Please run the program with correct number email addresses')
    exit()

print(' ')

## SMTP action
mailer = smtplib.SMTP('smtp.gmail.com', 587)
mailer.ehlo()
mailer.starttls()
password = getpass.getpass('Enter Password for ' + sender + ':')
not_logged_in = True


mailer.login(sender, password)
not_logged_in = False
print('Successfully logged in.')
print('Now, sit back and relax. Your internet connection is slow. I will send the mail')


receiver = email_id.rstrip()
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

msg['From'] = sender
msg['To'] = receiver
body = email.mime.Text.MIMEText(msg_text, 'html')
msg.attach(body)

try:
    mailer.sendmail(sender, receiver, msg.as_string())
    print('Mail successfully sent to ' + receiver)
    log_file.write('\n' + receiver + '      ' + format(datetime.datetime.now()) + '\n')
except Exception:
    print("Could not send mail to " + receiver)


mailer.quit()
log_file.close()
