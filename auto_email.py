import pandas as pd
import smtplib

#sender details
your_name = input("Enter your name:")
your_email = input("Enter your Email:")
your_password = input("Enter password:")

#setting up the connection
server = smtplib.SMTP_SSL('smtp.gmail.com', 465)
server.ehlo()
server.login(your_email, your_password)

#file information
file = input("Enter filename")
email_list = pd.read_excel(file)

#getting all the email infos
all_names = email_list['Name']
all_emails = email_list['Email']
all_subjects = email_list['Subject']
all_messages = email_list['Message']

for idx in range(len(all_emails)):
    name = all_names[idx]
    email = all_emails[idx]
    subject = all_subjects[idx]
    message = all_messages[idx]+"\nThis is a system generated mail, Please "+name+" do not reply\n"
    full_email = ("From: {0} <{1}>\n"
                  "To: {2} <{3}>\n"
                  "Subject: {4}\n\n"
                  "{5}"
                  .format(your_name, your_email, name, email, subject, message))
    try:
        server.sendmail(your_email, [email], full_email)
        print('Email to {} successfully sent!\n\n'.format(email))
    except Exception as e:
        print('Email to {} could not be sent :( because {}\n\n'.format(email, str(e)))
server.close()
