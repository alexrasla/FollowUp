from email.MIMEMultipart import MIMEMultipart
from email.MIMEText import MIMEText
from email.MIMEImage import MIMEImage
from email.mime.application import MIMEApplication

import pandas as pd
import smtplib
import datetime as dt

data = pd.read_excel('/Users/alexrasla/Documents/jobs.xlsx') 

# companies = pd.DataFrame(data, columns= ['Company Name'])
# emails = pd.DataFrame(data, columns= ['Email'])
last_comm = pd.DataFrame(data, columns=['Last Communication'])

today = dt.datetime.now()

server = smtplib.SMTP("smtp.gmail.com", 587)
server.starttls()
if(last_comm.size > 0):
    server.login("alexrasla@ucsb.edu", "474047VmS")

att = MIMEApplication(open("/Users/alexrasla/Documents/Rasla, Alex, Resume.pdf", 'rb').read())
att.add_header('Content-Disposition', 'attachment', filename='Rasla, Alex, Resume.pdf')

for index in range(last_comm.size):
    
    curr_last_comm = data.at[index, 'Last Communication']
    company_name = data.at[index, 'Company Name']
    difference = today - curr_last_comm

    dear = company_name + " Recruiter"
    
    body = [
    "Dear {},\n\nMy name is Alex Rasla, and I am a computer science student at UCSB in the process of obtaining a bachelors and masters degree in 5 years. I was wondering if you had any software engineering part-time or internship opportunities for this coming fall, or anything full-time for next summer. I love what {} does in the community and would love to work for the company. Please let me know if there is anything you know about, or anyone you can connect me with that might know. Thank you, I look forward to hearing back from you.".format(dear, company_name),
    "Dear {},\n\nI'm wondering if you've had a chance to read my email that I sent a couple days ago. I'm very interested in working for {} and hope to be working together in the future. If you know of any potential opportunities for which I am a competitive candidate, please let me know. I look forward to hearing back from you.".format(dear, company_name),
    "Dear {},\n\nI've sent you a few emails regarding employment opporunities last week. Please let me know about any opportunties where I could be helpful to the company.".format(dear, company_name),
    "Dear {},\n\nI would love if you could share any relevant employment opportunities with me. Thank you.".format(dear, company_name)
    ]

    signature = "Sincerely,\nAlex Rasla\nUCSB Computer Science B.S./M.S. Student\nalexrasla@ucsb.edu\n949-748-9505"
    
    if(data.at[index, 'Response'] == 'F'):
        message = MIMEMultipart()
        
        if(difference.days == 0):
            message['Subject'] = 'Employment Opportunities'
            curr_body = MIMEText('{}\n\n{}'.format(body[0], signature))
            message.attach(curr_body)
            message.attach(att)
            server.sendmail("alexrasla@ucsb.edu", data.at[index, 'Email'], message.as_string())
            print('Sent mail to {} at {} after {} days'.format(company_name, data.at[index, 'Email'], difference.days))
        elif(difference.days == 3):
            message['Subject'] = 'Follow Up'
            curr_body = MIMEText('{}\n\n{}'.format(body[1], signature))
            message.attach(curr_body)
            message.attach(att)
            server.sendmail("alexrasla@ucsb.edu", data.at[index, 'Email'], message.as_string())
            print('Sent mail to {} at {} after {} days'.format(company_name, data.at[index, 'Email'], difference.days))
        elif(difference.days == 9):
            message['Subject'] = 'Status Regarding Opportunities'
            curr_body = MIMEText('{}\n\n{}'.format(body[2], signature))
            message.attach(curr_body)
            message.attach(att)
            server.sendmail("alexrasla@ucsb.edu", data.at[index, 'Email'], message.as_string())
            print('Sent mail to {} at {} after {} days'.format(company_name, data.at[index, 'Email'], difference.days))
        elif(difference.days == 14):
            message['Subject'] = 'Empoloyment Opportunities Update'
            curr_body = MIMEText('{}\n\n{}'.format(body[3], signature))
            message.attach(curr_body)
            message.attach(att)
            server.sendmail("alexrasla@ucsb.edu", data.at[index, 'Email'], message.as_string())
            print('Sent mail to {} at {} after {} days'.format(company_name, data.at[index, 'Email'], difference.days))

server.quit()