import pandas as pd
import smtplib
import datetime as dt

data = pd.read_excel('/Users/alexrasla/Documents/jobs.xlsx') 

# companies = pd.DataFrame(data, columns= ['Company Name'])
# emails = pd.DataFrame(data, columns= ['Email'])
last_comm = pd.DataFrame(data, columns=['Last Communication'])

# print(companies)

today = dt.datetime.now()

server = smtplib.SMTP("smtp.gmail.com", 587)
server.starttls()
if(last_comm.size > 0):
    server.login("EMAIL", "PASSWORD")

for index in range(last_comm.size):
    
    curr_last_comm = data.at[index, 'Last Communication']
    company_name = data.at[index, 'Company Name']
    difference = today - curr_last_comm
    
    body = ["Dear {},\n\nI'm wondering if you've recieved my application that I submitted 3 days ago. I'm very interested in working for your company and hope to be working together in the future. If you would like any additional information about my experience or have any question, please don't hestitate to reach out. I look forward to hearing back from you.".format(company_name, company_name),
    "Dear {},\n\nI've sent you a few emails regarding a job application I submitted 5 days ago. Please let me know as soon as you make a decision on my application.".format(company_name, company_name),
    "Dear {},\n\nIf you've taken a look at my application, would you please be able to give me an update on your decision? Thank you.".format(company_name, company_name)]

    signature = "Sincerely,\nAlex Rasla\nUCSB Undergraduate\nalexrasla@ucsb.edu\n949-748-9505"
    
    if(data.at[index, 'Response'] == 'F'):
        if(difference.days == 5):
            message = 'Subject: Follow Up\n\n{}\n\n{}'.format(body[0], signature)
            server.sendmail("alexrasla@ucsb.edu", data.at[index, 'Email'], message)
        elif(difference.days == 9):
            message = 'Subject: Status Regarding Application\n\n{}\n\n{}'.format(body[1], signature)
            server.sendmail("alexrasla@ucsb.edu", data.at[index, 'Email'], message)
        elif(difference.days == 14):
            message = 'Subject: Application Update\n\n{}\n\n{}'.format(body[2], signature)
            server.sendmail("alexrasla@ucsb.edu", data.at[index, 'Email'], message)

server.quit()