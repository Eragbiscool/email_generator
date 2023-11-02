#Run command: python3 script.py -csv_filepath "<Enter your csv filepath here>" -name_of_the_reportee "<name_of_the_reportee>" -agenda "<Agenda here>" -name_of_the_attendee "<Name of the attendee>" -name_of_the_host "<name of the host>" -location <meeting location> -sender_name "<Sender's name>" -organization "<Company name Here>" -to "<email of the receiver>" -from_email "<email of sender" -subject "<Subject here>" -custom_domain "<either "outlook" or "gmail" as per your email service provider>"



import csv
import smtplib, ssl 
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import re
from datetime import datetime
import argparse
from simple_html_table import Table
import numpy as np
import html
from getpass import getpass




argparser = argparse.ArgumentParser(description='Description:\n This script takes up the tos and froms for the email. Also the subject of the email and the name of the person its referred to. It will also ask for the project name and meeting location.')


result = ''''''
head = ''''''
tail = ''''''
data = []
col_number = []
row_number = []

current_date = datetime.now()
formatted_date = current_date.strftime("%B %d, %Y")


argparser.add_argument("-csv_filepath",type=str,help="Enter your csv filepath here!",required=True)


##email based args

argparser.add_argument("-custom_domain",type=str,help="Enter your email vendor. Either outlook or gmail for now",required=True)
argparser.add_argument("-from_email",type=str,help="Write the (From:) email id here")
argparser.add_argument("-to",type=str,help="Write the (To:) email id here")
argparser.add_argument("-cc",type=str,help="Write the (cc:) email id here")
argparser.add_argument("-bcc",type=str,help="Write the (bcc:) email id here")
argparser.add_argument("-subject",type=str,help="Write the subject of the email here")


#email body based args
argparser.add_argument("-meeting_date",type=str,help="Write it in Month Data,Year format",default=formatted_date)
argparser.add_argument("-name_of_the_reportee",type=str,help="Write the name of the reportee here.",required=True)
argparser.add_argument("-name_of_the_attendee",type=str,help="Write the name of the attendee here.",required=True,nargs="+")
argparser.add_argument("-name_of_the_host",type=str,help="Write the name of the Host here.",required=True)
argparser.add_argument("-location",type=str,help="Type the meeting location here",required=True)
argparser.add_argument("-agenda",type=str,help="Type your Agenda Here",required=True)
argparser.add_argument("-sender_name",type=str,help="Type sender's name here",required=True)
argparser.add_argument("-organization",type=str,help="Type sender's organization here",required=True)


args = argparser.parse_args()








def table_generation(input_csv):
    global table
    global data
    global col_number
    global row_number
    global result
 
    
    
    with open(input_csv, 'r') as csv_file:
        csv_reader = csv.reader(csv_file)
        for row in csv_reader:
            data.append(row)


    def merge_cells(data):
        row_count = 0
        
        for row in data:
            col_count = 0
            row_count = row_count + 1
            for field in row:
                col_count = col_count + 1
                if field == "":
                    col_number.append(col_count-1)
                    row_number.append(row_count-1)
                else:
                    pass

    
    merge_cells(data)

  
    new_data = []

    def clean(data):
        for row in data:
            field_list = []
            for field in row:
                alph_list= []
                for alph in field:
                    if(re.findall('[^\x00-\x7F]+', alph)):
                        alph_list.append(" ")
                    else:
                        alph_list.append(alph)
                field_list.append("".join(alph_list))  
            new_data.append(field_list)

    clean(data)
    dict = {}
    count = 0
    for i in col_number:
        if(dict.get(i) is not None):
            dict[i].append(row_number[count])
        else:    
            dict[i] = [row_number[count]]
        count = count +1
    
    rowspan_dict = {}

    for i in list(dict.keys()):
        if(len(dict[i])>1):
            row= min(dict[i])
            rowspan_dict[(row-1,i)] = len(dict[i])+1
        else:
            
            rowspan_dict[(dict[i][0]-1,i)] = 2




    table= Table((4,7),table_contents=np.array(new_data),attrs = {"border":"1"},cell_attrs={"align":"center","valign":"middle"})
    
    for i in list(rowspan_dict.keys()):
  
        j = list(i)
        table[j[0],j[1]].rowspan= rowspan_dict[i]



    table[0].header = True
    result = table.render()


    


   


def email_body(arg1,arg2,arg3,arg4,arg5,arg6,arg7,arg8,input_csv):
    global head
    global rowspan
    global tail
    


    head = '''
Dear {},
        
    I hope you are doing well. In this email, I added the meeting minutes of {}.
    
    Agenda: {}
    Host: {}
    Attendee: {}
    Time: Everyday
    Meeting Place: {}
    \n
'''.format(arg1,arg2,arg3,arg4,arg5,arg6)
    
    tail = '''
    Thanks,

    {}
    {}
'''.format(arg7,arg8)
    


    html_code_head = html.escape(head).replace('\n', '<br>').replace("Host",'<b><u>Host</b></u>').replace("Agenda",'<b><u>Agenda</b></u>').replace("Attendee",'<b><u>Attendee</b></u>').replace("Time",'<b><u>Time</b></u>').replace("Meeting Place",'<b><u>Meeting Place</b></u>')
    html_code_tail = html.escape(tail).replace('\n', '<br>')



    
    head = '''<div>{}</div>'''.format(html_code_head)
    tail = '''<div>{}</div>'''.format(html_code_tail)
  

    table_generation(input_csv)


email_body(args.name_of_the_reportee,args.meeting_date,args.agenda,args.name_of_the_host,args.name_of_the_attendee,args.location,args.sender_name,args.organization,args.csv_filepath)




main = head + result + tail



sender_email = args.from_email
receiver_email = args.to

message = MIMEMultipart("alternative")
message["Subject"] = args.subject
message["From"] = sender_email
message["To"] = receiver_email


html_code = result


part2 = MIMEText(main, "html")


message.attach(part2)


print("\n\n****Please enter the app password for GMAIL, NOT the original gmail password. If you dont have it, please follow line 237 to know how to generate app password****\n\n")
password = getpass('Type your password here: ')


###########################################################################################################
###################################### Enter pass below ###################################################
###################################### in your vendor's ###################################################
######################################  Section if you  ###################################################
###################################### want to hardcode ###################################################
######################################    your pass     ###################################################
###########################################################################################################

if(re.findall("^.*gmail.com.*",sender_email) or args.custom_domain == "gmail"):
    # password = "" #Turn your 2fa on in your gmail account. After that you will get "App passowrd" feature on. Create an app password and paste that password here. Need to do all this fuss for gmail only. If you want to hardcode your password, then comment the command in line 225 and uncomment this line and put your password in the quote


    context = ssl.create_default_context()
    with smtplib.SMTP_SSL("smtp.gmail.com", 465, context=context) as server:
        server.ehlo()
        try:
            server.login(sender_email, password)
        except Exception:
            print("\n\nYou might have inserted the \033[1m\033[31mWRONG\033[0m password!\n\n")
            raise Exception
        print("\n\n\033[1m\033[32mLogged in successfully!\033[0m Your email is being sent.... Please wait!!")
        server.sendmail(
            sender_email, receiver_email, message.as_string()
        )
        print("\n\n\033[1m\033[32mLogged out successfully!\033[0m Have a nice day !!!")

elif(re.findall("^.*outlook.com.*",sender_email) or args.custom_domain == "outlook"):
    
    # password = "" #Enter your password here. If you want to hardcode your password, then comment the command in line 225 and uncomment this line and put your password in the quote
    server = smtplib.SMTP('smtp.office365.com', 587)
    server.starttls()
    try:
        server.login(sender_email, password)
    except Exception:
        print("\n\nYou might have inserted the \033[1m\033[31mWRONG\033[0m password!\n\n")
        raise Exception
    print("\n\n\033[1m\033[32mLogged in successfully!\033[0m Your email is being sent.... Please wait!!")
    server.sendmail(sender_email, receiver_email, message.as_string())
    server.quit()
    print("\n\n\033[1m\033[32mLogged out successfully!\033[0m Have a nice day !!!")
else:
    raise Exception
