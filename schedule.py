''' 
scheduler + sending mail  
trial1 
writer: Hyein Cho  
'''

import smtplib, imaplib, email, time, re, glob, os, sys 
import xlrd, xlwt 
from datetime import date, datetime, timedelta 
from email.mime.text import MIMEText 
from email.mime.multipart import MIMEMultipart 
from email.header import Header 
from apscheduler.schedulers.background import BackgroundScheduler 
from apscheduler.jobstores.base import JobLookupError 

# global variable 
dept_dict = dict(dept_1=[0,'dept_1@naver.com'],dept_2=[0,'dept_2@naver.com'], 
               dept_3=[0,'dept_3@naver.com']) 

n = datetime.now()

# counter, stop when it hit 0. 
num_dept=len(dept_dict)
html_list=list()

# In case program shut down for some reason. 
# check the DB xls files and check what department send email or not. 
def counter_check():

    global num_dept
    weeks=(str(datetime.now().year)+str(datetime.now().isocalendar()[1]))

    for key, val in dept_dict.items():
        file=glob.glob('./DB_'+key+'.xls')
        if os.path.exists(file[0]):
            wb=xlrd.open_workbook(file[0])
            sheets=wb.sheets()
            nrows=sheets[0].nrows 
            ncols=sheets[0].ncols
            if str(sheets[0].cell_value(nrows-1,0)) == weeks:
                val[0]=1
                num_dept-=1
                print(key+' did send email') 
            else:
                print('testing pass block') 
                print(key+' did not send email') 
                pass 
        else:
            print('The file for '+key+' does not exist. Call!') 

    if num_dept == 0:  
        print('Every department sent email for next work') 
        sys.exit() 

# From Wedseday to Friday, set the time interval of working period
def date_interval():
    date=str()
    # Wednseday
    if n.weekday() == 2:
        date=(str(n+timedelta(days=-2))[:10] + ' ~ ' + str(n+timedelta(days=2))[:10])
    # Thursday
    elif n.weekday() == 3:
        date=(str(n+timedelta(days=-3))[:10] + ' ~ ' + str(n+timedelta(days=1))[:10])
    # Friday
    elif n.weekday() == 4:
        date=(str(n+timedelta(days=-4))[:10] + ' ~ ' + str(n+timedelta(days=0))[:10]) 
    return date

# It actually send emails to employees 
def email_content(receiver): 
    print('sending email to: ', receiver) 

    # email title and content 
    title=(str(n.isocalendar()[1]) + ' th work should be added. (' + date_interval() + ')')
    content=('Please fill out the form. The important notes should be in (* *)\n\n' + 'Department: \n' + 'Next week\'s work: \n Important note:') 

    # message setting 
    msg=MIMEText(content) 
    msg['Subject']=Header(title, 'utf-8') 
    msg['From']='email_bot@naver.com'
    msg['To']=receiver 

    with smtplib.SMTP_SSL('smtp.naver.com') as smtp: 
        smtp.login('email_bot', 'password') 
        smtp.send_message(msg) 

# send emails to departments 
def send_email():
    
    # first look the counter
    if num_dept>0: 
        # also check the second counter, dictionary
        for val in dept_dict.values():
            if val[0] == 0:
                email_content(val[1])
                print("email has been sent to ", val[1], "| [time] "
                    , str(time.localtime().tm_hour) + ":"
                    + str(time.localtime().tm_min) + ":"
                    + str(time.localtime().tm_sec))
    else:
        print('Good! All departments are done')
        print('Saveing files. Time: '+str(datetime.now()))

        # If all departments send next week's work, save files.   
        # counter 에서 dept_num=0 가 되면 실행시킬 수도 있지만 고민해야 한다.  
        save_files()

# Check the inbox and change the dictionary counter from 0 to 1 if department sent email.
# Change Unseen email to be Seen. 
# Store all department next week's work individually in html.
def check_mail_imap():
    print('Checking email_bot\'s. Time: ' + str(datetime.now()))
    #imap server  
    user='email_bot@naver.com'
    password='password'
    
    imapsrv="imap.naver.com"
    imapserver=imaplib.IMAP4_SSL(imapsrv, "993")
    imapserver.login(user, password)
    imapserver.select('INBOX')

    res, unseen_data = imapserver.search(None, '(UNSEEN)')
    ids = unseen_data[0]
    id_list = ids.split()
    latest_email_id = id_list[-10:]

    for each_mail in latest_email_id:  
        # fetch the email body (RFC822) for the given ID 
        result, data = imapserver.fetch(each_mail, "(RFC822)") 
        dept=str()
        msg = email.message_from_bytes(data[0][1])
        sender=re.search("<.*>",msg['From']).group()[1:-1]

        # make counter to be replied after checking the department  
        for key, val in dept_dict.items():
            if val[1] == sender:
                val[0]=1
                dept=key

        global num_dept
        num_dept-=1
        imapserver.store(each_mail, '+FLAGS', '\Seen')

        while msg.is_multipart():
            msg = msg.get_payload(0)

        content = msg.get_payload(decode=True)
        file_path=("./email_" + date_interval() + '_' + dept + '.html')

        f = open(file_path, "wb")
        f.write(content)
        print(file_path + " has been saved. Time: "+ str(datetime.now())) 

        print("remained department is "+str(num_dept))
        f.close()

    imapserver.close()
    imapserver.logout()

# HTML Template that showing this week's work report. 
def html_template():
    html_str=(
        "<!DOCTYPE html>\n"+
        "<html>\n"+
        "<head>\n"+
        "    <meta charset='utf-8'>\n"+
        "    <meta http-equiv='Content-Type' content='text/html; charset=utf-8'>\n"+
        "<style>\n"+
        "table, th, td {\n"+
        "    border: 1px solid black;\n"+
        "    border-collapse: collapse; \n"+
        "}  \n"+
        "}\n"+
        "</style>\n"+
        "</head>\n"+
        "<body>\n"+
        "<h2>This weeks work! </h2>\n"+
        "<table style='width:80%''>\n"+ 
        "<tr>\n"+
        "<th>Division</th>\n"+
        "<th>Content</th> \n"+
        "</tr>\n"+
        "<tr>\n"+
        "<td align='center'>Next week's work</br>"+date_interval()+" \n"+
        "<td>")

    # for loop, only important note should be added in the final report
    global html_list
    for item in html_list:
        html_str+=(item[1].replace('\n', '')+'('+item[0][4:].replace(' ', '')+')</br></br>')

    html_str+=("</td>\n"+
        "</tr>\n"+
        "<tr>\n"+
        "<td align='center'>Important note</td>\n"+
        "<td>")

    # for loop 
    for item in html_list:
        html_str+=(item[2][5:].replace('\n', '')+'('+item[0][4:-1].replace(' ', '')+')</br>')

    html_str+=("</td>\n"+
        "</tr>\n"+
        "</table>\n"+
        "</body>\n"+
        "</html>\n"
    )

    return html_str

    # with open('./'+date_interval()+'_report.html', 'w') as html_write: 
    #     html_write.write(html_str) 
    #     print('Report has been made. Time: '+str(datetime.now())) 


# Send the final report to the all employee.
def send_report():

    for val in dept_dict.values():
        receiver=val[1]
        print('sending report email to: ', receiver)

        title=(str(n.isocalendar()[1]) + ' th week\'s (' + date_interval() + ') report!')

        msg=MIMEMultipart()
        msg['Subject']=Header(title, 'utf-8')
        msg['From']='email_bot@naver.com'
        msg['To']=receiver

        content=MIMEText('Regards', 'plain')
        html_str=MIMEText(html_template(), 'html')
        msg.attach(content)
        msg.attach(html_str)

        with smtplib.SMTP_SSL('smtp.naver.com') as smtp:
            smtp.login('email_bot', 'password')
            smtp.sendmail('email_bot@naver.com', receiver, msg.as_string())

        print("Report email has been sent to ", val[1], "| [time] "
            , str(time.localtime().tm_hour) + ":"
            + str(time.localtime().tm_min) + ":"
            + str(time.localtime().tm_sec))

    # Everything is done. Alarm.
    print('Everything is done. Time : '+ str(datetime.now())+ 'Shutting down...')
    sys.exit()


# individual file of department saved in email_[date_interval]_[department].html
def save_files(): 

    # File path 
    path = './email_*.html'
    files = glob.glob(path)
    save_interval=date_interval()

    # Open and read. Parsing.
    for file in files:
        with open(file, mode='r', encoding='utf-8') as f:
            lines=f.read().replace('&nbsp;', '')

            # department
            dept=re.compile(r"부서:[\s\w\S]*Next week\'s work", re.MULTILINE)
            dept_str=dept.findall(lines)[0][:-4]
            
            dept_file_path=str("./DB_"+dept_str.replace(' ', '')[3:])

            # content of next week's work
            nextwk=re.compile(r"Next week\'s work:[\s\w\S]*Important note", re.MULTILINE) 
            nextwk_str=nextwk.findall(lines)[0][:-4]

            impt=re.compile(r"['(*'] [\s\w\S]* ['*)']", re.MULTILINE)
            impt_str=impt.findall(nextwk_str)[0][1:-1]

            special=re.compile(r"Important note:[\s\w\S]*", re.MULTILINE)
            special_str=special.findall(lines)[0]

            # save files by department
            db=(dept_file_path+".txt")
            with open(db, mode='a', encoding='utf-8') as db_file:
                write_str=save_interval+'\t'+dept_str + '\t' + nextwk_str + '\t' + impt_str + '\t'    + special_str + '\n'
                db_file.write(write_str)
                print(dept_file_path+'.txt has been saved. Time: '+str(datetime.now()))

            # final report saving...
            global html_list
            html_list.insert(len(html_list), [dept_str, impt_str, special_str])

            # writing xlsx if file exist, jusg add. if not create new xlsx file
            if os.path.exists(dept_file_path+".xls"):
                excel_wb=xlrd.open_workbook(dept_file_path+".xls")
                sheets=excel_wb.sheets()

                nrows=sheets[0].nrows
                ncols=0

                # dictionary for adding next week's information to xlsx file.
                datadict={}
                for row_num in range(nrows):
                    datadict[row_num]={}
                    for col in range(5):
                        datadict[row_num][col]=sheets[0].cell_value(row_num,col)
                excel_write=xlwt.Workbook(encoding='utf-8')
                excel_ws=excel_write.add_sheet('Sheet1', cell_overwrite_ok=True)

                for row_num in range(nrows):
                    for col in range(5):
                        excel_ws.write(row_num,col,datadict[row_num][col])

                excel_ws.write(nrows, ncols, str(n.year)+str(n.isocalendar()[1]))
                excel_ws.write(nrows, ncols+1, save_interval)
                excel_ws.write(nrows, ncols+2, dept_str)
                excel_ws.write(nrows, ncols+3, nextwk_str)
                excel_ws.write(nrows, ncols+4, impt_str)
                excel_ws.write(nrows, ncols+5, special_str)
                excel_write.save(dept_file_path+".xls")

                print(dept_file_path + '.xls has been saved. Time: '+str(datetime.now()))

            else:
                excel_write=xlwt.Workbook(encoding='utf-8')
                nrows=0
                ncols=0

                excel_ws=excel_write.add_sheet('Sheet1', cell_overwrite_ok=True)
                excel_ws.write(nrows, ncols, str(n.year)+str(n.isocalendar()[1]))
                excel_ws.write(nrows, ncols+1, save_interval)
                excel_ws.write(nrows, ncols+2, dept_str)
                excel_ws.write(nrows, ncols+3, nextwk_str)
                excel_ws.write(nrows, ncols+4, impt_str)
                excel_ws.write(nrows, ncols+5, special_str)

                excel_write.save(dept_file_path+".xls") 

                print(dept_file_path + '.xls has been saved. Time: '+str(datetime.now()))

            # delete temporary saved file ... email_[date_interval]_[department].html
            os.remove(file)

    # if all departments sent email. create the final report.
    # after creation, sent the final report to employees. 
    with open('./report_'+date_interval()+'.html', 'w') as html_write:
        html_write.write(html_template())
        print('The final report has been created. Time: '+str(datetime.now()))

    send_report()


# main method
if __name__=="__main__":
    sched = BackgroundScheduler()
    sched.start()

    # wednseday
    if date.today().weekday() == 2:
        print('Today is Wednseday. You should send next week\'s work by day after tomorrow.')
        # counter in case of shutting down.
        try:
            counter_check()
        except:
            print('No file exist')
        
        # today is wedseday, everyone should send email.
        # notification to let employee send emails.
        send_email()
        sched.add_job(check_mail_imap, 'cron', second=30, id='check')

        # send notification email if any department did not send email. From 2~6 pm, every hour.
        sched.add_job(send_email, 'cron', day_of_week='wed', hour='14-18/1', id="wed_job") 

    # thursday
    elif date.today().weekday() == 3: 
        # check email when 9, and send notifications if any department did not send email. From 9 a.m. - 6 p.m. every hour.
        print('Today is Thursday. You should send next week\'s work by tomorrow.')
        try:
            counter_check()
        except:
            print('No file exist')
        sched.add_job(send_email, 'cron', day_of_week='thu', hour='9-18/1', id="thu_job")
        sched.add_job(check_mail_imap, 'cron', second=30, id='check')

    # friday
    elif date.today().weekday() == 4: 
        # check email when 9, and send notifications if any department did not send email. From 9 a.m. - 6 p.m. every hour.
        print('Today is Friday. You should send next week\'s work by today')
        try:
            counter_check()
        except:
            print('No file exist')
        sched.add_job(send_email, 'cron', day_of_week='fri', hour='9-18/1', id="fri_job")
        sched.add_job(check_mail_imap, 'cron', second=30, id='check')
    else:
        print('The process is scheduled between Wednseday to Friday')
        sys.exit()

    # this program should work on until all departments send email and create the final report. 
    while True:
        print('----- Running Main Process. Time: '+str(datetime.now())+'------')
        time.sleep(1)

