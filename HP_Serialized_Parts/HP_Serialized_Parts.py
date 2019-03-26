import  sys
import  smtplib
from    email.message import EmailMessage
from    email.headerregistry import Address
import  logging
import  time
import re
import panda as pd

def send_email(subject,content,to):
    msg = EmailMessage()
    msg.set_content(content)
    msg['Subject'] = subject
    msg['From'] = Address("Do Not Reply", "do_not_reply", "insight.com")
    msg['To'] =   to
    s = smtplib.SMTP('mailna.insight.com')
    s.send_message(msg)
    s.quit()
    return 

if __name__ == '__main__':

    try:
        #subject=sys.argv[2] 
        #files=sys.argv[4] 
        #sender=sys.argv[6]

        #mfr= re.search("mfr='([a-z0-9]+)'",subject,re.RegexFlag.IGNORECASE).group(1)

        parts=pd.read_excel(" C:/Users/Public/ScheduledScripts/Imap_Monitor/email_downloads/038/20190326T150237/HP Serialized Parts List.xlsx",'parts',header=False)
        

        content = ' \n The args past in are: \n {}'.format(parts.head())
        subject = "None"
        to = "yunior.rosellruiz@insight.com"
        send_email(subject,content,to)
                   
    except Exception as e:
        logging.basicConfig(filename='app.log', filemode='w', format='%(name)s - %(levelname)s - %(message)s')
        logging.warning(str(e))
        

