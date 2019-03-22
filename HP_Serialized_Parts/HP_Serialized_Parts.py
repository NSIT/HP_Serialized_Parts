import sys
import smtplib
from email.message import EmailMessage
from email.headerregistry import Address
import logging

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
        subject=sys.argv[2] 
        files=sys.argv[4] 
        sender=sys.argv[6]

        content = ' \n The args past in are: \n {}'.format(the_args)
        subject = subject
        to = "yunior.rosellruiz@insight.com"
    except Exception as e:
        logging.basicConfig(filename='app.log', filemode='w', format='%(name)s - %(levelname)s - %(message)s')
        logging.warning(obj.params["reseller_name"]+" : "+ str(material) + " : "+  str(e))

