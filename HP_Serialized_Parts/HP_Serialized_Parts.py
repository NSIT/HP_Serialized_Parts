import  sys
import  logging
import  time
import  re
import  pandas as pd
import  smtplib
from    email.message import EmailMessage
from    email.mime.multipart import MIMEMultipart
from    email.headerregistry import Address
import  mimetypes
from    email.mime.audio import MIMEAudio
from    email.mime.base import MIMEBase
from    email.mime.image import MIMEImage
from    email.mime.text import MIMEText
from    email import encoders



def send_email(subject,content,to,files):
    msg = MIMEMultipart()
    #msg.set_content(content)
    msg['Subject'] = subject
    msg['From'] = "do_not_reply@insight.com" #Address("Do Not Reply", "do_not_reply", "insight.com")
    msg['To'] =   to
    msg.attach(MIMEText(content))
    for f in files:
        a=attachment(f)
        msg.attach(a)
    s = smtplib.SMTP('mailna.insight.com')
    s.send_message(msg)
    s.quit()
    return 

def attachment(fileToSend):
    type = mimetypes.guess_type(fileToSend)
    ctype=type[0] if type[0] is not None else "application/octet-stream"
    encoding =type[1]

    maintype, subtype = ctype.split("/", 1)

    if maintype == "text":
        fp = open(fileToSend)
        # Note: we should handle calculating the charset
        attachment = MIMEText(fp.read(), _subtype=subtype)
        fp.close()
    elif maintype == "image":
        fp = open(fileToSend, "rb")
        attachment = MIMEImage(fp.read(), _subtype=subtype)
        fp.close()
    elif maintype == "audio":
        fp = open(fileToSend, "rb")
        attachment = MIMEAudio(fp.read(), _subtype=subtype)
        fp.close()
    else:
        fp = open(fileToSend, "rb")
        attachment = MIMEBase(maintype, subtype)
        attachment.set_payload(fp.read())
        fp.close()
        encoders.encode_base64(attachment)
    attachment.add_header("Content-Disposition", "attachment", filename=fileToSend)

    return attachment

if __name__ == '__main__':

    try:
        #subject=sys.argv[2] 
        #files=sys.argv[4] 
        #sender=sys.argv[6]

        #mfr= re.search("mfr='([a-z0-9]+)'",subject,re.RegexFlag.IGNORECASE).group(1)

        #Read file from email
        ##parts:pd.DataFrame=pd.read_excel("C:/Users/Public/ScheduledScripts/Imap_Monitor/email_downloads/038/20190326T150237/HP Serialized Parts List.xlsx",'parts',index_col=False)

        #Update parts table
        ##parts.to_sql(con=engine,schema="dbo",if_exists="replace",index=False,name=table_name)

        ##query = "SET NOCOUNT ON;EXEC [dbo].[sp_HP_Serialized_Parts] @mfr = '{0}'".format('0007081791')

        #pull data from stored procedure
        ##result= pd.read_sql_query(sql=query,con=engine)

        ##result.to_csv(output_file)
        
        content = ""
        subject = "None"
        to = "yunior.rosellruiz@insight.com"
        send_email(subject,content,to,(output_file,))

                   
    except Exception as e:
        logging.basicConfig(filename='app.log', filemode='w', format='%(name)s - %(levelname)s - %(message)s')
        logging.warning(str(e))
        

