import  sys
import  os
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
from    config import *
from    sqlalchemy import create_engine


def send_email(subject,content,to,files):
    msg = MIMEMultipart()
    #msg.set_content(content)
    msg['Subject'] = subject
    msg['From'] = "do_not_reply@insight.com" #Address("Do Not Reply", "do_not_reply", "insight.com")
    msg['To'] =   to
    msg.attach(MIMEText(content))
    if files is not None:
        for f in files:
            a=attachment(f)
            msg.attach(a)
    s = smtplib.SMTP('mailna.insight.com')
    s.send_message(msg)
    s.quit()
    return 

def attachment(fileToSend):
    fn=os.path.basename(fileToSend)
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
    attachment.add_header("Content-Disposition", "attachment", filename=fn)

    return attachment

if __name__ == '__main__':

    subject=sys.argv[2] 
    files=sys.argv[4] 
    sender=sys.argv[6]

    ##testing
    #subject='None'
    #files=r'\\insight.com\team\finance\Business Intelligence\Working Folders\Yunior\HP Serialized Parts List.xlsx'
    #sender='yunior.rosellruiz@insight.com'
    #mfr="0007081792"

    try:
        mfr= re.search("mfr=('[a-z0-9]+')",subject,re.RegexFlag.IGNORECASE).group(1)

        #Read file from email
        parts:pd.DataFrame=pd.read_excel(files,'parts',index_col=False)

        #Update parts table
        parts.to_sql(con=engine,schema="dbo",if_exists="replace",index=False,name=table_name)

        query = "SET NOCOUNT ON;EXEC [dbo].[sp_HP_Serialized_Parts] @mfr = {0}".format(mfr)

        #pull data from stored procedure
        result= pd.read_sql_query(sql=query,con=engine)
        result.to_csv(output_file)
        
        content = "Please see attached"
        subject = "HP_Serialized_Parts"
        to = sender
        send_email(subject,content,to,(output_file,))

                   
    except Exception as e:
        logging.basicConfig(filename='app.log', filemode='w', format='%(name)s - %(levelname)s - %(message)s')
        logging.warning(str(e))

    content = "Something went wrong, please reach out to the BI Team\n{}".format(str(e))
    subject = "Error"
    to = sender
    send_email(subject,content,to,None)



        

