#---------------------------------------------------------------------------------------------------------------------#
def main_log(log_folder):
    import logging
    import datetime
    
    
    try:
        #create a log file with today's date and define its format and set the level to DEBUG
        logging.basicConfig(filename=log_folder + datetime.datetime.now().strftime( '%d-%m-%Y.log' ),
                            filemode='a',
                            format='Line: %(lineno)d - Time: %(asctime)s - Position: %(name)s - Status: %(levelname)s - Message: %(message)s',
                            datefmt='%H:%M:%S',
                            level=logging.DEBUG)
    except Exception as error:
        return error
   
    
    # create the main log names 
    loggerMain = logging.getLogger( '__main__' )
    loggerInit = logging.getLogger( '__Init__' )
    loggerProcess = logging.getLogger( '__Process__' )


    
    return loggerMain, loggerInit, loggerProcess
#---------------------------------------------------------------------------------------------------------------------#
def custom_log(log_folder, log_format, log_name):
    import logging
    import datetime
    
    # if the log_foramt string empty the format will be set to default 
    if log_format == None:
       log_format = 'Line: %(lineno)d - Time: %(asctime)s - Position: %(name)s - Status: %(levelname)s - Message: %(message)s'
        
       
    # create a log file with today's date and name of your log and define its format as specified and set the level to DEBUG
    logging.basicConfig(filename=log_folder + '__'+log_name+'__' +datetime.datetime.now().strftime( '%d-%m-%Y.log' ),
                        filemode='a',
                        format=log_format,
                        datefmt='%H:%M:%S',
                        level=logging.DEBUG)

    
    
    #create the main log name
    customLog = logging.getLogger('__'+log_name+'__')
    
    return customLog
#---------------------------------------------------------------------------------------------------------------------#
def get_config(Config_Path, SheetName, Key, Value):
    import pandas
    
        
    # Read the config file with sheet name and making sure all keys and values are string type
    df = pandas.read_excel(Config_Path,sheet_name = SheetName) 
    df[Key] = df[Key].astype(str)
    df[Value] = df[Value].astype(str)
    
    #iterate through the config dataframe and initialize every Key with its value  
    for index,row in df.iterrows():
        Config = df.set_index(Key)[Value].to_dict()
   
    
    return Config
#---------------------------------------------------------------------------------------------------------------------#
def SMTP(config):
    from email.mime.multipart import MIMEMultipart
    from email.mime.text import MIMEText
    from email.mime.base import MIMEBase
    from email import encoders
    import smtplib
    
    # getting all the email components
    file_name = config['file_name']
    sender_email = config['sender_email']
    to_email = config['to_email'].split(';')
    email_subject = config['email_subject']
    sender_email = config['sender_email']
    body_type = config['body_type']
    try: 
        body = config['body']
    except Exception as error:
        print("ATTENTION: Email has no body or the implementation in the config is wrong.")
        print("The catched error is: " + str(error))      
    try:
        cc = config['cc'].split(';')
    except Exception as error:
        print("ATTENTION: Email has no CC recipient/s or the implementation in the config is wrong.")
        print("The catched error is: " + str(error))
    
    try:
        attachments = config['attachments'].split(';')
    except Exception as error:
        print("ATTENTION: Email has no attachment/s or the implementation in the config is wrong.")
        print("The catched error is: " + str(error))

    print("Config read successfully \n")
    ###################################################################################################################
    # Build the body of the Email
    message = MIMEMultipart('alternative')
    message["From"] = sender_email
    message['To'] = ', '.join(to_email)
    message["Subject"] = email_subject
    try:    
        message["CC"] = ', '.join(cc)
        recipients = to_email + cc
    except:
        recipients = to_email
    message.attach(MIMEText(body, body_type))

    print("Email setup successfully")
    ###################################################################################################################
    # iterate through all the files and attache them with the email
    try:
        for attachment_path in attachments:
            # attach the files
            with open(attachment_path, "rb") as attachment:
                file = MIMEBase("application", "octet-stream")
                file.set_payload(attachment.read())
        
            # Encode file in ASCII characters to send by email
            encoders.encode_base64(file)
            # add name for the file
            file.add_header("Content-Disposition",f"attachment; filename= {file_name}",)
            # Add attachment to message and convert message to string
            message.attach(file)
            text = message.as_string()
    except Exception as error:
        print(str(error))

    print("Added attachments successfully")
    ################################################################################################################### 
    # connect to the server and send the email
    with smtplib.SMTP(config['server'], config['port']) as server:
         server.sendmail(sender_email, recipients, text)
    
    print("Email Sent \n")
#---------------------------------------------------------------------------------------------------------------------#
def SFTP_download(config):
    import pysftp

    # to get the hostkeys for the sftp
    cnopts = pysftp.CnOpts()
    cnopts.hostkeys = None
    
    
    # SFTP connection
    with pysftp.Connection(host = config['host'], username = config['user_name'], password= config['password'] , port = int(config['SFTP_port']) , cnopts = cnopts) as sftp:
        print("SFTP download - connection successful")
        
        # this to enter the file location before iteration  
        sftp.cwd(config['server_folder_name'])
        ##############################
        # this for loop iterate through all ziped files until it fineds the required file
        directory_structure = sftp.listdir_attr()
        for attr in directory_structure:
            attr_file_name = str(attr.filename)
            if attr_file_name == config['server_file_name']:
               remoteFilePath = attr_file_name
        ###############################
        # gets the required ziped file        
        result =  sftp.get(remoteFilePath, config['local_file_path'])
        print("SFTP download - downloaded file successfully")

        # close the connection
        sftp.close()

        print("SFTP download - connection closed successful \n")
    
    return result
#---------------------------------------------------------------------------------------------------------------------#
def SFTP_upload(config):
    import pysftp

    # to get the hostkeys for the sftp
    cnopts = pysftp.CnOpts()
    cnopts.hostkeys = None

    # SFTP connection
    with pysftp.Connection(host = config['host'], username = config['user_name'], password= config['password'] , port = int(config['SFTP_port']) , cnopts = cnopts) as sftp:
        print("SFTP upload - connection successful")
        
        #call sftp.put() method to upload file to server
        sftp.put(config['local_file'], config['target_location'])
        print("SFTP upload - upload successful")

        # close the connection
        sftp.close()      
        print("SFTP upload - connection closed successful \n")
#---------------------------------------------------------------------------------------------------------------------#
def oracle_download(Config):
    import cx_Oracle
    import pandas
    
    dsn = cx_Oracle.makedsn(Config['host'], Config['port'], service_name = Config['data_source'])
    connection = cx_Oracle.connect(user=Config['user_ID'], password=Config['password'], dsn=dsn)
    print("oracle_connection - connected to oracle database")
    
    df = pandas.read_sql(Config['query'], con=connection)
    print("oracle_connection - query found")
              
    connection.close()
    print("oracle_connection - connection closed. \n")
    
    return df
    