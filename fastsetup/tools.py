#import all dependencies needed for the tools functions
try:
    import cx_Oracle
    import pandas
    import logging
    import datetime
    from email.mime.multipart import MIMEMultipart
    from email.mime.text import MIMEText
    from email.mime.base import MIMEBase
    from email.header import decode_header
    from email import encoders
    import smtplib
    import pysftp
    from sqlalchemy import create_engine
    import win32com.client
    import os

except Exception as error:
    print("import fastsetup - could not import library dependencies.")
    raise Exception(str(error))



def check_type(_argument, _type, _function_name):
    if not isinstance(_argument, _type):
        raise TypeError("fastsetup "+ _function_name +" - wrong data type passed in: " + str(_type))

#---------------------------------------------------------------------------------------------------------------------#
def main_log(_log_folder_path:str):
    check_type(_log_folder_path, str, 'main_log')

    try:
        #create a log file with today's date and define its format and set the level to DEBUG
        logging.basicConfig(filename=_log_folder_path + datetime.datetime.now().strftime( '%d-%m-%Y.log' ),
                            filemode='a',
                            format='Line: %(lineno)d - Time: %(asctime)s - Position: %(name)s - Status: %(levelname)s - Message: %(message)s',
                            datefmt='%H:%M:%S',
                            level=logging.DEBUG)
    except Exception as error:
        print("fastsetup main_log - could not create a log.")
        raise Exception("error: " + str(error))
   
    
    # create the main log names 
    loggerMain = logging.getLogger( '__main__' )
    loggerInit = logging.getLogger( '__init__' )
    loggerProcess = logging.getLogger( '__process__' )


    
    return loggerMain, loggerInit, loggerProcess
#---------------------------------------------------------------------------------------------------------------------#
def custom_log(_log_folder:str, _log_name:str, _log_format:str = None):
    check_type(_log_folder, str, 'custom_log')
    check_type(_log_name, str, 'custom_log')
    if not _log_format == None:
        check_type(_log_format, str, 'custom_log')

    # if the log_foramt string empty the format will be set to default 
    if _log_format == None:
       _log_format = 'Line: %(lineno)d - Time: %(asctime)s - Position: %(name)s - Status: %(levelname)s - Message: %(message)s' 
       print("no special format picked, "+str(_log_name)+" will have the main_log format")
    
    try:  
        # create a log file with today's date and name of your log and define its format as specified and set the level to DEBUG
        logging.basicConfig(filename=_log_folder + '__'+_log_name+'__' +datetime.datetime.now().strftime( '%d-%m-%Y.log' ),
                        filemode='a',
                        format=_log_format,
                        datefmt='%H:%M:%S',
                        level=logging.DEBUG)
    except Exception as error:
        print("fastsetup custom_log - could not create a log.")
        raise Exception("error: " + str(error))
    
    
    #create the main log name
    customLog = logging.getLogger('__'+_log_name+'__')
    
    return customLog
#---------------------------------------------------------------------------------------------------------------------#
def get_config(_config_path:str):
    check_type(_config_path, str, 'get_config')
    # gets data from config excel sheet file
    dfs = []
    try:
        xls = pandas.ExcelFile(_config_path)
    except Exception as error:
        print("fastsetup get_config - could not read excel file path")
        raise Exception("error: " + str(error))
    
    for sheet_name in xls.sheet_names:
        df = xls.parse(sheet_name)
        dfs.append(df)

    try:
        combined_df = pandas.concat(dfs, ignore_index=True)
        combined_df['key'] = combined_df['key'].astype(str)
        combined_df['value'] = combined_df['value'].astype(str)
    except Exception as error:
        print("fastsetup get_config - could not find key or value columns")
        raise Exception("error: " + str(error))
    
    for index,row in df.iterrows():
        config = combined_df.set_index('key')['value'].to_dict()

    return config
#---------------------------------------------------------------------------------------------------------------------#
def SMTP(_config:dict):
    check_type(_config, str, 'SMTP')

    # sents an email with or without an attachment
    # getting all the email components
    file_name = _config['file_name']
    sender_email = _config['sender_email']
    to_email = _config['to_email'].split(';')
    email_subject = _config['email_subject']
    body_type = _config['body_type']
    try: 
        body = _config['body']
    except Exception as error:
        print("fastsetup SMTP - ATTENTION: Email has no body or the implementation in the config is wrong.")
    try:
        cc = _config['cc'].split(';')
    except Exception as error:
        print("fastsetup SMTP - ATTENTION: Email has no CC recipient/s or the implementation in the config is wrong.")
    try:
        attachments = _config['attachments'].split(';')
    except Exception as error:
        print("fastsetup SMTP - ATTENTION: Email has no attachment/s or the implementation in the config is wrong.")
    
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

    print("fastsetup SMTP - email setup successfully")
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
        print("fastsetup SMTP_send - could not add attachments")
        raise Exception("error: " + str(error))

    print("fastsetup SMTP - added attachments successfully")
    ################################################################################################################### 
    # connect to the server and send the email
    try:
        with smtplib.SMTP(_config['server'], _config['port']) as server:
             server.sendmail(sender_email, recipients, text)
    except Exception as error:
        print("fastsetup SMTP - could not send the email")
        raise Exception("error: " + str(error))

    print("fastsetup SMTP - email Sent \n")
#---------------------------------------------------------------------------------------------------------------------#
def IMAP(_config:dict):
    check_type(_config, dict, "IMAP")

    email_found = False
    try:
        outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    except Exception as error:
        print("fastsetup IMAP - could not connect to the outlook")
        raise Exception("error: " + str(error))        
    inbox = outlook.GetDefaultFolder(6)
    messages = inbox.Items  
    for message in messages:
        if message.Subject == _config['IMAP_subject']:
            email_found = True
            for attachment in message.Attachments:
                attachment.SaveAsFile(os.path.join(_config['IMAP_path'], str(attachment)))
    if email_found == False:
        raise Exception("fastsetup IMAP - could not found the email download from the attachments needed")
#---------------------------------------------------------------------------------------------------------------------#
def SFTP_download(_config:dict):
    check_type(_config, dict, "SFTP_download")

    # to get the hostkeys for the sftp
    cnopts = pysftp.CnOpts()
    cnopts.hostkeys = None
    try:
        # SFTP connection
        with pysftp.Connection(host = _config['SFTP_download_host'], username = _config['SFTP_download_user_name'], password= _config['SFTP_download_password'] , port = int(_config['SFTP_download_port']) , cnopts = cnopts) as sftp:        
            try:
                # this to enter the file location before iteration  
                sftp.cwd(_config['server_folder_name'])
            except Exception as error:
                print("fastsetup SFTP_download - could not find server_folder_name")
                raise Exception("error: " + str(error))
            #######################################################
            # this for loop iterate through all ziped files until it fineds the required file
            directory_structure = sftp.listdir_attr()
            for attr in directory_structure:
                attr_file_name = str(attr.filename)
                if attr_file_name == _config['SFTP_download_server_file_name']:
                   remoteFilePath = attr_file_name
            #######################################################
            # gets the required ziped file
            try:        
                file =  sftp.get(remoteFilePath, _config['SFTP_download_local_file_path'])
                print("SFTP download - downloaded file successfully")
            except Exception as error:
                print("fastsetup SFTP_download - could not get file from the server")
                raise Exception("error: " + str(error))
            # close the connection
            sftp.close()
    except Exception as error:
        print("fastsetup SFTP_download - could not connect to the server")
        raise Exception("error: " + str(error))    
    return file
#---------------------------------------------------------------------------------------------------------------------#
def SFTP_upload(_config:dict):
    check_type(_config, dict, "SFTP_upload")

    # to get the hostkeys for the sftp
    cnopts = pysftp.CnOpts()
    cnopts.hostkeys = None
    try:
        # SFTP connection
        with pysftp.Connection(host = _config['SFTP_upload_host'], username = _config['SFTP_upload_user_name'], password= _config['SFTP_upload_password'] , port = int(_config['SFTP_upload_port']) , cnopts = cnopts) as sftp:
            print("SFTP upload - connection successful")

            try:
                #call sftp.put() method to upload file to server
                sftp.put(_config['SFTP_upload_local_file'], _config['SFTP_upload_target_location'])
                print("SFTP upload - upload successful")
            except Exception as error:
                print("fastsetup SFTP_upload - could not upload the file")
                raise Exception("error: " + str(error))

            # close the connection
            sftp.close()      
            print("SFTP upload - connection closed successful \n")
    except Exception as error:
        print('fastsetup SFTP_upload - could not connect to the server')
        raise Exception("error: " + str(error))
#---------------------------------------------------------------------------------------------------------------------#
def oracle_download(_config:dict):
    check_type(_config, dict, "oracle_download")

    try:
        dsn = cx_Oracle.makedsn(_config['oracle_download_host'], _config['oracle_download_port'], service_name = _config['oracle_download_data_source'])
    except Exception as error:
        print("fastsetup oracle_download - could not make dsn")
        raise Exception("error: " + str(error))
    
    try:
        connection = cx_Oracle.connect(user=_config['oracle_download_user_ID'], password=_config['oracle_download_password'], dsn=dsn)
        print("fastsetup oracle_download - connected to oracle database")
    except Exception as error:
        print("fastsetup oracle_download - could not connect to oracle")
        raise Exception("error: " + str(error))    
    
    try:
        df = pandas.read_sql(_config['oracle_download_query'], con=connection)
        print("fastsetup oracle_download - query found")
    except Exception as error:
        print("fastsetup oracle_download - could not find the query")
        raise Exception("error: " + str(error))
              
    connection.close()
    print("fastsetup oracle_download - connection closed. \n")
    
    return df
#---------------------------------------------------------------------------------------------------------------------#
def oracle_upload(_config:dict, dataframe:pandas.DataFrame = None):
    check_type(_config, dict, "oracle_upload")

    if not dataframe == None:
        check_type(dataframe, pandas.DataFrame, "oracle_upload")
    else:    
        dataframe = _config["oracle_upload_file_path"]
    try:
        dsn = cx_Oracle.makedsn(_config['oracle_upload_host'], _config['oracle_upload_port'], service_name = _config['oracle_upload_data_source'])
    except Exception as error:
        print("fastsetup oracle_download - could not make dsn")
        raise Exception("error: " + str(error))
    
    try:
        connection = cx_Oracle.connect(user=_config['oracle_upload_user_ID'], password=_config['oracle_upload_password'], dsn=dsn)
        print("fastsetup oracle_download - connected to oracle database")
    except Exception as error:
        print("fastsetup oracle_download - could not connect to oracle")
        raise Exception("error: " + str(error))    
    
    try:
        dataframe.to_sql(_config['oracle_upload_query'], con=connection)
        print("fastsetup oracle_download - query found")
    except Exception as error:
        print("fastsetup oracle_download - could not find the query")
        raise Exception("error: " + str(error))
              
    connection.close()
    print("fastsetup oracle_download - connection closed. \n")
#---------------------------------------------------------------------------------------------------------------------#
def sql_upload(_config:dict, df_to_upload:pandas.DataFrame = None):
    check_type(_config, dict, "sql_upload")
    if not df_to_upload == None:
        check_type(df_to_upload, pandas.DataFrame, "sql_upload")
        file = df_to_upload
    else:    
        file = _config["sql_upload_file_path"]

    try:
        engine = create_engine("mssql+pyodbc://"+_config['sql_upload_username']+":"+_config['sql_upload_password']+"@"+_config['sql_upload_servername']+"/"+_config['sql_upload_database']+ "?driver=SQL+Server") 
    except Exception as error:
        print("fastsetup SQL_upload - could not create engine")
        raise Exception("error: " + str(error))
    try:
        file.to_sql(_config['sql_upload_table_name'], con=engine, if_exists='append')
        print("sql upload successfully")
    except Exception as error:
        print("fastsetup sql_upload - could not upload file")
        raise Exception("error: " + str(error))
#---------------------------------------------------------------------------------------------------------------------#   
def sql_download(_config:dict):
    check_type(_config, dict, "sql_download_")
    dataframe = None
    try:
        engine = create_engine("mssql+pyodbc://"+_config['sql_download_username']+":"+_config['sql_download_password']+"@"+_config['sql_download_servername']+"/"+_config['sql_download_database']+ "?driver=SQL+Server") 
    except Exception as error:
        print("fastsetup SQL_upload - could not create engine")
        raise Exception("error: " + str(error))
    try:
        dataframe = pandas.to_sql(_config['sql_download_table_name'], con=engine, if_exists='append')
        print("sql upload successfully")
    except Exception as error:
        print("fastsetup sql_upload - could not upload file")
        raise Exception("error: " + str(error))
    return dataframe