def _check_dependencies():
    _hard_dependencies = ("cx_Oracle", "pandas", "logging", "datetime", "email", "smtplib", "pysftp", "sqlalchemy", "requests","os","openpyxl")
    _missing_dependencies = []

    for _dependency in _hard_dependencies:
        try:
            __import__(_dependency)
        except ImportError as _e:  
            _missing_dependencies.append(f"{_dependency}: {_e}")

    if _missing_dependencies:  
        raise ImportError("Unable to import required dependencies:\n" + "\n".join(_missing_dependencies))
    
    del _hard_dependencies, _dependency, _missing_dependencies

def _create_config(path):
    import openpyxl
    import os
    if not os.path.exists(path):

        workbook = openpyxl.Workbook()

        sheet_names = ["settings","SMTP", "IMAP", "SFTP_download", "SFTP_upload", "oracle_download", "oracle_upload", "sql_upload", "sql_download", "authenticate", "add_queue_item", "bulk_add_queue_item"]

        for name in sheet_names:
            sheet = workbook.create_sheet(name)
            sheet[f"A1"] = "key"
            sheet[f"B1"] = "value"

            if name == "SMTP":
                sheet[f"A2"] = "file_name"
                sheet[f"A3"] = "sender_email"
                sheet[f"A4"] = "to_email"
                sheet[f"A5"] = "email_subject"
                sheet[f"A6"] = "body_type"
                sheet[f"A7"] = "body"
                sheet[f"A8"] = "cc"
                sheet[f"A9"] = "attachments" 
                sheet[f"A10"] = "port" 
                sheet[f"A11"] = "server" 
        
            if name == "IMAP":
                sheet[f"A2"] = "IMAP_subject"
                sheet[f"A3"] = "IMAP_path"
        
            if name == "SFTP_download":
                sheet[f"A2"] = "SFTP_download_host"
                sheet[f"A3"] = "SFTP_download_user_name"
                sheet[f"A4"] = "SFTP_download_password"
                sheet[f"A5"] = "SFTP_download_port"
                sheet[f"A6"] = "SFTP_download_server_file_name"
                sheet[f"A7"] = "SFTP_download_local_file_path"

            if name == "SFTP_upload":
                sheet[f"A2"] = "SFTP_upload_host"
                sheet[f"A3"] = "SFTP_upload_user_name"
                sheet[f"A4"] = "SFTP_upload_password"
                sheet[f"A5"] = "SFTP_upload_local_file"
                sheet[f"A6"] = "SFTP_upload_target_location"
        
            if name == "oracle_download":
                sheet[f"A2"] = "oracle_download_host"
                sheet[f"A3"] = "oracle_download_port"
                sheet[f"A4"] = "oracle_download_data_source"
                sheet[f"A5"] = "oracle_download_user_ID"
                sheet[f"A6"] = "oracle_download_password"
                sheet[f"A7"] = "oracle_download_query"

            if name == "oracle_upload":
                sheet[f"A2"] = "oracle_upload_host"
                sheet[f"A3"] = "oracle_upload_port"
                sheet[f"A4"] = "oracle_upload_data_source"
                sheet[f"A5"] = "oracle_upload_user_ID"
                sheet[f"A6"] = "oracle_upload_password"
                sheet[f"A7"] = "oracle_upload_query"

            if name == "sql_upload":
                sheet[f"A2"] = "sql_upload_file_path"
                sheet[f"A3"] = "sql_upload_username"
                sheet[f"A4"] = "sql_upload_password"
                sheet[f"A5"] = "sql_upload_servername"
                sheet[f"A6"] = "sql_upload_database"
                sheet[f"A7"] = "sql_upload_table_name"

            if name == "sql_download":
                sheet[f"A2"] = "sql_download_file_path"
                sheet[f"A3"] = "sql_download_username"
                sheet[f"A4"] = "sql_download_password"
                sheet[f"A5"] = "sql_download_servername"
                sheet[f"A6"] = "sql_download_database"
                sheet[f"A7"] = "sql_download_table_name"

            if name == "authenticate":
                sheet[f"A2"] = "authenticate_tenancy_name"
                sheet[f"A3"] = "authenticate_username"
                sheet[f"A4"] = "authenticate_password"
                sheet[f"A5"] = "authenticate_url"

            if name == "add_queue_item":
                sheet[f"A2"] = "AddQueueItem_QueueName"
                sheet[f"A3"] = "AddQueueItem_X-UIPATH-TenantName"
                sheet[f"A4"] = "AddQueueItem_Authorization"
                sheet[f"A5"] = "AddQueueItem_Content"
                sheet[f"A6"] = "AddQueueItem_AddQueueItem_url"

            if name == "add_queue_item":
                sheet[f"A2"] = "BulkAddQueueItem_QueueName"
                sheet[f"A3"] = "BulkAddQueueItem_listItem"
                sheet[f"A4"] = "BulkAddQueueItem_Result"
                sheet[f"A5"] = "BulkAddQueueItem_bulk_url"

        default_sheet = workbook["Sheet"]
        workbook.remove(default_sheet)

        workbook.save(path)

    else:
        print("Excel file is already created")


# Let users know if they're missing any of our hard dependencies
_check_dependencies()

import tools
import orchestrator_api


