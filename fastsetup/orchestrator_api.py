import requests

def check_type(_argument, _type, _function_name):
    if not isinstance(_argument, _type):
        raise TypeError("fastsetup "+ _function_name +" - wrong data type passed in: " + str(_type))

def authenticate(_config:dict):
        check_type(_config, dict, 'authenticate')
    
        payload="{\r\n    \"TenancyName\": \""+_config['authenticate_tenancy_name']+"\",\r\n    \"UsernameOrEmailAddress\": \""+_config['authenticate_username']+"\",\r\n    \"Password\": \""+_config['authenticate_password']+"\"\r\n}"
        headers = {
            'Content-Type': 'application/json',
            }
        print("sending request")
        response = requests.request("POST", _config['authenticate_url'], headers=headers, data=payload, verify=False)
        print("requist succssefully")
        return response


def add_queue_item(_config:dict,row):

    check_type(_config, dict, 'add_queue_item')

    payload="{'itemData': {'Priority': 'High','Name': '"+_config['AddQueueItem_QueueName']+"','SpecificContent': {   'Name': '"+str(row["Name"])+"', 'OST Shop Id': '"+str(row["OST Shop Id"])+"', 'Bank': '"+str(row["Bank"])+"', 'Company': '"+str(row["Company"])+"', 'Payment': '"+str(row["Payment"])+"'}}}"
    headers = {
            'X-UIPATH-TenantName': _config['AddQueueItem_X-UIPATH-TenantName'],
            'Authorization': _config['AddQueueItem_Authorization'],
            'Content-Type': _config['AddQueueItem_Content-Type']
            }
    response = requests.request("POST", _config['AddQueueItem_AddQueueItem_url'], headers=headers, data=payload, verify=False)
    return response


def bulk_add_queue_item(_config:dict):
    check_type(_config, dict, 'bulk_add_queue_item')

    payload="{'queueName': '"+_config['BulkAddQueueItem_QueueName']+"','commitType': 'AllOrNothing','queueItems': "+str(_config['BulkAddQueueItem_listItem'])+"}"
    headers = {
            'X-UIPATH-TenantName': 'OrangeProduction',
            'Authorization': 'Bearer '+ _config['BulkAddQueueItem_Result']+'',
            'Content-Type': 'application/json'
            }
    response = requests.request("POST", _config['BulkAddQueueItem_bulk_url'], headers=headers, data=payload, verify=False)
    return response