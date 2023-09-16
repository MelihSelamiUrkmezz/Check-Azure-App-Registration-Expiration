from azure.identity import ClientSecretCredential
from msgraph.core import GraphClient
from datetime import datetime
import json
import msal
import base64
import os
import requests
import csv
import pandas as pd
import sys

tenant_supp=["AS","DTG1","DTG2","INO","INS","OTOG1","OTOS","TS","TR","VDF","ZE","ZI","DR"]
userid_index = sys.argv.index("--USERID")
userid=sys.argv[userid_index+1]
tenant_ids=[]
client_ids=[]
client_secrets=[]
for sup in tenant_supp:
    tenant_index = sys.argv.index("--"+sup+"TENANTID")
    clientid_index = sys.argv.index("--"+sup+"CLIENTID")
    client_index = sys.argv.index("--"+sup+"CLIENTSECRET")
    tenant_id = sys.argv[tenant_index+1]
    client_id = sys.argv[clientid_index+1]
    client_secret = sys.argv[client_index+1]
    tenant_ids.append(tenant_id)
    client_ids.append(client_id)
    client_secrets.append(client_secret)

# The name of the excel and sheet name you want to read
df = pd.read_excel('', sheet_name='')

tenant_to_app_dict = {}
tenant_count=0
day_threshold=10 #How many days before the e-mail should be sent?

for index, row in df.iterrows():
    tenant_id = row['Tenant ID']
    app_id = row['App ID']
    
    if tenant_id not in tenant_to_app_dict:
        tenant_to_app_dict[tenant_id] = [app_id]
    else:
        tenant_to_app_dict[tenant_id].append(app_id)
        

graph_url = 'https://graph.microsoft.com/v1.0'

rows=[]
def send_mail(file_path,to):
    tenant_id=tenant_ids[0]
    client_secret=client_secrets[0]
    client_id=client_ids[0]
    userId="" #The e-mail address or object id to whom the e-mail will be sent is required.
    sending_file_name="" # Name of the file to be sent (ex: app.csv)
    authority = f"https://login.microsoftonline.com/{tenant_id}"
    csv_file_path = file_path
    with open(csv_file_path, "rb") as file:
        csv_content = base64.b64encode(file.read()).decode()
    app = msal.ConfidentialClientApplication(
        client_id=client_id,
        client_credential=client_secret,
        authority=authority)

    scopes = ["https://graph.microsoft.com/.default"]

    result = None
    result = app.acquire_token_silent(scopes, account=None)

    if not result:
        print(
            "No suitable token exists in cache. Let's get a new one from Azure Active Directory.")
        result = app.acquire_token_for_client(scopes=scopes)

    if "access_token" in result:
        endpoint = f'https://graph.microsoft.com/v1.0/users/{userId}/sendMail'
        toUserEmail = to
        email_msg = {
        'Message': {
            'Subject': "App Registration Keys Expiring Soon Alert secret",
            'Body': {
                'ContentType': 'Text',
                'Content': "The secrets in the registers of these apps have expired or are very close to being expired.",
            },
            'ToRecipients': [{'EmailAddress': {'Address': toUserEmail}}],
            'Attachments': [
                {
                    "@odata.type": "Microsoft.Graph.FileAttachment",
                    "Name": sending_file_name,
                    'ContentBytes': csv_content,
                }
            ],
        },
        'SaveToSentItems': 'true',
    }
        r = requests.post(endpoint,
                        headers={'Authorization': 'Bearer ' + result['access_token']}, json=email_msg)
        if r.ok:
            print('Sent email successfully')
        else:
            print(r.json())
    else:
        print(result.get("error"))
        print(result.get("error_description"))
        print(result.get("correlation_id"))


def calculate_date_difference(target_date_str):
    target_date_format = "%Y-%m-%d"
    target_date = datetime.strptime(target_date_str, target_date_format)
    current_date = datetime.utcnow()
    date_difference = target_date - current_date
    return date_difference.days

for tenant_id, app_ids in tenant_to_app_dict.items():
    try:
        index=tenant_ids.index(tenant_id)
        token_url = f'https://login.microsoftonline.com/{tenant_ids[index]}/oauth2/token'
        token_data = {
            'grant_type': 'client_credentials',
            'client_id': client_ids[index],
            'client_secret': client_secrets[index],
            'resource': 'https://graph.microsoft.com'
        }
        token_r = requests.post(token_url, data=token_data)
        token = token_r.json().get('access_token')
        row_count=0
        apps_url = f'{graph_url}/applications?$select=appId,displayName,passwordCredentials'
        all_apps = []
        required_apps=app_ids
        while apps_url:
            apps_response = requests.get(
                apps_url,
                headers={'Authorization': 'Bearer ' + token}
            )
            
            if apps_response.status_code == 200:
                apps_data = apps_response.json()
                all_apps.extend(apps_data.get('value', []))
                apps_url = apps_data.get('@odata.nextLink')
            else:
                print(f'Error Code: {apps_response.status_code}, Error Message: {apps_response.text}')
                break
        for x in range(len(all_apps)):
            app_id=all_apps[x]
            application_id=app_id['appId']
            application_name=app_id['displayName']
            if  app_id:
                if len(app_id['passwordCredentials'])==1:
                    given_date =app_id['passwordCredentials'][0]['endDateTime']
                    difference_in_days = calculate_date_difference(str(given_date.split('T')[0]))+1
                    if (difference_in_days>=0 and difference_in_days <day_threshold and application_id in required_apps):
                        row_count+=1
                        row=[tenant_ids[index],application_id,application_name,difference_in_days,app_id['passwordCredentials'][0]['keyId']]
                        rows.append(row)
                elif(len(app_id['passwordCredentials'])>1):
                    for m in range(len(app_id['passwordCredentials'])):
                        given_date =app_id['passwordCredentials'][m]['endDateTime']
                        difference_in_days = calculate_date_difference(str(given_date.split('T')[0]))+1
                        if (difference_in_days>=0 and difference_in_days <day_threshold and application_id in required_apps):
                            row_count+=1
                            row=[tenant_ids[index],application_id,application_name,difference_in_days,app_id['passwordCredentials'][0]['keyId']]
                            rows.append(row)
    except: 
        print("Tenant ID not found!")
    

with open(r'C:\app_registration-expiry.csv', mode='w', newline='') as dosya:
        csv_writer = csv.writer(dosya)
        csv_writer.writerow(['Tenant ID','App ID','Application Name','Remaining Day', 'Key ID'])
        csv_writer.writerows(rows)


if(row_count>0):
    send_mail(r"C:\app_registration-expiry.csv","") #On the Azure side, you can enter the e-mail address or object id to whom you want to send the e-mail.
    send_mail(r"C:\app_registration-expiry.csv","") #On the Azure side, you can enter the e-mail address or object id to whom you want to send the e-mail.

        

            

