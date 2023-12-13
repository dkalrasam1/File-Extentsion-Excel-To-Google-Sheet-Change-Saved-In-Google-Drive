from __future__ import print_function

import os.path

from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from apiclient.http import MediaFileUpload
from apiclient.http import MediaIoBaseDownload
import io
import shutil
import time
import re

# If modifying these scopes, delete the file token.json.
SCOPES = ['https://www.googleapis.com/auth/drive']
location = '/home/xxxxx/xls-to-sheet/files/'
#location = '/var/www/html/drive_files_conversion/files/'

def login():
    creds = None
    # The file token.json stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    
    # if os.path.exists('token.json'):
    #     creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    if os.path.exists('/home/xxxxx/xls-to-sheet/token.json'):
        creds = Credentials.from_authorized_user_file('/home/xxxxx/xls-to-sheet/token.json', SCOPES)
        
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file('/home/xxxxx/xls-to-sheet/client.json', SCOPES)
            #flow = InstalledAppFlow.from_client_secrets_file('client_secret_255041354914-4dis5bgv1v9lu6clhk8577mm591onu46.apps.googleusercontent.com.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.json', 'w') as token:
            token.write(creds.to_json())
        
    service = build('drive', 'v3', credentials=creds)
    shutil.rmtree(location)
    os.mkdir(location)
    main('XXXXXXXXXx' ,service) #xxxxxxxxxx id of folder to start scanning with

def main(folderName, service):
    

    try:
        #print("'"+folderName+"' in parents and name != 'converted-files'")
        # Call the Drive v3 API
        results = service.files().list(
            #q="mimeType = 'application/vnd.google-apps.folder'",
            q = "'"+folderName+"' in parents",
            pageSize=1000, fields="nextPageToken, files(id, name,mimeType,id)").execute()
        items = results.get('files', [])

        if not items:
            print('No files found.')
            return
        #print('Files:')
        for item in items:

            #print(u'{0} ({1})'.format(item['name'],item['id']))
            if item['mimeType'] == 'application/vnd.ms-excel': #for excel file
                downloadFile(service, item['id'], item['mimeType'], item['name'])
                continue
            if item['mimeType'] == 'application/vnd.google-apps.folder': #for folder
                #print(item['name'])
                if item['name'] == 'Google Sheets' or item['name'] == 'converted-files':
                    #print(item['name']+'------skip')
                    continue
                checkNumberInName = re.search(r'\d+', item['name'])
                if checkNumberInName is not None:
                    checkName = int(checkNumberInName.group())
                    if(checkName >= 2018):#read all folders greater than 2018  
                        main(item['id'], service)
        
    except HttpError as error:
        # TODO(developer) - Handle errors from drive API.
        print(f'An error occurred: {error}')



def downloadFile(service, fileId, mimeType, fileName):
    file_id = fileId
    request = service.files().get_media(fileId=file_id)
    fh = io.FileIO(location+fileName, 'wb')
    downloader = MediaIoBaseDownload(fh, request)
    done = False
    while done is False:
        status, done = downloader.next_chunk()
        print(status.progress() * 100)

    fh.seek(0)
    # Write the received data to the file
    # with open(fileName, 'wb') as f:
    #     shutil.copyfileobj(fh, f)

    file_metadata = {
        'name': fileName,
        "parents" : ['XXXXXXXXXXXXXXXXX'], #convert-files---- folder's id
        'mimeType': 'application/vnd.google-apps.spreadsheet'
    }
    media = MediaFileUpload(location+fileName,
                            mimetype=mimeType,
                            resumable=True)
    try:
        service.files().delete(fileId=file_id).execute()
        print('file deleted')
    except HttpError as error:
        print(error)
        

    file = service.files().create(body=file_metadata,media_body=media).execute()
    print('upload done')
    time.sleep(5) 
    return "true"


if __name__ == '__main__':
    login()