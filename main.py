from __future__ import print_function
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from googleapiclient.errors import HttpError
import httplib2
import io
from io import StringIO
from googleapiclient.http import MediaIoBaseDownload


import csv
from datetime import datetime, timedelta
import getpass
import logging
import os
import boto3
import re
import time

from config import *

def get_file_id(g_url):
    regex = "[-\w]{25,}(?!.*[-\w]{25,})"
    file_id = re.search(regex,g_url).group(0)

    return file_id


def remove_illegal_chars(filename):
    invalid = '<>:"/\|?*'

    for char in invalid:
        filename = filename.replace(char, '')
        
    return filename


def upload_to_s3(filename, folder):
    logging.info("Uploading to S3")
    result = s3.Bucket(BUCKET_NAME).upload_file(os.path.join(download_location, filename), '{}/{}/{}'.format(ROOT_FOLDER, folder, filename))
    #return result


def create_or_validate_creds():
    creds = None
    # The file token.json stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.json', 'w') as token:
            token.write(creds.to_json())
    
    return creds


#create datetime vars
today = datetime.now()
yesterday = datetime.now() - timedelta(days=1)
timestamp = today.strftime("%d-%m-%Y %H;%M")

#get username
username = getpass.getuser()

#collect path vars
root_folder = os.path.abspath(os.getcwd())

logs = os.path.join(root_folder, "logs")
download_location = os.path.join(root_folder, "temp")

if not os.path.exists(logs):
    # Create a new directory because it does not exist 
    os.makedirs(logs)

if not os.path.exists(download_location):
    # Create a new directory because it does not exist 
    os.makedirs(download_location)

sheet_url = input("Please paste the url of the google sheet:")

#collect batch size / limit
dl_batch_limit = 1

#setup logging
logging.basicConfig(filename = os.path.join(logs, timestamp + ".txt"), 
                    filemode = 'w', 
                    format = '%(asctime)s - %(message)s', 
                    datefmt = '%d-%b-%y %H:%M:%S', 
                    level = logging.DEBUG)
logging.getLogger().addHandler(logging.StreamHandler())

for name in ['boto', 'urllib3', 's3transfer', 'boto3', 'botocore', 'nose', 'boto3']:
    logging.getLogger(name).setLevel(logging.ERROR)

logging.info(download_location)


#deleteing old logs (-7 day)
logging.info("Looking for old logs...")
for i in os.listdir(logs):
    filename, file_extension = os.path.splitext(i)
    
    date_time_obj = datetime.strptime(filename, "%d-%m-%Y %H;%M")

    time_since_insertion = today - date_time_obj

    if time_since_insertion.days >= 7:
        logging.info("Deleting {}".format(i))
        os.remove(os.path.join(logs, i))



#Creating Session With Boto3.
session = boto3.Session(
aws_access_key_id = ACCESSKEY,
aws_secret_access_key = SECRETKEY
)

#Creating S3 Resource From the Session.
s3 = session.resource('s3')


# Set Google Drive API Tokens / Scopes
SCOPES = ['https://www.googleapis.com/auth/drive.readonly', 'https://www.googleapis.com/auth/spreadsheets.readonly', 'https://www.googleapis.com/auth/documents.readonly', 'https://www.googleapis.com/auth/documents'] 

creds = create_or_validate_creds()


# Run it
total_expected_files = 0

#if len(os.listdir(source_csv_folder)) != 1:
#    logging.error("Please make sure there is only 1 source csv file in the source csv folder")
#    exit()

#----

SHEET_ID = get_file_id(sheet_url)

service = build('sheets', 'v4', credentials=creds)

spreadsheet = service.spreadsheets().get(spreadsheetId=SHEET_ID, includeGridData=True).execute()

for sheet in spreadsheet['sheets']:
    sheet_title = sheet['properties']['title']
    logging.info(sheet_title)


sheet = service.spreadsheets()

upload_folder = str(sheet.get(spreadsheetId=SHEET_ID, fields="properties/title").execute()["properties"]["title"])
logging.info("Spreadsheet: {}".format(upload_folder))


result = sheet.values().get(spreadsheetId=SHEET_ID, range="{}!A:Z".format(str(sheet_title))).execute()
values = result.get('values', [])

line_count = 1
cell_count = 0

failed_dl_dict = {}

for row in values:
    creds = create_or_validate_creds()

    title = row[0]
    for cell in row:
        logging.info("-" * 50)
        logging.info("Row: {}".format(line_count))
        logging.info("Cell: {}".format(cell))

        try:
            if get_file_id(cell):
                file_id = get_file_id(cell)
                logging.info("File ID: {}".format(str(file_id)))
        except AttributeError:
            logging.info("Nothing to download")
            continue
        except Exception as e:
            logging.error(e)
            continue

        current_dl_count = len(os.listdir(download_location))

        if len(os.listdir(download_location)) > 0:
            logging.error("Please clear the temp upload folder!")
            exit()
        
        if "document" in cell:
            total_expected_files += 1
            
            try:
                downloadService = build('drive', 'v3', credentials=creds)
                results = downloadService.files().get(fileId=file_id, fields="id, name,mimeType,createdTime", supportsAllDrives=True).execute()
                original_assetname = results['name']
                logging.info("Filename: {}".format(original_assetname))
                assetname = remove_illegal_chars(original_assetname)
                logging.info("Formatted filename: {}".format(assetname))
                docMimeType = results['mimeType']

                mimeTypeMatchup = {
                "application/vnd.google-apps.document": {
                    "exportType":"application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    "docExt":"docx"
                    },
                "application/vnd.openxmlformats-officedocument.wordprocessingml.document": {
                    "exportType":"application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                    "docExt":"docx"
                    }
                }

                exportMimeType =mimeTypeMatchup[docMimeType]['exportType']
                docExt =mimeTypeMatchup[docMimeType]['docExt']

                if docMimeType == "application/vnd.google-apps.document":
                    request = downloadService.files().export_media(fileId=file_id, mimeType=exportMimeType) # Export formats : https://developers.google.com/drive/api/v3/ref-export-formats
                else:
                    request = downloadService.files().get_media(fileId=file_id)

                fh = io.FileIO(os.path.join(download_location, assetname+"."+docExt), mode='w')

                downloader = MediaIoBaseDownload(fh, request)
                done = False
                while done is False:
                    status, done = downloader.next_chunk()
                    logging.info("Download %d%%." % int(status.progress() * 100))
                fh.close()

                for f in os.listdir(download_location):
                    if not f.endswith(".crdownload"):
                        logging.info("Uploading to S3...")
                        upload_to_s3(f, upload_folder)
                        time.sleep(10)
                        os.remove(os.path.join(download_location, f))

                cell_count += 1

            except Exception as e:
                logging.error(e)
                if 'original_assetname' in locals():
                    failed_dl_dict[original_assetname] = cell
                else:
                    cell_index = row.index(cell)
                    failed_dl_dict["ROW:{}, CELL:{}".format(line_count, cell_index)] = cell
        
        elif "file" in cell:
            total_expected_files += 1

            try:            
                downloadService = build('drive', 'v3', credentials=creds)
                results = downloadService.files().get(fileId=file_id, fields="id, name", supportsAllDrives=True).execute()
                
                original_assetname = results['name']
                logging.info("Filename: {}".format(original_assetname))
                assetname = remove_illegal_chars(original_assetname)
                logging.info("Formatted filename: {}".format(assetname))
                
                request = downloadService.files().get_media(fileId=file_id)
                fh = io.FileIO(os.path.join(download_location, assetname), 'wb') 
                downloader = MediaIoBaseDownload(fh, request)
                done = False
                while done is False:
                    status, done = downloader.next_chunk()
                    logging.info("Download %d%%." % int(status.progress() * 100))
                fh.close()

                for f in os.listdir(download_location):
                    if not f.endswith(".crdownload"):
                        upload_to_s3(f, upload_folder)
                        time.sleep(10)
                        os.remove(os.path.join(download_location, f))
                
                cell_count += 1

            except Exception as e:
                logging.error(e)
                if 'original_assetname' in locals():
                    failed_dl_dict[original_assetname] = cell
                else:
                    cell_index = row.index(cell)
                    failed_dl_dict["ROW:{}, CELL:{}".format(line_count, cell_index)] = cell

        
    
    line_count += 1



logging.info("-" * 50)
logging.info("Finished!")
logging.info(f'Downloaded / Uploaded {cell_count} assets out of {total_expected_files}.')

if len(failed_dl_dict) > 0:
    logging.warning("{} assets failed to be ingested!".format(str(len(failed_dl_dict))))
    failed_count = 1
    for key, value in failed_dl_dict.items():
        if str(key).startswith("ROW:"):
            logging.warning("{}. Filename unknown: {}".format(str(failed_count), str(key)))
        else:
            logging.warning("{}. Filename: {}".format(str(failed_count), str(key)))
        logging.warning("{}. URL: {}".format(str(failed_count), str(value)))

        failed_count += 1