## Moonbug Google sheet download links to S3
This script detects google drive download links in a google sheet, downloads them and uploads them to a bucket in S3 inside a folder named after the spreadsheet.
1. User submits a google sheet URL
1. Search each cell for google download URLs
1. Download the asset to a holding folder
1. Upload to a bucket in S3 inside a sub-folder named after the spreadsheet name
1. Delete the file in the holding folder
1. Print a count of assets successfully ingested and a report of files which were unable to be ingested (filename and URL)

### Installation
In order for the script to run you'll need to create the following unversioned files:
- config.py
- credentials.json

**Creating the config file**:
1. Create a new file named 'config.py' in the root directory.
1. Open the file and add the following details:  
ACCESSKEY = *"AWS ACCESS KEY"*  
SECRETKEY = *"AWS SECRET KEY"*  
BUCKET_NAME = *"THE BUCKET NAME"*  
ROOT_FOLDER = *"THE ROOT FOLDER NAME"*  

Create a user with correct S3 permissions in AWS IAM and generate an access and secret key. 

Add the bucket name and root folder name's tothe config file and save.

**Generating credentials.json**:  
Follow the first part of this tutorial to create credentials file:
https://www.casuallycoding.com/download-docs-from-google-drive-api/

### Steps to run
1. Download codebase 
1. Run setup.bat to install python libraries
1. Run START.bat to run the python script
1. Paste your google sheet URL into the console window
1. Once it has completed you will be presented with a report of successes and any failures.
