from firebase_admin import credentials
from firebase_admin import storage
import os
import firebase_admin
import requests
import openpyxl
import re



project_id = 'heimdall-333'
key_file = '/Users/iar/Documents/Heimdall/python/....json'

# Load the workbook
wb = openpyxl.load_workbook('/Users/iar/Documents/.../excel_database/data_base.xlsx')

# Select the first sheet in the workbook
worksheet = wb.worksheets[1]



cred = credentials.Certificate(key_file)
# Initialize the app with a service account, granting admin privileges
app = firebase_admin.initialize_app(cred, {
    'storageBucket': '....appspot.com'
})
bucket = storage.bucket()

# Get a reference to the storage bucket
bucket = storage.bucket(app=app)

# Iterate through all the blobs in the bucket

blobs = bucket.list_blobs(prefix='Real Estate/')
x=0
for blob in blobs:
    x+=1
    name_pic = blob.name.replace('Real Estate/','').lower()
    blob_name = blob.name
    print(blob_name)
    


    # Cellule à editer
    cell = worksheet.cell(row=x, column=14)
    
    blob = bucket.get_blob(blob_name)
    metadata = blob.metadata
    token = metadata['firebaseStorageDownloadTokens']
    name_database = blob.name.replace("Real Estate/", "")
    url = 'https://firebasestorage.googleapis.com/..../'+'Real%20Estate%2F'+name_database+'?alt=media&token='+token
    print(url)



    write_string = url
    cell.value = write_string
    # Save the changes to the workbook
    wb.save('/Users/iar/Documents/.../excel_database/data_base_photo_tempo2.xlsx')

# Close the workbook
wb.close()

