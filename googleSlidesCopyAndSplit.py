# Script to split a Google Slides presentation into different sections.
# Utilizes the Google Slides API and Google Drive API to make copies of
# certain sections of an existing Google Slides presentation. Useful in
# separating content shared among teachers for use in Google Classrooms.

from __future__ import print_function
import pickle
import os.path
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request

# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/presentations']
SCOPES2 = ['https://www.googleapis.com/auth/drive']

# The ID of a presentation
PRESENTATION_ID = '1EAYk18WDjIG-zp_0vLm3CsfQh_i8eXc67Jo2O9C6Vuc'

def auth_flow(file_name1, file_name2):
    # The files token.pickle and tokekn2.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first time
    if os.path.exists('token.pickle'):
        with open('token.pickle', 'rb') as token:
            creds_slides = pickle.load(token)
    # If there are no (valid) credentials available, let the user log in.
    if not creds_slides or not creds_slides.valid:
        if creds_slides and creds_slides.expired and creds_slides.refresh_token:
            creds_slides.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                file_name1, SCOPES)
            creds_slides = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.pickle', 'wb') as token:
            pickle.dump(creds_slides, token)

    if os.path.exists('token2.pickle'):
        with open('token2.pickle', 'rb') as token:
            creds_drive = pickle.load(token)
    # If there are no (valid) credentials available, let the user log in.
    if not creds_drive or not creds_drive.valid:
        if creds_drive and creds_drive.expired and creds_drive.refresh_token:
            creds_drive.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                file_name2, SCOPES2)
            creds_drive = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token2.pickle', 'wb') as token:
            pickle.dump(creds_drive, token)
    return creds_slides, creds_drive


# Create a new folder in Google Drive
def make_new_folder(drive_service):
    file_metadata = {
        'name': 'Sample Folder',
        'mimeType': 'application/vnd.google-apps.folder'
    }
    folder = drive_service.files().create(body=file_metadata,
                                        fields='id').execute()
    return folder.get('id')


# Share a Google Drive folder with another account. Give other account write permissions.
def share_folder(folder_id, drive_service):
    batch = drive_service.new_batch_http_request()
    user_permission = {
        'type': 'user',
        'role': 'writer',
        'emailAddress': 'sampleEmail@sample.com'
    }
    batch.add(drive_service.permissions().create(
        fileId=folder_id,
        body=user_permission,
        fields='id',
    ))
    batch.execute()


# Make a copy of a Google Slides Presentation in your Drive
def make_presentation_copy(drive_service, file_num, folder_id):
    body = {
        'name': "Sample Presentation Copy" + ' Week-' + str(file_num+1),
         'parents': [folder_id]
    }
    copy_presentation = drive_service.files().copy(
        fileId=PRESENTATION_ID, body=body).execute()
    return copy_presentation


# Ignore certain pages in all copies
def init_delete_requests(copy_slides):
    requests = []
    for i in range(0,4):
            first_slides = {
                "deleteObject": {
                        "objectId": copy_slides[i].get('objectId')
                } 
            }
            requests.append(first_slides);

    last_slide = {
        "deleteObject": {
            "objectId": copy_slides[len(copy_slides)-1].get('objectId')
        } 
    }
    requests.append(last_slide);
    return requests


def main():
    creds_slides = None
    creds_drive = None
    
    creds_slides, creds_drive = auth_flow('credential_slides.json', 'credentials_drive.json')

    # Slides API and Drive API
    slide_service = build('slides', 'v1', credentials=creds_slides)
    drive_service = build('drive', 'v3', credentials=creds_drive)

    # Create a folder
    folder_id = make_new_folder(drive_service)

    # Share folder
    share_folder(folder_id, drive_service)

    # Get the presentation to be copied
    presentation = slide_service.presentations().get(
        presentationId=PRESENTATION_ID).execute()
    

    for i in range(0, 20):
        
        copy_presentation = make_presentation_copy(drive_service, i, folder_id)
        presentation_copy_id = copy_presentation.get('id')
        copy_slides = slide_service.presentations().get(presentationId=presentation_copy_id).execute().get('slides')
        print(copy_slides)
        requests = init_delete_requests(copy_slides)
        
        if i < 13:
            start_range = i*6
            end_range = (i+1)*6
        elif i == 13:
            start_range = i*6
            end_range = (i+1)*6 -1
        else:
            start_range = i*6 -1
            end_range = (i+1)*6 -1

        important_slides = copy_slides[4:123]
        temp_slides = important_slides[:start_range] + important_slides[end_range:]

        for slide in temp_slides: 
            temp = {
                "deleteObject": {
                    "objectId": slide.get('objectId')
                } 
            }
            requests.append(temp)

        batch_body = {
            "requests": requests
        }
        print(len(requests));
        print(requests)

        slide_service.presentations().batchUpdate(
            presentationId=presentation_copy_id, body=batch_body).execute()

    
if __name__ == '__main__':
    main()





    

