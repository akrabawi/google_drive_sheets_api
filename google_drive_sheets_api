from __future__ import print_function
import httplib2
from apiclient import errors
import os, io
import pandas as pd
import apiclient
import pprint
from apiclient import discovery
from google.oauth2 import service_account
from oauth2client import client
from oauth2client import tools
from oauth2client.service_account import ServiceAccountCredentials
from oauth2client.file import Storage
import json
import gspread
import pprint
from urllib.request import Request
from apiclient.http import MediaFileUpload, MediaIoBaseDownload

##############################################################################################################
##############################################################################################################
###################################### Google APIs Authentication ############################################
##############################################################################################################
##############################################################################################################

'''
---- (IN API Console) ----
1. First, you need to make sure to create Google API (Project).
2. Make sure to (Enable) any services to (Libraries).
2. Secondly, in (Credentials) module, you are required to create Credentials (.JSON) file. 
        --> Must be a (Service Key Account) with (Editor/Organizer/Owner) priviliges.
---- (IN Code) ----
3. Specify the (SCOPE) of your API access, and get path to (.JSON) file.
3. Then to use API, you need to create (credentials) object [below] using (SCOPE)&(.JSON).
4. Lastly, create a (service/client) by authorizing it using the created (credentials object).
'''

##############################################################################################################
##############################################################################################################

'''
Create authorization class:
The object should take (SCOPE) and (.JSON) credential file
'''
class auth:
    # Create constructor
    def __init__(self,SCOPES,CLIENT_SECRET_FILE):
        self.SCOPES = SCOPES
        self.CLIENT_SECRET_FILE = CLIENT_SECRET_FILE
        
    # get the (credentials) and create (client) and return it
    def getCredentials(self):
        # Read (JSON) file into here
        with open(self.CLIENT_SECRET_FILE) as f:
            secrets = json.load(f)
        # Create (Credentials) object
        credentials = service_account.Credentials.from_service_account_file(self.CLIENT_SECRET_FILE, 
                                                                            scopes=self.SCOPES)
        return credentials

##############################################################################################################
##############################################################################################################
########################################## Google Drive APIs #################################################
##############################################################################################################
##############################################################################################################

'''
Create google_drive class:
The object should take (SERVICE_DRIVE) object
'''

class google_drive:
    # Create constructor
    def __init__(self,SERVICE_DRIVE):
        self.drive_service = SERVICE_DRIVE
    
    ########################################################################################################
    ########################################################################################################

    '''
    This function will return a list of files in Google Drive with their respected file IDs
    set (printList = True), if you want to print the files names and IDS
    '''

    def list_files(self,*positional_parameters, **keyword_parameters):
        # Make API request 
        results = self.drive_service.files().list(
            pageSize=500,fields="nextPageToken, files(id, name)").execute()

        # Extract the files and thier IDs
        items = results.get('files', [])

        # Convret them into a dictionary with fileID as key and fileName as values
        files = {}
        if not items:
            files = {}
        else:
            for item in items:
                files[item['id']] = item['name']

        # If (printList) parameter have been passed, then the function will print the dictionary
        if (
            (
                ('printList' in keyword_parameters) & 
                (len(keyword_parameters) == 1)
            ) |
            (
                ('printList' in keyword_parameters) & 
                ('printPermissions' in keyword_parameters)
            )
        ):
            if(
                (keyword_parameters['printList'] == True) & 
                (keyword_parameters['printPermissions'] == False)
            ):
                i = 1
                for fileId in sorted(files):
                    print("%i. %-20s  ==> \t%-20s"%(i,files[fileId], fileId))
                    i = i+1

        # If (permissions) is also added and is (True), then print the permissions of each file
        if (
            ('printList' in keyword_parameters) & 
            ('printPermissions' in keyword_parameters) & 
            (len(keyword_parameters) == 2)
        ):
            if(
                (keyword_parameters['printList'] == True) & 
                (keyword_parameters['printPermissions'] == True) 
            ):
                i = 1
                for fileId in sorted(files):
                    print("%i. %-20s  ==> \t%-20s"%(i,files[fileId], fileId))
                    permission_list = self.list_file_permission(fileId)
                    for perm in permission_list:
                        if(perm['user_email'] != -1):
                            print("\t%-15s  ==> %-10s: [%s]"%(perm['id'],perm['role'],perm['user_email']))
                    print('\n')
                    i = i+1
        # Finally, return the dictionary of files    
        return files

    ########################################################################################################
    ########################################################################################################

    '''
    This function will create a Google Spreadsheet in Google Drive.
    You can overwirte an existing file of same name by setting (overwrite = True)
        ** The spreadsheet will have a sheet with a name similar to (fileName)
    '''

    def create_spreadsheet(self,df,fileName,*positional_parameters, **keyword_parameters):
        if ('overwrite' in keyword_parameters):
            if(keyword_parameters['overwrite']):
                self.delete_file(fileName=fileName)
            else:
                print('WARNING! Overwrite is FALSE! Duplicates may occur')
        else:
            print('WARNING! No overwrite! Duplicates may occur')

        # Change the DataFrame to (csv)
        df.to_csv(fileName)

        # This helps us convert any csv file into Google Sheets
        file_metadata = {
            'name': fileName,
            'mimeType': 'application/vnd.google-apps.spreadsheet'
        }

        # Create the body of request
        media = MediaFileUpload(fileName,
                                mimetype='text/csv',
                                resumable=True)

        # Execute API to create Google sheet
        file = self.drive_service.files().create(body=file_metadata,
                                            media_body=media,
                                            fields='id').execute()

        # Print and Return file ID
        print('File Successfully Created! File ID ==>  ', file.get('id'))
        return file.get('id')

    ########################################################################################################
    ########################################################################################################
    
    '''
    This function will delete any file with passed file ID
    '''

    def delete_file_id(self,fileId):
        return self.drive_service.files().delete(fileId=fileId).execute()

    ########################################################################################################
    ########################################################################################################
    
    '''
    This function will delete any file with passed file ID
    '''

    def delete_file(self,*positional_parameters, **keyword_parameters):
        if (
            ('fileName' in keyword_parameters) & 
            ('fileId' in keyword_parameters) & 
            (len(keyword_parameters) != 1)
        ):
            raise Exception('Please use either (fileName) or (fileID) EXECLUSIVILY')
            return None
        elif (
            (('fileName' in keyword_parameters)) & 
            (not('fileId' in keyword_parameters)) 
        ):
            if(keyword_parameters['fileName']):
                return self.delete_file_name(keyword_parameters['fileName'])
        elif (
            (not('fileName' in keyword_parameters)) & 
            (('fileId' in keyword_parameters)) 
        ):
            if(keyword_parameters['fileId']):
                return self.delete_file_id(keyword_parameters['fileId'])
        else:
            raise Exception('Please enter a file indicator: Either (fileName) or (fileID) EXECLUSIVILY')
            return None
    
    ########################################################################################################
    ########################################################################################################
    
    '''
    This function will delete all the files in Google drive that have the same passed name 
    '''

    def delete_file_name(self,fileName):
        # The process involves searching for file with given name.
        # If files with same name are available, then we need to delete previous versions

        # First list all the files in Google Drive
        files_list = self.list_files()

        # Check if file exists in Google Drive
        if(fileName in files_list.values()): 
            # loop through all the files and delete the matched file names
            for i in files_list.keys():
                if (fileName == files_list[i]):
                    print('File name', fileName, ' with file ID ', i,' have been found and deleted!')
                    self.delete_file_id(i)
            return None
        else:
            print('Name has NOT been Found!')
            return None

    ########################################################################################################
    ########################################################################################################
    
    '''
    This function will take pandas DataFrame and create a (csv) file of it in Google Drive
    You can overwirte an existing file of same name by setting (overwrite = True)
    '''

    def create_csv(self,df,fileName,*positional_parameters, **keyword_parameters):
        if ('overwrite' in keyword_parameters):
            if(keyword_parameters['overwrite']):
                self.delete_file(fileName=fileName)
            else:
                print('WARNING! Overwrite is FALSE! Duplicates may occur')
        else:
            print('WARNING! No overwrite! Duplicates may occur')

        # Change the DataFrame to (csv)
        df.to_csv(fileName)
        # This helps us convert any csv file into Google Sheets
        file_metadata = {
            'name': fileName,
        }

        # Create the body of request
        media = MediaFileUpload(fileName,
                                mimetype='text/csv',
                                resumable=True)

        # Execute API to create Google sheet
        file = self.drive_service.files().create(body=file_metadata,
                                            media_body=media,
                                            fields='id').execute()

        # Print and Return file ID
        print('File Successfully Created! File ID ==>  ', file.get('id'))
        return file.get('id')


    ########################################################################################################
    ########################################################################################################
    
    '''
    This function is used to read (csv) file from Google Drive and return a pandas DataFrame for it
    '''

    def read_csv(self,*positional_parameters, **keyword_parameters):
        if (
            ('fileName' in keyword_parameters) & 
            ('fileId' in keyword_parameters) & 
            (len(keyword_parameters) != 1)
        ):
            raise Exception('Please use either (fileName) or (fileID) EXECLUSIVILY')
            return None
        elif (
            ('fileName' in keyword_parameters) &
            (not('fileId' in keyword_parameters))
        ):
            if(keyword_parameters['fileName']):
                files_list = self.list_files()
                files_list = dict(zip(files_list.values(),files_list.keys()))
                fileId = files_list[keyword_parameters['fileName']]
                return self.read_csv_id(fileId)
        elif (
            (not('fileName' in keyword_parameters)) & 
            ('fileId' in keyword_parameters)
        ):
            if(keyword_parameters['fileId']):
                return self.read_csv_id(keyword_parameters['fileId'])
        else:
            raise Exception('Please enter a file indicator: Either (fileName) or (fileID) EXECLUSIVILY')
            return None

    ########################################################################################################
    ########################################################################################################
    
    '''
    This function is used to read (csv) file from Google Drive using file ID and return it as a DataFrame
    '''

    def read_csv_id(self,fileId):
        # Initiate the request and proceed it
        request = self.drive_service.files().get_media(fileId=fileId)
        fh = io.BytesIO()
        downloader = MediaIoBaseDownload(fh, request)
        
        # Get the fileName
        files_list = self.list_files()
        fileName = files_list[fileId]
        
        # Read the ByteIO into a string
        done = False
        while done is False:
            status, done = downloader.next_chunk()
        with io.open(fileName,'wb') as f:
            fh.seek(0)
            inFile = fh.read()

        # Convert the string into DataFrame
        df = pd.read_csv(pd.compat.StringIO(inFile.decode('utf-8')), sep=',', header=None)
        df.columns = df.iloc[0]
        df = df.reindex(df.index.drop([0]))
        df = df.loc[:, df.columns.notnull()]
        return df
    
    ########################################################################################################
    ########################################################################################################
    
    
    '''
    This function provides a permission for (user_account) to [read/edit/orginze/own] a file with given (fileId)
    Permission type (perm_type) can either be: [user / group / domain / anyone]
        ==> HOWEVER, the function is made to set permission for (user) ONLY.
    Permission role (perm_role) can either be: [owner / organizer / fileOrganizer / writer / commenter / reader]
    '''

    def add_permission(self,fileId,perm_type,perm_role,user_account):
        def callback(request_id, response, exception):
            if exception:
                # Handle error
                print(exception)
            else:
                print("Permission Id: %s" % response.get('id'))
        batch = self.drive_service.new_batch_http_request(callback=callback)

        # Set the permission body
        user_permission = {
            'type': perm_type,
            'role': perm_role,
            'emailAddress': user_account
        }
        
        # Apply the API to set and send permission
        batch.add(self.drive_service.permissions().create(fileId=fileId,
                                                          body=user_permission,
                                                          fields='id',
        ))
        batch.execute()


    ########################################################################################################
    ########################################################################################################
    
    '''
    This function delete permission for given (fileId) and (permissionId)
    '''

    def delete_permission(self,fileId,permissionId):
        # Delete the permission for given (fileID) and (permissionId)
        result = self.drive_service.permissions().delete(fileId=fileId, 
                                                         permissionId=permissionId).execute()
        print('Permission ID [',permissionId,'] had been deleted!')
        return result

    ########################################################################################################
    ########################################################################################################
    
    '''
    This function will list all the permissions for a certian file ID
    '''

    def list_file_permission(self,fileId,*positional_parameters, **keyword_parameters):
        # Get the list of all permissions for given fileId
        perms = self.drive_service.permissions().list(fileId=fileId).execute()['permissions']

        # Add the (emailAdress) to permissions dictionary
        for perm in perms:
            user_email = self.get_user_email_permission(fileId,perm['id'])
            perm['user_email'] = user_email

        # If (printList) parameter have been passed, then the function will print the dictionary
        if ('printList' in keyword_parameters):
            if(keyword_parameters['printList'] == True):
                for perm in perms:
                    if(perm['user_email'] != -1):
                        print("%-25s  ==> %-10s: [%s]"%(perm['id'],perm['role'],perm['user_email']))

        return perms
    
    ########################################################################################################
    ########################################################################################################
    
    '''
    This function will return user email for a given permission ID and file ID
    '''

    def get_user_email_permission(self,fileId,permissionId):
        # Get the email address of user for given (fileID) and (permissionId)
        try:
            return self.drive_service.permissions().get(fileId=fileId, 
                                                         permissionId=permissionId,
                                                         fields='emailAddress').execute()['emailAddress']
        except errors.HttpError:
            return -1
        return None


##############################################################################################################
##############################################################################################################
########################################## Google Sheet APIs #################################################
##############################################################################################################
##############################################################################################################

'''
Create google_sheet class:
The object should take (SERVICE_SHEET) object
'''

class google_sheet:
    
    # Create constructor
    def __init__(self,SHEET_SERVICE,SHEET_CLIENT,DRIVE):
        self.sheet_service = SHEET_SERVICE
        self.sheet_client = SHEET_CLIENT
        self.drive = DRIVE
        
    ########################################################################################################
    ########################################################################################################

    '''
    This function will read a specific sheet in a Google spreadsheet and convert it to a DataFrame
    Hence, the (spreadsheetID OR spreadsheetName) and (sheetName) are both required.
    '''

    def read_sheet(self, *positional_parameters, **keyword_parameters):
        if (
            ('spreadsheetName' in keyword_parameters) & 
            ('spreadsheetId' in keyword_parameters) 
        ):
            raise Exception('Please use either (spreadsheetName) or (spreadsheetId) EXECLUSIVILY')
            return None
        elif (
            (not('spreadsheetName' in keyword_parameters)) & 
            (not('spreadsheetId' in keyword_parameters)) 
        ):
            raise Exception('Please enter a file indicator: Either (spreadsheetName) or (spreadsheetId) EXECLUSIVILY')
            return None
        elif (
            (not('sheetName' in keyword_parameters))
        ):
            raise Exception('Please enter a sheet name')
            return None
        elif (
            (not('spreadsheetId' in keyword_parameters)) &
            ('spreadsheetName' in keyword_parameters)
        ):
            if(keyword_parameters['spreadsheetName']):
                # Retrieve metadata of spreadsheet
                files_list = self.drive.list_files()
                files_list = dict(zip(files_list.values(),files_list.keys()))
                spreadsheetId = files_list[keyword_parameters['spreadsheetName']]
        elif (
            ('spreadsheetId' in keyword_parameters) &
            (not('spreadsheetName' in keyword_parameters))
        ):
            if(keyword_parameters['spreadsheetId']):
                spreadsheetId  = keyword_parameters['spreadsheetId']
        
        
        # Use client to ge the spreadsheet and sheet
        result = self.sheet_client.values().get(spreadsheetId=spreadsheetId,
                                        range=keyword_parameters['sheetName']).execute()

        # Extract the sheet values and convert it to pandas DataFrame
        values = result.get('values', [])
        headers = values.pop(0)
        return pd.DataFrame(values, columns=headers)

    ########################################################################################################
    ########################################################################################################

    '''
    This function returns the sheets names and sheets ID of given spreadsheets ID
    set (printList = True), if you want to print the sheets names and IDS
    '''

    def list_sheets(self, *positional_parameters, **keyword_parameters):

        if (
            ('spreadsheetName' in keyword_parameters) & 
            ('spreadsheetId' in keyword_parameters) 
        ):
            raise Exception('Please use either (spreadsheetName) or (spreadsheetId) EXECLUSIVILY')
            return None
        elif (
            (not('spreadsheetName' in keyword_parameters)) & 
            (not('spreadsheetId' in keyword_parameters)) 
        ):
            raise Exception('Please enter a file indicator: Either (spreadsheetName) or (spreadsheetId) EXECLUSIVILY')
            return None
        elif (
            (not('spreadsheetId' in keyword_parameters)) &
            ('spreadsheetName' in keyword_parameters)
        ):
            if(keyword_parameters['spreadsheetName']):
                # Retrieve metadata of spreadsheet
                files_list = self.drive.list_files()
                files_list = dict(zip(files_list.values(),files_list.keys()))
                spreadsheetId = files_list[keyword_parameters['spreadsheetName']]
        elif (
            ('spreadsheetId' in keyword_parameters) &
            (not('spreadsheetName' in keyword_parameters))
        ):
            if(keyword_parameters['spreadsheetId']):
                spreadsheetId  = keyword_parameters['spreadsheetId']

        
        # Retrieve metadata of spreadsheet
        sheet_metadata = self.sheet_service.spreadsheets().get(spreadsheetId=spreadsheetId).execute()

        # Extract sheets and put them into a dictionary
        sheets = sheet_metadata.get('sheets', '')
        sheets_ids = []
        sheets_name = []
        for i in range(len(sheets)):
            title = sheets[i].get("properties", {}).get("title", '')
            sheets_name.append(title)
            sheet_id = sheets[i].get("properties", {}).get("sheetId", '')
            sheets_ids.append(sheet_id)
        sheets = dict(zip(sheets_ids, sheets_name))

        # If (printList) parameter have been passed, then the function will print the dictionary
        if ('printList' in keyword_parameters):
            if(keyword_parameters['printList'] == True):
                i = 1
                for sheet in sorted(sheets):
                    print("%i. %-20s  ==> \t%-20s"%(i,sheets[sheet], str(sheet)))
                    i = i+1

        return sheets

    ########################################################################################################
    ########################################################################################################

    def clear_sheet(self, *positional_parameters, **keyword_parameters):
        
        if (
            ('spreadsheetName' in keyword_parameters) & 
            ('spreadsheetId' in keyword_parameters) 
        ):
            raise Exception('Please use either (spreadsheetName) or (spreadsheetId) EXECLUSIVILY')
            return None
        elif (
            (not('spreadsheetName' in keyword_parameters)) & 
            (not('spreadsheetId' in keyword_parameters)) 
        ):
            raise Exception('Please enter a file indicator: Either (spreadsheetName) or (spreadsheetId) EXECLUSIVILY')
            return None
        elif (
            (not('sheetName' in keyword_parameters))
        ):
            raise Exception('Please enter a sheet name')
            return None
        elif (
            (not('spreadsheetId' in keyword_parameters)) &
            ('spreadsheetName' in keyword_parameters) &
            ('sheetName' in keyword_parameters)
        ):
            if(keyword_parameters['spreadsheetName']):
                # Retrieve metadata of spreadsheet
                files_list = self.drive.list_files()
                files_list = dict(zip(files_list.values(),files_list.keys()))
                spreadsheetId = files_list[keyword_parameters['spreadsheetName']]
                sheetName = keyword_parameters['sheetName']
        elif (
            ('spreadsheetId' in keyword_parameters) &
            (not('spreadsheetName' in keyword_parameters)) &
            ('sheetName' in keyword_parameters)
        ):
            if(keyword_parameters['spreadsheetId']):
                spreadsheetId  = keyword_parameters['spreadsheetId']
                sheetName = keyword_parameters['sheetName']
        
        # Specify all the range
        rangeAll = '{0}!A1:ZZZ'.format(sheetName)
        # Set the sheet body to be empty
        body = {}
        # Apply clarification setting to sheet
        resultClear = self.sheet_service.spreadsheets().values().clear(spreadsheetId=spreadsheetId,
                                                                       range=rangeAll,
                                                                       body=body).execute()

        print('Done clearing sheet!')
        return resultClear
    
    ########################################################################################################
    ########################################################################################################

    '''
    This function will write DataFrame to a specific sheet in a Google spreadsheet
    Hence, the (spreadsheetID OR spreadsheetName) and (sheetName) are both required.
    '''
    
    def write_sheet(self,df,*positional_parameters, **keyword_parameters):
        
        if (
            ('spreadsheetName' in keyword_parameters) & 
            ('spreadsheetId' in keyword_parameters) 
        ):
            raise Exception('Please use either (spreadsheetName) or (spreadsheetId) EXECLUSIVILY')
            return None
        elif (
            (not('spreadsheetName' in keyword_parameters)) & 
            (not('spreadsheetId' in keyword_parameters)) 
        ):
            raise Exception('Please enter a file indicator: Either (spreadsheetName) or (spreadsheetId) EXECLUSIVILY')
            return None
        elif (
            (not('sheetName' in keyword_parameters))
        ):
            raise Exception('Please enter a sheet name')
            return None
        elif (
            (not('spreadsheetId' in keyword_parameters)) &
            ('spreadsheetName' in keyword_parameters) &
            ('sheetName' in keyword_parameters)
        ):
            if(keyword_parameters['spreadsheetName']):
                # Retrieve metadata of spreadsheet
                files_list = self.drive.list_files()
                files_list = dict(zip(files_list.values(),files_list.keys()))
                spreadsheetId = files_list[keyword_parameters['spreadsheetName']]
                sheetName = keyword_parameters['sheetName']
        elif (
            ('spreadsheetId' in keyword_parameters) &
            (not('spreadsheetName' in keyword_parameters)) &
            ('sheetName' in keyword_parameters)
        ):
            if(keyword_parameters['spreadsheetId']):
                spreadsheetId  = keyword_parameters['spreadsheetId']
                sheetName = keyword_parameters['sheetName']
                
        # Make sure to clean the sheet first
        self.clear_sheet(self.sheet_service,spreadsheetId=spreadsheetId,sheetName=sheetName)
        # Convert the DataFrame into list of list for uploading
        values = df.astype(str).values.tolist()
        # Make sure to include the headers
        headers = df.columns.values.tolist()
        values.insert(0,headers)
        Body = {
                'values':values,
                'majorDimension':'ROWS',
               }
        result = self.sheet_service.spreadsheets().values().update(spreadsheetId=spreadsheetId, 
                                                                   range=sheetName, 
                                                                   valueInputOption='USER_ENTERED', 
                                                                   body=Body).execute()

        # Print output of API if specified
        if ('printResult' in keyword_parameters):
            if(keyword_parameters['printResult']):
                files_list = self.drive.list_files()
                print('Done Writing to spreadsheet (',files_list[spreadsheetId], ') with ID ==> [',spreadsheetId,']')
                print('Writing range is: ',result['updatedRange'])
                print('Number of updated rows is: ',result['updatedRows'])
                print('Number of updated columns is: ',result['updatedColumns'])
                print('Number of updated cells is: ',result['updatedCells'])
                
        return result

    ########################################################################################################
    ########################################################################################################

    '''
    This function will add an empty sheet in a Google spreadsheet
    Hence, the (spreadsheetID OR spreadsheetName) and (sheetName) of new sheet are both required.
    '''
    
    def add_sheet(self,*positional_parameters, **keyword_parameters):
        
        if (
            ('spreadsheetName' in keyword_parameters) & 
            ('spreadsheetId' in keyword_parameters) 
        ):
            raise Exception('Please use either (spreadsheetName) or (spreadsheetId) EXECLUSIVILY')
            return None
        elif (
            (not('spreadsheetName' in keyword_parameters)) & 
            (not('spreadsheetId' in keyword_parameters)) 
        ):
            raise Exception('Please enter a file indicator: Either (spreadsheetName) or (spreadsheetId) EXECLUSIVILY')
            return None
        elif (
            (not('sheetName' in keyword_parameters))
        ):
            raise Exception('Please enter a sheet name')
            return None
        elif (
            (not('spreadsheetId' in keyword_parameters)) &
            ('spreadsheetName' in keyword_parameters) &
            ('sheetName' in keyword_parameters)
        ):
            if(keyword_parameters['spreadsheetName']):
                # Retrieve metadata of spreadsheet
                files_list = self.drive.list_files()
                files_list = dict(zip(files_list.values(),files_list.keys()))
                spreadsheetId = files_list[keyword_parameters['spreadsheetName']]
                sheetName = keyword_parameters['sheetName']
        elif (
            ('spreadsheetId' in keyword_parameters) &
            (not('spreadsheetName' in keyword_parameters)) &
            ('sheetName' in keyword_parameters)
        ):
            if(keyword_parameters['spreadsheetId']):
                spreadsheetId  = keyword_parameters['spreadsheetId']
                sheetName = keyword_parameters['sheetName']
            
        # Specify the request body of API with (addSheet) option
        batch_update_spreadsheet_request_body = {
            'requests':[
                {
                    "addSheet": 
                     {
                         "properties": 
                         {
                             "title": sheetName,
                         }
                     }
                }
            ],
        }
        request = self.sheet_service.spreadsheets().batchUpdate(spreadsheetId=spreadsheetId, 
                                                                body=batch_update_spreadsheet_request_body
                                                               ).execute()
        
        sheetId = request['replies'][0]['addSheet']['properties']['sheetId']
        
        # Print output of API if specified
        if ('printResult' in keyword_parameters):
            if(keyword_parameters['printResult']):
                files_list = self.drive.list_files()
                print('New sheet named (', sheetName ,') have been added to spreadsheet (',
                      files_list[spreadsheetId], ')')
                print('New sheet ID is: ', sheetId)
                
        return sheetId

    ########################################################################################################
    ########################################################################################################

    '''
    This function will delete a sheet in a Google spreadsheet
    Hence, the (spreadsheetID OR spreadsheetName) and (sheetName)  are both required.
    '''
    
    def delete_sheet(self,*positional_parameters, **keyword_parameters):
        
        if (
            ('spreadsheetName' in keyword_parameters) & 
            ('spreadsheetId' in keyword_parameters) 
        ):
            raise Exception('Please use either (spreadsheetName) or (spreadsheetId) EXECLUSIVILY')
            return None
        elif (
            (not('spreadsheetName' in keyword_parameters)) & 
            (not('spreadsheetId' in keyword_parameters)) 
        ):
            raise Exception('Please enter a file indicator: Either (spreadsheetName) or (spreadsheetId) EXECLUSIVILY')
            return None
        elif (
            (not('sheetName' in keyword_parameters))
        ):
            raise Exception('Please enter a sheet name')
            return None
        elif (
            (not('spreadsheetId' in keyword_parameters)) &
            ('spreadsheetName' in keyword_parameters) &
            ('sheetName' in keyword_parameters)
        ):
            if(keyword_parameters['spreadsheetName']):
                # Retrieve metadata of spreadsheet
                files_list = self.drive.list_files()
                files_list = dict(zip(files_list.values(),files_list.keys()))
                spreadsheetId = files_list[keyword_parameters['spreadsheetName']]
                sheetName = keyword_parameters['sheetName']
        elif (
            ('spreadsheetId' in keyword_parameters) &
            (not('spreadsheetName' in keyword_parameters)) &
            ('sheetName' in keyword_parameters)
        ):
            if(keyword_parameters['spreadsheetId']):
                spreadsheetId  = keyword_parameters['spreadsheetId']
                sheetName = keyword_parameters['sheetName']
        
        # Extract (sheetID) from (SheetName)
        sheets = self.list_sheets(spreadsheetId=spreadsheetId)
        sheets = dict(zip(sheets.values(),sheets.keys()))
        sheetId = sheets[sheetName]
        
        # Specify the request body of API with (deleteSheet) option
        batch_update_spreadsheet_request_body = {
            'requests':[
                {
                    "deleteSheet": 
                    {
                         "sheetId": sheetId
                    }
                }
            ],
        }
        request = self.sheet_service.spreadsheets().batchUpdate(spreadsheetId=spreadsheetId,
                                                                body=batch_update_spreadsheet_request_body
                                                               ).execute()
        print('Done deleting (',sheetName,') sheet!')
        return request

    ########################################################################################################
    ########################################################################################################
