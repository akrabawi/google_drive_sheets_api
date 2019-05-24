# Google Drive & Google Sheets APIs Classes

A python classes to control the Google Drive API and Google Sheets API with direct functions to do the basic documents retrieval, overwrite, delete and permissioning control. The classes are built mainly to manipulate Pandas dataframes with Google Drive and Sheets APIs.

# How To Use

### The following script initiate the working directories and the apis classes

###################################################################

    Work_DIR = os.getenv('Code_Directory', '/home/jovyan/work')       # Current .py code or notebook directory
    Data_DIR = os.path.join(WORK_DIR, 'Data_Directory')               # Subdirectory to store the data
    Credentials_DIR = os.path.join(WORK_DIR, 'Credentials_Directory') # Subdirectory to store the credentials json files

    SCOPES = ['https://www.googleapis.com/auth/drive',
              'https://www.googleapis.com/auth/spreadsheets']

    CLIENT_SECRET_FILE = 'client_service.json'                        # Generated from Google Cloud APIs Console - Service Account

    # Create authuentication
    CLIENT_SECRET_FILE = os.path.join(Credentials_DIR, CLIENT_SECRET_FILE)

    # Create (auth) object
    authInst = auth(SCOPES,CLIENT_SECRET_FILE)

    # Create (client) object
    credentials = authInst.getCredentials()

    # Build Drive Client
    drive_service = discovery.build('drive', 'v3', credentials=credentials)

    # Build Sheets Client
    sheet_service = discovery.build('sheets', 'v4', credentials=credentials)
    sheet_client = sheet_service.spreadsheets()

    # Create the service objects
    drive = google_drive(drive_service)
    sheet = google_sheet(sheet_service,sheet_client,drive)

# Examples

###################################################################

    # Get test dataframe - to be used to test the uploading to Google Drive function
    df = pd.read_csv(Data_DIR + '/some_test_file.csv', sep=',')

###################################################################

    # # Upload the file
    # fileName = 'test'
    # t = drive.create_spreadsheet(df,fileName,overwrite=True)
    # print(t)

###################################################################

    # # Delete the file by name, even if it was multiple files with same name
    # fileName = 'test'
    # drive.delete_file(fileName=fileName)

    # # Delete file by providing the fileID
    # fileId = '1NLrqIFWWyz56Ks7szzAIOjVD52GY4Vdk'
    # drive.delete_file(fileId=fileId)

###################################################################

    # # Give permission to certain user
    # fileId = '1cKNu1l797SQXfyZGXF1Xh4OiuWfSrtZgNpilCPVt994'
    # permission_type = 'user'
    # permission_role = 'writer'
    # user_email = 'ufsdatateam@gmail.com'
    # drive.add_permission(fileId,permission_type,permission_role,user_email)

###################################################################

    # # Delete permission of a file
    # fileId = '1cKNu1l797SQXfyZGXF1Xh4OiuWfSrtZgNpilCPVt994'
    # permissionId = '05351037175981423841i'
    # drive.delete_permission(fileId,permissionId)

###################################################################

    # # Print the list of permissions for a file
    # list = drive.list_file_permission(fileId, printList=True)

###################################################################

    # # Upload the file
    # fileName = 'test.csv'
    # t = drive.create_csv(df,fileName, overwrite=True)

###################################################################

    # # Read the (csv) file by it's name
    # fileName = 'test.csv'
    # df = drive.read_csv(fileName=fileName)
    # print(df.head(1).T)

    # # Read the (csv) file by it's ID
    # fileId = '1H-MCtfxlclucLvYTC4Spg4uJ0g6cZPF4'
    # df = drive.read_csv(fileId=fileId)
    # print(df.head(1).T)

###################################################################

    # # List the files names and IDs
    # files = drive.list_files(printList=True)
    # files = drive.list_files(printList=True, printPermissions=True)
    # # print(files)

###################################################################  

    SPREADSHEET_ID = 'Some_spreadsheet_ID'
    SPREADSHEET_NAME = 'Sheet1'

###################################################################

    # # retrieve the sheets of spreadsheet by providing (spreadsheetId)
    # sheets = sheet.list_sheets(spreadsheetId=SPREADSHEET_ID, printList=True)

    # # retrieve the sheets of spreadsheet by providing (spreadsheetName)
    # sheets = sheet.list_sheets(spreadsheetName=SPREADSHEET_NAME, printList=True)
    # # print(sheets)

###################################################################

    # # Read sheet to pandas by (spreadsheetName) or by (spreadsheetName)
    # SHEET_NAME = 'Sheet1'
    # df = sheet.read_sheet(spreadsheetId=SPREADSHEET_ID,sheetName=SHEET_NAME)
    # # df = sheet.read_sheet(spreadsheetName=SPREADSHEET_NAME,sheetName=SHEET_NAME)
    # print(df.head(1).T)

###################################################################

    # # Clean a sheet in Google spreadsheet
    # sheetName = 'Sheet1'
    # result = sheet.clear_sheet(spreadsheetId=SPREADSHEET_ID,sheetName=sheetName)

###################################################################

    # # Write a DataFrame to existing Google sheet
    # result = sheet.write_sheet(df,spreadsheetId=SPREADSHEET_ID,sheetName=sheetName,printResult=True)

###################################################################

    # # To delete sheet in a spread sheet
    # sheetName = 'Sheet1'
    # result = sheet.delete_sheet(spreadsheetId=SPREADSHEET_ID,sheetName=sheetName)

###################################################################

    # # To add sheet in a spread sheet
    # sheetName = 'Sheet1'
    # result = sheet.add_sheet(spreadsheetId=SPREADSHEET_ID,sheetName=sheetName,printResult=True)
    # # print(result)


* Code written by @sweileho (Omar Sweileh)
