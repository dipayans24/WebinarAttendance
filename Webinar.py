from datetime import date, datetime, time, timedelta
from io import BytesIO
import glob
import os
import re
import tempfile
import warnings
from math import floor
import gdown
import gspread
import pandas as pd
import streamlit as st
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload, MediaIoBaseDownload
from google.oauth2 import service_account
from gspread_dataframe import get_as_dataframe, set_with_dataframe
from oauth2client.service_account import ServiceAccountCredentials
from stqdm import stqdm

warnings.filterwarnings("ignore")
warnings.filterwarnings('ignore')


SCOPES = ['https://www.googleapis.com/auth/drive']

def authenticate(SERVICE_ACCOUNT_FILE, SCOPES):
    creds = service_account.Credentials.from_service_account_file(SERVICE_ACCOUNT_FILE, scopes=SCOPES)
    return creds

def getFilesList(PARENT_FOLDER_ID, service):
  results = service.files().list(
      q=f"'{PARENT_FOLDER_ID}' in parents and trashed=false",
      fields="nextPageToken, files(id, name)"
  ).execute()
  file_list = results.get('files', [])

  return file_list



def downloadFiles(ProcessBatch, MainDataFileDate, clearPreviousData, GdriveCredentials):
 

  SCOPES = ['https://www.googleapis.com/auth/drive']

  SERVICE_ACCOUNT_FILE = GdriveCredentials
  creds = authenticate(SERVICE_ACCOUNT_FILE, SCOPES)
  service = build('drive', 'v3', credentials=creds)

  st.write("Downloading the attendee reports....")
  filePaths, service  = getFilefromGdrive('1hmBMW_SkVUQeckKfFOjZdnBXJKDFTZUN', service, ProcessBatch, clearPreviousData)

  #if len(glob.glob("Main*.xlsx")) == 0:
  st.write("Downloading the Main File....")
  MainFilePath, service  = getFilefromGdrive('13-pSnmkVenS-g693qKd_Txvzq489Zsq4', service, MainDataFileDate, clearPreviousData)
#   else:
#     MainFilePath  = glob.glob("Main*.xlsx")
   
  return filePaths, MainFilePath

# @title Generates formatted file from the attendance report in an excel file
def getNumber(number):
    try:
        mobile =  re.findall(r'\d+', str(number))[0]
        if len(mobile) == 11 and mobile[0] == "0":
            return "91"+str(mobile[1:])
        elif len(mobile)==10:
            return "91"+str(mobile)
            #return str(mobile)
        else:
            return str(mobile)
    except:
        return pd.NA

def getCleanPhone(number):
    try:
      return  re.findall(r'\d+', str(number))[0]
    except:
      return pd.NA

def CountIf(Main_File, Current_File, MFCol, CFCol, filename, ValueCol=None):  # CHANGED: added optional ValueCol parameter (e.g. "SessionDuration")

    MainFileColCleaned = MFCol.replace(" ", "") #Replaces any space in column names
    CurrentFileColCleaned = CFCol.replace(" ", "") #Replaces any space in column names

    NewColName = MainFileColCleaned[:5]+"_"+re.sub(r"\W", "", filename[:5])+"_"+CurrentFileColCleaned #Creates a new column name to store output

    if ValueCol is not None:  # NEW: only when a value column is requested
        NewColName = NewColName+"_"+ValueCol  # NEW: add the value column name so it doesn't clash with the 0/1 flag column

    newcol = 1
    while NewColName in Main_File.columns:  #Checks whether the newly generated column name already exists or not
        NewColName = NewColName+"_"+str(newcol)
        newcol = newcol+1

    # ✅ OPTIMIZED: Build a lookup set for O(1) membership check instead of O(n) row-by-row scan
    lookup_set = set(Current_File[CFCol].astype(str).str.lower().str.strip())

    MainKeys = Main_File[MFCol].astype(str).str.lower().str.strip()  # NEW: cleaned email/phone keys from the main file

    if ValueCol is None:  # NEW: default behaviour -> 0/1 presence flag, same as before
        value = MainKeys.isin(lookup_set).astype(int)  # NEW: 1 if the key exists in the attendee file, else 0
    else:  # NEW: value mode -> return the max of ValueCol per key
        CurrentKeys = Current_File[CFCol].astype(str).str.lower().str.strip()  # NEW: cleaned keys from the attendee file
        MaxLookup = Current_File[ValueCol].groupby(CurrentKeys).max()  # NEW: max value per key (handles people who joined more than once)
        value = MainKeys.map(MaxLookup).fillna(0)  # NEW: look up each main-file key's max value, 0 if not found

    Main_File.insert(loc=len(Main_File.columns), column=NewColName,
                     value=value,  # CHANGED: uses the value computed above (flag or max)
                     allow_duplicates=True)

    return Main_File, NewColName

def getAttendanceFormat(filePath, AttendanceTimeThreshold = 0):
  
  fileName = filePath.split(".")[0].replace("attendee_", "")
  MaxSessionDuration = None
  try:
    Sessiondetails = pd.read_csv(filePath, sep=",", index_col=False, skiprows=2, nrows = 1)
    MaxSessionDuration = Sessiondetails["Actual Duration (minutes)"][0]
  except:
    st.write(filePath)

  #While loop for reading the file untill the parser error is not encountered.
  cnt = 0
  while cnt != None:
      try:
          attendance = pd.read_csv(filePath, sep="," , index_col=False ,  skiprows=cnt )
          
          cnt = None
      except:
          cnt = cnt+1


  #st.write(filePath)
  #drops all the join time rows which are blank.
  attendance = attendance[(attendance["Join Time"] != "--")].dropna(subset=["Join Time"])

  #drops all the irrelevant rows
  attendance = attendance.where(lambda x : x["Attended"] == "Yes", other=pd.NA).dropna(subset=["Attended"])

  #Change the datatype of Join Time and Leave Time
  attendance[["Join Time", "Leave Time"]] = attendance.loc[:, ["Join Time", "Leave Time"]].map(lambda x : pd.to_datetime(x, format="%m/%d/%Y %I:%M:%S %p"))

  attendanceDate = attendance["Join Time"].dt.date.unique()[0]

  fileNamePart = attendance["Join Time"].dt.strftime("%d%b%Y").unique()[0]

  #Extracts only the digits from the Phone column
  attendance["FullNumber"] = attendance.Phone.map(lambda x: getCleanPhone(x))

  attendance["Phone"] = attendance["Phone"].map(lambda x: getNumber(x))

  #Change the datatype to float for the Time in session column
  attendance["Time in Session (minutes)"] = attendance["Time in Session (minutes)"].fillna(0).astype(float, errors="ignore")

  attendance["WebinarID"] = fileName.replace(r"\W", "")
  #Group by using Email to get the total duration for that user
  Duration  = attendance[["Email", "Time in Session (minutes)"]].groupby(by="Email", as_index=False).agg( SessionDuration = ("Time in Session (minutes)", "sum") )

  #Drop the duplicate rows using the Email column
  cleanAttendance = attendance.drop_duplicates(subset="Email", keep="first").loc[:, ["Email", "Phone", "FullNumber", "WebinarID"]]

  #Merges the both the dataframe
  final = cleanAttendance.merge(Duration, how="left", left_on="Email", right_on="Email")

  #Change the datatype to float for the Phone column
  final.Phone = final.Phone.fillna(0).astype(float, errors="ignore")

  #Filters the data by the value mentioned in the AttendanceTimeThreshold threshold
  final = final[final["SessionDuration"] > AttendanceTimeThreshold]

  if MaxSessionDuration is not None:
    final["SessionDuration"] = final["SessionDuration"].map(lambda x : float(x) if x<= MaxSessionDuration else MaxSessionDuration)

  #Sorts the dataframe using the Time in session(minutes)
  final.sort_values(by="SessionDuration", ascending=False, inplace=True)

  final["Date"] = None
  final["Date"] = final["Date"].apply(lambda x: attendanceDate if x is None else x)

  localFileSaveName = f"{fileNamePart}_{fileName}_Formatted.csv"
  outputFileName = f"attendee_{fileName}_Formatted.xlsx"

  #final.to_csv(localFileSaveName, sep=",", index=False)

  #Saves the file to an excel file.
  with pd.ExcelWriter(outputFileName, engine="openpyxl", mode="w") as f:
      final.to_excel(f, sheet_name="EligibleAttendee", index=False)
      #attendance.to_excel(f, sheet_name="OriginalSheet", index=False)

  #files.download(outputFileName) #Downloads the excel sheets

  #st.write(f"File downloaded {outputFileName}.")

  return outputFileName

def update_dict(d, key, value):
    if key in d:
        d[key].add(value)  # Add new value to existing set
    else:
        d[key] = {value}   # Initialize with a new set containing the value

    return d


def getWebinarSheet(ProcessBatch, credential_Upload):
    #output_filename = "WebinarDetails"
  output_filename = "WebinarDetails.xlsx"
    #   # Remove existing file if it exists to avoid conflicts

    #     # sheet_id = "1wUviIGWnfOeTTYW8dlnIspAi2G91mgMiP607i6PGncE"
    #     # url = f"https://drive.google.com/uc?id={sheet_id}"
    #     # gdown.download(url, output_filename, quiet=True)

    #   scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    #   creds = ServiceAccountCredentials.from_json_keyfile_name(credential_Upload, scope)
    #   client = gspread.authorize(creds)

    #   sheet_id = "1wUviIGWnfOeTTYW8dlnIspAi2G91mgMiP607i6PGncE"
    #   workbook = client.open_by_key(sheet_id)
    #   values = workbook.worksheet("WebinarDetails").get_all_values()
    #   records = workbook.worksheet("WebinarDetails").get_all_records()
    # sheet_name = datetime.now().strftime("%b-%Y")

  # Read the downloaded XLSX file into a pandas DataFrame
  try:
      WebinarDetails, records, workbook = getSheet("1wUviIGWnfOeTTYW8dlnIspAi2G91mgMiP607i6PGncE", "WebinarDetails", credential_Upload)
      WebinarDetails = WebinarDetails[WebinarDetails["Cancelled"] != "Yes"]
      # WebinarDetails = pd.read_excel(output_filename, parse_dates = True)
      # WebinarDetails = WebinarDetails[WebinarDetails["Cancelled"].isna()]
      WebinarDetails["WebinarID"] = WebinarDetails["WebinarID"].str.replace(r"\W", "", regex=True) #removing any empty space.
      WebinarDetails.drop_duplicates(subset=WebinarDetails.columns, inplace=True)
  except Exception as e:
      st.error(f"\nError reading '{output_filename}' with pandas: {e}")


  BatchesWebinar = {}


  MultiWebinar = WebinarDetails.groupby(by=["Date",	"BatchName"], as_index=False).agg(CountWebinar = ("WebinarID", "count"))
  MultiWebinar = MultiWebinar[MultiWebinar["CountWebinar"]>1]
  WebinarList = MultiWebinar.merge(WebinarDetails, on=["Date", "BatchName"])

  WebinarDict  = WebinarList.loc[WebinarList["BatchName"] == ProcessBatch].groupby(by="Date")["WebinarID"].apply(list).to_dict()

  WebinarDict = {
      datetime.strptime(k, '%d-%m-%Y').strftime('%Y_%m_%d'): v
      for k, v in WebinarDict.items()
  }

  WebinarList = WebinarList.loc[WebinarList["BatchName"] == ProcessBatch]["WebinarID"].to_list()

  concatRequired = {}

  #Checks for the file which starts with "attendee_" and is an excel file and it checks for the batches for those files.
  for files in os.listdir(os.path.curdir):
      if files.startswith("attendee_") & files.endswith("xlsx"):
          WebinarID = files.replace("attendee_", "").split("_")[0]
          BatchNames = WebinarDetails[WebinarDetails["WebinarID"] == WebinarID]["BatchName"]

          date_ = "_".join(files.replace("attendee_", "").split("_")[1:-1])
          if WebinarID in WebinarList:
              update_dict(concatRequired, date_, files)

          if len(BatchNames) == 1:
              update_dict(BatchesWebinar, BatchNames.to_string(index=False), files)

          elif len(BatchNames) > 1:
              for BatchName in BatchNames:
                  update_dict(BatchesWebinar, BatchName, files)
  #st.write(concatRequired)
  return BatchesWebinar, concatRequired, WebinarList,WebinarDict, records, WebinarDetails, workbook

#@title Updates the WebinarDetails Sheet with the duration of the session and ActualStartTime

#Converts the date to specific syntax
def getdatetime(date):
  try:
    try:
      return datetime.strptime(date, "%m/%d/%Y %I:%M:%S %p")
    except:
      return datetime.strptime(date, "%m-%d-%y %H:%M:%S")
  except:
    return pd.NaT

def setWebinardetails(filePaths, records, WebinarDetails, workbook, ProcessBatch):
  Sessiondetails = pd.DataFrame() #Creates an empty dataframe

  #Iterates over the filepaths to get the required details
  for file in filePaths:
    try:
      attendance = pd.read_csv(file, sep=",", index_col=False, skiprows=2, nrows = 1)
      Sessiondetails = pd.concat([Sessiondetails, attendance], axis = 0)
    except:
      st.write(file)

  #Modifies some changes to the Sessiondetails dataframe
  Sessiondetails = Sessiondetails.drop_duplicates(subset=Sessiondetails.columns)
  Sessiondetails["Webinar ID"] = Sessiondetails["Webinar ID"].astype(str).str.replace(" ", "")
  Sessiondetails.rename(columns = {"Webinar ID":"WebinarID", "Actual Start Time":"Start Time"}, inplace = True)

  #Merges the dataframe and filters the data
  UpdatedDetails = WebinarDetails.merge(Sessiondetails.loc[:, ["Topic", 	"WebinarID", 	"Start Time", 	"Actual Duration (minutes)"]], on = "WebinarID", how = "left")
  SessionCondition = ((UpdatedDetails["BatchName"] == ProcessBatch) & (~UpdatedDetails["Start Time"].isna()))
  UpdatedDetails.loc[SessionCondition, "Start Time"] = UpdatedDetails.loc[SessionCondition, "Start Time"].apply(lambda x: getdatetime(x))

  #Creates a new dataframe with only the required details
  NewDetails = UpdatedDetails.loc[((SessionCondition) & (UpdatedDetails["Duration"] == "")), ["Date", "WebinarID", "Actual Duration (minutes)", "Start Time" , "BatchName"]]

  #Iterates with the NewDetails dataframe
  for newDate in NewDetails.itertuples():
    for i, row in  enumerate(records):

      #Gets the target date
      TARGET_DATE = newDate[1]
      if str(row["Date"]) == TARGET_DATE and row["Duration"] == "" and row["BatchName"] == ProcessBatch and \
            row["WebinarID"].replace(" ", "") == newDate[2]:
        sheet_row = i + 2  # +2 accounts for 0-index + header row

        # Find the column index of "Duration"
        headers = workbook.worksheet("WebinarDetails").row_values(1)

        #Gets the session duration
        NEW_DURATION = newDate[3]
        duration_col = headers.index("Duration") + 1  # gspread cols are 1-indexed
        # Update the specific cell
        workbook.worksheet("WebinarDetails").update_cell(sheet_row, duration_col, NEW_DURATION)

        #Gets the session actual start date
        ActualStartTime = headers.index("Actual Start Time") + 1
        START_DATE = newDate[4]
        workbook.worksheet("WebinarDetails").update_cell(sheet_row, ActualStartTime, str(START_DATE))

        #Gets the BatchType
        BatchType = headers.index("BatchType") + 1
        BATCH_TYPE = "Evening" if START_DATE.time() > time(14, 0, 0) else "Morning"
        workbook.worksheet("WebinarDetails").update_cell(sheet_row, BatchType, str(BATCH_TYPE))

# @title -----OPTIONAL CASE----- When a 2 or more webinars has been assigned to a single batch. Else DON'T RUN.

def checkOptionalCase(BatchesWebinar, concatRequired, WebinarDict, WebinarList, GeneratedFiles, ProcessBatch):

  updatedBatchesWebinar = {}
  dropFiles = []

  if len(concatRequired)>0:
    for dates in WebinarDict.keys():

      concat = pd.DataFrame()
      found = None
      newFileName =  "-".join(WebinarDict[dates])


      for batchFile in BatchesWebinar.values():
          for files in batchFile:
              found = any(files in val_set for val_set in concatRequired[dates])
              if found:
                  fileLoc = os.path.curdir
                  dropFiles.append(files)
                  df = pd.read_excel(rf"{fileLoc}/{files}")
                  concat = pd.concat([df, concat], axis=0)


      date = concat.head(1).Date.astype("M8[ns]").dt.strftime("%Y_%m_%d").to_string(index=False)
      newFileName = f"attendee_{newFileName}_{date}.xlsx"

      concat.to_excel(f"{newFileName}", index=False)

      for b in BatchesWebinar.keys():
          for f in BatchesWebinar.get(b):
              found = any(f in val_set for val_set in list(concatRequired[dates]))
              if found:
                  update_dict(updatedBatchesWebinar, b, newFileName)

              else:
                  update_dict(updatedBatchesWebinar, b, f)

  NewBatchesWebinar = {}

  for b in updatedBatchesWebinar.keys():
    for f in updatedBatchesWebinar.get(b):
      if f not in dropFiles:
        update_dict(NewBatchesWebinar, b, f)

  if len(updatedBatchesWebinar) > 0:
    BatchesWebinar = NewBatchesWebinar

  st.write(f"Files need to concatenated {len(BatchesWebinar.get(ProcessBatch))}")

#   for i in GeneratedFiles:
#     if i not in BatchesWebinar.get(ProcessBatch):
#         st.write(i)

  return BatchesWebinar

# @title Filters the main file according to the attendee file uploaded and generates an excel file for every batch

def generateData(MainFilePath, BatchesWebinar, ProcessBatch, BacthesWO_W):

  Main_File = pd.read_excel(MainFilePath) #Reads the filtered main data
  Final_Generated_File = [] #Stores the names of the generated file.
  final_df = Main_File[(~Main_File["High Ticket Activity"].astype(str).str.strip().str.lower().str.contains("pause"))\
                       &(Main_File["Status"].isna() )].reset_index(drop=True)

  if len(BatchesWebinar)>0:
    #for batches in BatchesWebinar.keys():
    for batches in [ProcessBatch]:

      if batches in BacthesWO_W:
        batchName = batches
        filterCondition = final_df["batch name"].fillna("empty").str.strip().isin([batchName])

      else:
        batchName = batches+"W"
        filterCondition = final_df["batch name"].fillna("empty").str.strip().str.contains(batchName)

      #Filters the columns if Main Data is uploaded ,else if existing attendance file is uploaded, it skips the filtering.

      final_df = final_df[filterCondition].reset_index(drop=True)

      st.write(f"Total Count for {batches} is - {len(final_df)}")
      #'Refunded'
      removeCol = ['Schedule',  'sat link', 'sun link', 'LMS link of recording','Telegram', 'High Ticket Activity', 'update', 'Reason']
      final_df.drop(columns=removeCol, inplace=True, errors="ignore")

      final_df.dropna(subset=["batch name"], inplace=True, how="all") #Drops any rows which doesn't have any values for the batch name column

      MainFileBatches = final_df["batch name"].unique().tolist() #Gets the unique batches in the main file

      try:  #Gets the batchname by removing splitting from the "W" part.
        batchname = set()
        for i in MainFileBatches:
          batchname.add(i.split("W")[0].replace(" ", ""))

        batchname = ["_".join(batchname)][0]

        MainFileBatches = batchname
      except:
        pass

      resortedCol = ['Registered Number', 'Registered mail', 'CountryCode', 'Whatsaap Number', 'broadcast mail', 'Funnel','Refunded', 'batch name']

      final_df = final_df.reindex(columns=resortedCol)

      indexer = final_df.columns.get_loc("batch name") #Get the column index for the "batch name" column

      MainDataCol = ['Registered Number', 'Registered mail', 'Whatsaap Number', 'broadcast mail'] #Column names required for the main data
      final_df["Registered mail"] = final_df["Registered mail"].astype(str).str.strip() #Converts the data type to string datatype
      final_df["Whatsaap Number"] = final_df["Whatsaap Number"].astype(str).str.strip() #Converts the data type to string datatype
      final_df["Registered Number"] = final_df["Registered Number"].astype(str).str.strip() #Converts the data type to string datatype
      final_df["broadcast mail"] = final_df["broadcast mail"].astype(str).str.strip() #Converts the data type to string datatype

      finalCol = ['Email', 'Phone'] #Column names required for the attendance data

      # Load all webinar files once upfront to avoid repeated disk I/O
      webinar_cache = {f: pd.read_excel(f) for f in BatchesWebinar[batches]}

      SumColNames = []
      for f in stqdm(BatchesWebinar[batches]):

        WebinarId = f.split("_")[1] #Extracts the webinar id from the file name
        Current_File = webinar_cache[f] #Retrieved from cache — no disk read

        #Creating the combination for using the countif function
        MFCombinations = ['Registered mail', 'broadcast mail', 'Registered Number', 'Whatsaap Number', 'Registered Number','Whatsaap Number' ]
        AttendeeCombinations = ['Email',  'Email', 'Phone', 'Phone', 'FullNumber', 'FullNumber']

        Current_File["Phone"] = Current_File["Phone"].astype(str).str.strip() #Converts the data type to string datatype
        Current_File["FullNumber"] = Current_File["FullNumber"].astype(str).str.strip() #Converts the data type to string datatype
        Current_File["Email"] = Current_File["Email"].astype(str).str.strip() #Converts the data type to string datatype

        Current_File["SessionDuration"] = pd.to_numeric(Current_File["SessionDuration"], errors="coerce").fillna(0)  # NEW: makes sure duration is numeric so max() works

        CurrentDateColName =   Current_File["Date"].dt.strftime("%d%b%Y").unique()[0]  #Extracts the date from the attendee report

        # CurrentFileSumColumns = [] #Stores all the columns for the summation
        # for MFCol, CFCol  in zip(MFCombinations, AttendeeCombinations):
        #     final_df, NewColName = CountIf(final_df, Current_File, MFCol, CFCol, f)
        #     SumColNames.append(NewColName)
        #     CurrentFileSumColumns.append(NewColName)


        CurrentFileSumColumns = [] #Stores all the columns for the summation
        CurrentFileDurationColumns = []  # NEW: stores all the session duration columns for this file
        for MFCol, CFCol  in zip(MFCombinations, AttendeeCombinations):
            final_df, NewColName = CountIf(final_df, Current_File, MFCol, CFCol, f)
            SumColNames.append(NewColName)
            CurrentFileSumColumns.append(NewColName)
            final_df, DurColName = CountIf(final_df, Current_File, MFCol, CFCol, f, "SessionDuration")  # NEW: same match, but returns max SessionDuration
            CurrentFileDurationColumns.append(DurColName) 

        #Assigns the webinarId if the countif sum is greater than 0 else "Absent"
        #final_df[CurrentDateColName] = final_df.apply(lambda x: WebinarId if x[CurrentFileSumColumns].sum() >0 else "Absent", axis=1)

        final_df[CurrentDateColName] = final_df.apply(
                lambda x: x[CurrentFileDurationColumns].max()  # CHANGED: returns the max SessionDuration across the email/phone matches instead of WebinarId
                if x[CurrentFileSumColumns].sum() > 0  # same presence check as before: at least one email/phone match
                else "Absent",  # unchanged: no match means "Absent"
                axis=1) 
        #Drops all the column which was required for the countif function
        final_df.drop(columns=CurrentFileSumColumns + CurrentFileDurationColumns, inplace=True)  # NEW: drops the temporary flag and duration columns immediately, so only the date-named column remains

      final_df["Whatsaap Number"] = final_df["Whatsaap Number"].astype(float, errors="ignore")

      final_df.drop(columns="AttendancePercentage", inplace=True, errors="ignore")

      #Generates the Attendance Percentage for each customer
      _att_cols = final_df.iloc[:, indexer+1:]
      _present_counts = (_att_cols != "Absent").sum(axis=1)
      final_df["AttendancePercentage"] = (_present_counts / _att_cols.shape[1] * 100).round(2).astype(str) + "%"

      #Generates the Attendance Percentage for each session
      _session_cols = final_df.columns[indexer+1:]
      final_df[_session_cols] = final_df[_session_cols].astype(object)
      _refunded_mask = final_df["Refunded"].notna() & (final_df["Refunded"] != "")

      _valid_df = final_df[~_refunded_mask]
      _session_pct = (_valid_df[_session_cols] != "Absent").sum(axis=0)
      _session_pct = (_session_pct / len(_valid_df) * 100).round(2).astype(str) + "%"
      final_df.loc[len(final_df), _session_cols] = _session_pct

      final_df.loc[len(final_df)-1, "AttendancePercentage"] = pd.NA

      MainFileOutputName = f"{MainFileBatches}_AttendanceRecord.xlsx"

      col = final_df.columns[indexer+1:-1] #Gets the column after the batch name and before the AttendancePercentage for sorting
      sorted_col = sorted(col, key=lambda x: datetime.strptime(x, "%d%b%Y"), reverse=True) #Sorts the column
      final_df = final_df.reindex(labels=final_df.columns[:indexer+1].to_list()+sorted_col+[final_df.columns[-1]] , axis="columns") #Reorders the columns

      Final_Generated_File.append(MainFileOutputName)

      final_df.to_excel(MainFileOutputName, index=False)

      st.dataframe(final_df.iloc[[-1], 7:12])

  return Final_Generated_File

# @title Uploads data to the google sheet



def updateAttendeeSheet(Final_Generated_File, credential_Upload):
    

  #credential_Upload = upload_file()
  # Authentication
  scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
  creds = ServiceAccountCredentials.from_json_keyfile_name(credential_Upload , scope)
  client = gspread.authorize(creds)

  sheet_id =  "1rbLW59CGLKxlIBM3d7lNopLWpP5oiD3hddWKRzku7Ow"
  workbook = client.open_by_key(sheet_id)


  df = pd.read_excel(Final_Generated_File[0])

  MainFileBatches = df["batch name"].dropna().unique().tolist()

  try:  #Gets the batchname by removing splitting from the "W" part.
      batchname = set()
      for i in MainFileBatches:
        batchname.add(i.split("W")[0].replace(" ", ""))

      batchname = ["_".join(batchname)][0]

      MainFileBatches = batchname
  except:
      pass


  existing_sheet_titles = [ws.title for ws in workbook.worksheets()]

  if MainFileBatches not in existing_sheet_titles:
      AddNewWS = workbook.add_worksheet(title=MainFileBatches, rows='100', cols='20')
      batchData = AddNewWS # Use the newly created worksheet object
  else:

      batchData = workbook.worksheet(MainFileBatches)
      batchData.clear()

  # Ensure set_with_dataframe receives the worksheet object
  set_with_dataframe(batchData, df)

  st.write("Process Completed")
  return True

def save_upload(fileupload, fileType = None):
    temp_dir = tempfile.mkdtemp()
    tmp_path = os.path.join(temp_dir, fileupload.name)

    with open(tmp_path, "wb") as f:
        f.write(fileupload.getvalue())
   
    return tmp_path

def last_friday():
    """
    Returns the date of the most recent Friday in 'YYYY-MM-DD' format.
    If today is already a Friday, returns today's date.
    """
    today = date.today()

    # Monday=0 ... Friday=4
    days_since_friday = (today.weekday() - 4) % 7
    result = today - timedelta(days=days_since_friday)

    return result.strftime("%Y-%m-%d")

def getGdriveService(GdriveCredentials, delegated_user=None):
    # Authenticates with Google Drive using a service account file
    # Pass delegated_user="someone@yourdomain.com" to impersonate a real user (needed if
    # uploading/downloading against a personal My Drive folder rather than a Shared Drive)

    creds = service_account.Credentials.from_service_account_file(GdriveCredentials, scopes=SCOPES)

    if delegated_user:
        creds = creds.with_subject(delegated_user)

    return build('drive', 'v3', credentials=creds)

def getFilesList(parent_folder_id, service):
    # Retrieves ALL files/folders within a parent folder (paginated, Shared-Drive aware)
    file_list = []
    page_token = None
    while True:
        results = service.files().list(
            q=f"'{parent_folder_id}' in parents and trashed=false",
            fields="nextPageToken, files(id, name)",
            pageToken=page_token,
            supportsAllDrives=True,
            includeItemsFromAllDrives=True,
            corpora='allDrives'
        ).execute()
        file_list.extend(results.get('files', []))
        page_token = results.get('nextPageToken')
        if not page_token:
            break
    return file_list

def getSubfolderId(parent_folder_id, folder_name, service):
    # Looks up a named subfolder's ID within a parent folder
    for item in getFilesList(parent_folder_id, service):
        if item['name'] == folder_name:
            return item['id']
    return None

# ---------- Download ----------

def download_file(service, file_id, file_name, clear):
    # Downloads a single file straight to disk, with a Streamlit progress bar
    if file_name not in st.session_state or clear:
        
        request = service.files().get_media(fileId=file_id, supportsAllDrives=True)

        #progress_bar =  st.progress(0, text=f"Downloading {file_name}...")

        with open(file_name, 'wb') as f:
            downloader = MediaIoBaseDownload(fd=f, request=request)
            done = False
            while not done:
                status, done = downloader.next_chunk()
                if status:
                    pct = int(status.progress() * 100)
                    #progress_bar.progress(pct, text=f"Downloading {file_name}... {pct}%")

        #progress_bar.progress(100, text=f"{file_name} downloaded")

        st.session_state[file_name] = file_name
        return file_name

    else:
        return st.session_state[file_name]


def getFilefromGdrive(folder_id, service, ProcessParameter, clear):
    # Downloads all files from a named subfolder within folder_id
    subfolder_id = getSubfolderId(folder_id, ProcessParameter, service)
    file_list = getFilesList(subfolder_id, service)

    filePaths = []
     
    for f in stqdm(file_list):
        fileName = download_file(service, f['id'], f['name'], clear)
        filePaths.append(fileName)

    return filePaths, service


def getSheet(sheet_id, sheet_name, credential_Upload):
  scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
  creds = ServiceAccountCredentials.from_json_keyfile_name(credential_Upload, scope)
  client = gspread.authorize(creds)

  workbook = client.open_by_key(sheet_id)
  values = workbook.worksheet(sheet_name).get_all_values()
  records = workbook.worksheet(sheet_name).get_all_records()
  # sheet_name = datetime.now().strftime("%b-%Y")

  # Read the downloaded XLSX file into a pandas DataFrame
  try:
      paymentSlugs = pd.DataFrame(values[1:], columns=values[0])
      return paymentSlugs, records, workbook
  except:
      return None

def check_session_state(sheet_id  ,sessionVarName , sheet_name , credential_Upload , clear):
        if sessionVarName not in st.session_state or clear:
            st.write(f"Downloading {sessionVarName}.. ")
            st.session_state[sessionVarName], records, workbook = getSheet( sheet_id, sheet_name, credential_Upload)
            return st.session_state[sessionVarName]
        else:
            return  st.session_state[sessionVarName]
    

st.set_page_config("Webinar Attendance", layout="wide")
st.header("📊 Webinar Attendance", divider=True, text_alignment="center")
MainDataFileDate =  str(st.date_input("Select the Last Friday date",value=last_friday()))

col1, col2 = st.columns(2)

with col1:
     credential_Upload = st.file_uploader("Upload Credentials File", type = ["json"]) 
     
with col2:
     GdriveCredentials =  st.file_uploader("Upload GDrive File", type = ["json"]) 
 

col1, col2= st.columns(2)    
with col1:
   clearPreviousData = st.checkbox("Clear Data?", width = "stretch")

if MainDataFileDate and GdriveCredentials and credential_Upload:
    credential_Upload = save_upload(credential_Upload)
    WebinarDetails = check_session_state("1wUviIGWnfOeTTYW8dlnIspAi2G91mgMiP607i6PGncE", "WebinarDetails", "WebinarDetails", credential_Upload, clearPreviousData)
    WebinarDetails = WebinarDetails[WebinarDetails["Cancelled"] != "Yes"]
    WebinarDetails["WebinarID"] = WebinarDetails["WebinarID"].str.replace(r"\W", "", regex=True)
    WebinarDetails.drop_duplicates(subset=WebinarDetails.columns, inplace=True)

    Batches = WebinarDetails["BatchName"].apply(lambda x : "AI CAP B0" + re.sub(r"\D",  "",x) if len(x) == 9 else x).sort_values(ascending=False).unique()

    BatchesList = st.multiselect(label="Select the ProcessBatch", options= Batches, max_selections=5)

    genbtn = st.button("Generate Data", type="primary", on_click=None )


# Example usage:
    if genbtn and BatchesList:
        st.session_state["credential_Upload"] = credential_Upload
        
        BacthesWO_W = [f"AI CAP B{i}" for i in range(10,90)]
        GdriveCredentials = save_upload(GdriveCredentials)
        #st.session_state["GdriveCredentials"] = GdriveCredentials
        for ProcessBatch in BatchesList:
            with st.status(f"Processing for Batch {ProcessBatch}....") as status:
                filePaths, MainFilePath = downloadFiles(ProcessBatch,MainDataFileDate , clearPreviousData, GdriveCredentials)            
                AttendanceTimeThreshold = 0
                GeneratedFiles = [] #Stores the generated file path
                for f in filePaths :

                    filesGen = getAttendanceFormat(f, AttendanceTimeThreshold)
                    
                    GeneratedFiles.append(filesGen)

                st.write("Done Appending Files List")
                BatchesWebinar, concatRequired, WebinarList, WebinarDict, records, WebinarDetails, workbook = getWebinarSheet(ProcessBatch, credential_Upload)
                #setWebinardetails(filePaths, records, WebinarDetails, workbook)

                st.write("Checking for optional case")
                BatchesWebinar = checkOptionalCase(BatchesWebinar, concatRequired,WebinarDict, WebinarList, GeneratedFiles, ProcessBatch)
                Final_Generated_File = generateData(MainFilePath[0], BatchesWebinar, ProcessBatch, BacthesWO_W)

                updateAttendeeSheet(Final_Generated_File, credential_Upload)

                status.update(label=f"🟢 Updated batch {ProcessBatch}.", state="complete")

        # with st.status("Processing..", expanded=True) as status:
        #     service = getGdriveService(GdriveCredentials)  # or getGdriveService(delegated_user="owner@yourdomain.com")
            #filePath, service = getFilefromGdrive('0AHGO663tIOm5Uk9PVA', service, WSDate, clearPreviousData)

        # @title Downloading All Sheets
        # Remove existing file if it exists to avoid conflicts
        
    
        #st.link_button(f"Go to Sheet- {sheet_id}", f"https://docs.google.com/spreadsheets/d/{sheet_id}/edit?usp=sharing", type = "secondary")

             
else:
    pass
    #  with st.status("Links", expanded=False):
    #     col1, col2, col3, col4, col5, col6  = st.columns(6, vertical_alignment = "center",  width="stretch") 
     
    #     with col1:
    #        st.link_button("Open 10xStats", "https://10xstats.com/", width  = "stretch")
    #     with col2:
    #        st.link_button("Open DirectUS", "https://directus-production-62b2.up.railway.app/admin/users/", width  = "stretch") 
    #     with col3:
    #        st.link_button("Open MEGA Exotic", "https://megaexotic.streamlit.app/", width  = "stretch") 
    #     with col4:
    #        st.link_button("Open MEGA AC", "https://megaac.streamlit.app/", width  = "stretch") 
    #     with col5:
    #        st.link_button("Open GdriveUpload", "https://gdriveupload.streamlit.app/", width  = "stretch")
    #     with col6:
    #        st.link_button("Open PaymentSlugsUpdate", "https://paymentslugs.streamlit.app/", width  = "stretch")

