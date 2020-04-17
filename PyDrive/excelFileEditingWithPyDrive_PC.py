from pydrive.auth import GoogleAuth
from pydrive.drive import GoogleDrive
import    pandas     as pd

# Authenticate and create the PyDrive client.
# This only needs to be done once per notebook.


gauth = GoogleAuth()
gauth.LocalWebserverAuth()
drive = GoogleDrive(gauth)


# PUT YOUR FILE ID AND ANY-NAME HERE
file_id   = '15IEFRTSH-9JniOkdZm-Hon9QQayXluYX'  
file_name      = "multiple-sheets-experiment By ANT.xlsx"  # You can use existing drive file name / totally different name

# Get contents of your drive file into the desired file. Here contents are stored in the file specified by 'file_name'
downloaded = drive.CreateFile({'id': file_id})
downloaded.GetContentFile(file_name)

df = pd.read_excel(file_name, usecols=None, sheet_name=None) 
print(df) 
# df['2018'].count() 


sheetNames = ['2018', '2019', '2020']
for sheet in sheetNames:
    df[sheet].drop_duplicates(subset=['Name'], keep="first", inplace=True)
    df[sheet]['Calculated Fine'] = df[sheet]['Absent Days'] * 10

# START [1] ------------- TO UPDATE EXISTING FILE --------------

def updateFileInColab(colabFolder):
    file_list = drive.ListFile({'q': "'%s' in parents and trashed=false" % colabFolder}).GetList()
    for f in file_list:
      if f['title'] == file_name:
        with pd.ExcelWriter('output.xlsx') as writer:
            for sheet in sheetNames:
                df[sheet].to_excel(writer, sheet_name=sheet) 
            writer.save()
            writer.close()
        f.SetContentFile("output.xlsx")
        f.Upload() 
        break


file_list = drive.ListFile({'q': "'root' in parents and trashed=false"}).GetList()
for f in file_list:
  if f['title'] == 'Colab Notebooks':
    print('File:', f)
    updateFileInColab(f['id'])
    break


# END [1]

# START [2] ------------- TO CREATE A NEW FILE --------------


def createNewFile(colabFolder):
    with pd.ExcelWriter('output.xlsx') as writer:
        for sheet in sheetNames:
            df[sheet].to_excel(writer, sheet_name=sheet) 
        writer.save()
        writer.close()

    myFile = drive.CreateFile({'title':'NEW multiple-sheets-experiment.xlsx', 
                                "parents": [{"kind": "drive#fileLink","id": colabFolder}] })  
    myFile.SetContentFile("output.xlsx")
    myFile.Upload()


file_list = drive.ListFile({'q': "'root' in parents and trashed=false"}).GetList()
for f in file_list:
    if f['title'] == 'Colab Notebooks':
        print('File:', f)
        createNewFile(f['id'])
        break
  

# END [2]
