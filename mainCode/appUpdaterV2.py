import glob
import os, shutil, traceback
from mail import send_mail
import numpy as np
from azure.storage.blob import BlobServiceClient
import tkinter as tk





# UPDATE_LOC = r'C:\Users\imam.khan\Documents\EAGS_AppUpdate'
# UPDATE_LOC = r'C:\Users\imam.khan\OneDrive - BioUrja Trading LLC\Documents\EAGS\FinalCodePreRequisites\FinalCodePrep\dist\EAGS_Quote_Generator'

CONNECTION_STRING = 'DefaultEndpointsProtocol=https;AccountName=biourjadayzerstorage01;AccountKey=7IjVelm0PNjGrPiiaTSoC8f3XZMmCQE4Xth+r5STLcHgRoBnAmCCANmkRIL6z5XpBDuRMopiBrEwPDciJI+GPQ==;EndpointSuffix=core.windows.net'

BLOB_CONTAINER = "eags-temp"





"""
    Logic
    Delete files and files folders with _Old in their name

    Connect to Blob service client string using connection_string and sas_key(connect_str = connection_string + sas_key)
    
    Check if current version file exixts in azure blobstorage or not (Check EAGS_App_version.exe version and compare with code file)

    if file not exists: (If version not matches then:)
        Get list of all files available

        Get List of files and folder files available in blob storage

        rename those files and folder files as _Old

        Download all data from Blob storage

        Restart new EAGS_App.exe and delete all files and folder files with _Old

    """





def appUpdater(root, photo,curr_version, curr_location, curr_directory, currFilename):
    """Updates app version based on app version present inside exe and compare it with version present in name of exe parked in update location

    Args:
        curr_version (str): Fetches app curentversion from main app to be updated
        curr_location (str): current location of app to be Updated
        curr_directory (str): current directory of app to be Updated

    Raises:
        e: No version number was present in exe parked for updating other exe's    And Other Code Exceptions 
    """
    try:
        #Deleting files and folders with _Old
        for file in glob.glob(curr_directory+"\\*_Old.*"):
            os.remove(file)
        for file in glob.glob(curr_directory+"\\*\\*_Old.*"):
            os.remove(file)
        """Connect to Blob service client string using connection_string and sas_key(connect_str = connection_string + sas_key)

            Check EAGS_App_version.exe version and compare with code file"""


        # Initialize the connection to Azure storage account
        blob_service_client =  BlobServiceClient.from_connection_string(CONNECTION_STRING)
        my_container = blob_service_client.get_container_client(BLOB_CONTAINER)
        currFilename = os.path.splitext(currFilename)[0]
        
        #Check if current version is same or not
        if my_container.get_blob_client(f"{currFilename}_{curr_version}.exe").exists():
            return False

        else:
            top = tk.Toplevel(root)
            top.title('EAGS Quote Generator App Update')
            # messagebox.showinfo("App is Updating, Please Wait...")
            # top.geometry('320x180')
            width2 = 420
            height2 = 190
            screen_width = root.winfo_screenwidth()
            screen_height = root.winfo_screenheight()
            x2 = (screen_width/2) - (width2/2)
            y2 = (screen_height/2) - (height2/2)
            top.geometry('%dx%d+%d+%d' % (width2, height2, x2, y2))
            top.iconphoto(True, photo)
            top["bg"]= "white"

            

            updateLabel = tk.Label(top, text="App is Updating, Please Wait...", anchor="center", bg = "white", font=("Segoe UI", 12))
            # updateLabel.grid(row=1, column=1, sticky="nsew")
            updateLabel.place(relx=0.5, rely=0.5, anchor=tk.CENTER)
            
            top.update()
            
            """Get List of files and folder files available in blob storage"""
            #Renaming files as _Old in current directory and in subfolders
            f_list=[]
            my_blobs = my_container.list_blobs()
            for blob in my_blobs:
                print(blob.name)
                if ("EAGS_Quote_Generator" in blob.name) and (".exe" in blob.name):
                    updatedFilename = blob.name
                    #renaming curr exe
                    # os.rename(curr_location.replace('.py','.exe'), curr_directory+"\\EAGS_Quote_Generator_Old.exe")
                    os.rename(curr_location, curr_directory+"\\EAGS_Quote_Generator_Old.exe")
                f_list.append(blob.name)

            #Renaming current files as _Old and then downloading new ones
            for file in f_list:
                if file != '':
                    filename, fileExtesion = os.path.splitext(curr_directory+"\\"+file)
                    print(filename)
                    print(fileExtesion)
                    if os.path.exists(curr_directory+"\\"+file):
                        os.rename(curr_directory+"\\"+file, filename+"_Old"+fileExtesion)
                     
                    """Download all data from Blob storage"""
                    file_content = my_container.get_blob_client(blob).download_blob().readall()
                    # Get full path to the file
                    download_file_path = os.path.join(curr_directory, file)
                
                    # for nested blobs, create local path as well!
                    os.makedirs(os.path.dirname(download_file_path), exist_ok=True)
                
                    with open(download_file_path, "wb") as file:
                        file.write(file_content)
                    
            

           
            #rename exe version updatedFilename to EAGS_Quote_Generator.exe
            os.rename(curr_directory+"\\"+updatedFilename, curr_directory+"\\EAGS_Quote_Generator.exe")
            #trigger to open new EAGS_Quote_generator.exe
            os.startfile(curr_directory+"\\EAGS_Quote_Generator.exe")
            #sys.exit() current exe
            return True

            

    except Exception as e:
        msg = traceback.format_exc()
        msg = msg.replace('\n', '<br>')
        nl = '<br>'
        error_id = np.random.randint(1000000,9999999)
        tempDir = os.path.join(os.environ["HOMEPATH"], "Temp")
        tempDir = os.path.join('C:', tempDir)

        if not os.path.exists(tempDir):
            os.mkdir(tempDir)
       
        

        
        # dsp_msg = f"Error: {error_id}\nPlease send a screenshot of this error message along with the app window to devsupport@biourja.com"
        dsp_msg = f"A Report for this issue has been sent to IT India Team with Error ID: {error_id}, it will be resolved soon\nPlease capture this error id for future reference"
        
        
        
        send_mail(receiver_email='imam.khan@biourja.com, yashn.jain@biourja.com, devsupport@biourja.com', mail_subject=f"EAGS APP ERROR FOUND in Updater {error_id}", 
        mail_body=f"<strong>User: {curr_directory} Error ID: {error_id}</strong>{nl}{msg}")
        
        return False
    
    