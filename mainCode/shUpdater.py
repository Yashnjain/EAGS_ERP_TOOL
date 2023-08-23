import sharepy
import glob
import os, shutil, traceback, time
from mail import send_mail
import tkinter as tk
import numpy as np
from tkinter import messagebox
from office365.runtime.auth.client_credential import ClientCredential
from office365.runtime.client_request_exception import ClientRequestException
from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.authentication_context import AuthenticationContext
from office365.sharepoint.files.file import File
import io
import datetime
import pandas as pd







# sp_path1 = "/it/_api/web/GetFolderByServerRelativeUrl"
# sp_path2 = "Shared%20Documents/EAGS%20Quote%20Generator/Updater/"



baseurl = 'https://ealloys.sharepoint.com'

basesite = '/it' # every share point has a home.

siteurl = baseurl + basesite

# filename= localpath.split("\\")[-1]


# localpath = r'C:\temp\Almez_000001.pdf'

relative_url = "Shared%20Documents/EAGS%20Quote%20Generator/Updater/" # existing folder path under sharepoint site.

# username="priyanshi.jhawar@ealloys.com"
username="itdevsupport@ealloys.com"

# password="PJWelcome1$"
password="4Gi.osDYQjM4C92x9iXnkW9w"

sp_path1 = "/it/_api/web/GetFolderByServerRelativeUrl"
sp_path2 = "Shared%20Documents/EAGS%20Quote%20Generator/Updater/"


def printAllContents(ctx, relativeUrl,f_list):

    try:
        
        libraryRoot = ctx.web.get_folder_by_server_relative_url(relativeUrl)
        ctx.load(libraryRoot)
        ctx.execute_query()

        folders = libraryRoot.folders
        ctx.load(folders)
        ctx.execute_query()

        for myfolder in folders:
            print("Folder name: {0}".format(myfolder.properties["ServerRelativeUrl"]))
            f_list=printAllContents(ctx, relativeUrl + '/' + myfolder.properties["Name"],f_list)
            
        files = libraryRoot.files
        ctx.load(files)
        ctx.execute_query()

        for myfile in files:
            #print("File name: {0}".format(myfile.properties["Name"]))
            print("File name: {0}".format(myfile.properties["ServerRelativeUrl"]))
            f_list.append(myfile.properties["ServerRelativeUrl"])
        return f_list
    except Exception as e: 
        raise e



# def get_f_list_from_sp(s):
#   # Get list of all files and folders in library
#   r = s.get(baseurl + f"""{sp_path1}('"""+sp_path2+"""')/Files""")
#   files = r.json()['d']['results']
#   f_lst = []
#   for file in files:
#     print(file["Name"])
#     f_lst.append(file["Name"])
#   print("done")
#   return f_lst



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


        # Initialize the connection to Sharepoint account
        # s = sharepy.connect(baseurl, username, password)
        # f_lst = get_f_list_from_sp(s)
        retry=0
        while retry<3:
            try:
                ctx_auth = AuthenticationContext(baseurl)

                ctx_auth.acquire_token_for_user(username, password)

                ctx = ClientContext(siteurl, ctx_auth) # make sure you auth to the siteurl.]
                break
            except Exception as e:
                time.sleep(2)
                if retry==2:
                    raise e
                retry+=1

        currFilename = os.path.splitext(currFilename)[0]

        #Get all files and directories present in sharepoint location
        f_list = []
        
        f_list = printAllContents(ctx, relative_url,f_list)
        #Check if current version is same or not
        if f"{currFilename}_{curr_version}.exe" in " ".join(f_list):
            return False
        elif len(f_list)==0:
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
            top.grab_set()
            

            updateLabel = tk.Label(top, text="App is Updating, Please Wait...", anchor="center", bg = "white", font=("Segoe UI", 12))
            # updateLabel.grid(row=1, column=1, sticky="nsew")
            updateLabel.place(relx=0.5, rely=0.5, anchor=tk.CENTER)
            
            top.update()
            
            """Get List of files and folder files available in blob storage"""
            #Renaming files as _Old in current directory and in subfolders
            
            #Renaming current files as _Old and then downloading new ones
            for file_url in f_list:
                if file_url != '':
                    file = file_url.split('Updater/')[-1]
                    if ("EAGS_Quote_Generator" in file) and (".exe" in file):
                        updatedFilename = file
                        
                    
                    filename, fileExtesion = os.path.splitext(curr_directory+"\\"+file)
                    print(filename)
                    print(fileExtesion)
                    if os.path.exists(curr_directory+"\\"+file):
                        os.rename(curr_directory+"\\"+file, filename+"_Old"+fileExtesion)
                     
                    """Download all data from Sharepoint"""
                    
                    # Get full path to the file
                    download_file_path = os.path.join(curr_directory, file)
                
                    # for nested blobs, create local path as well!
                    os.makedirs(os.path.dirname(download_file_path), exist_ok=True)
                
                    #Downloading file
                    with open(download_file_path, "wb") as local_file:
                        file = ctx.web.get_file_by_server_relative_url(file_url)  
                        file.download(local_file)
                        ctx.execute_query()
                    # path1 = "/itdev/_api/web/GetFolderByServerRelativeUrl"
                    # path2 = "('Shared Documents{0}')/Files('{1}')/$value"

                    # sp_path1 = "/it/_api/web/GetFolderByServerRelativeUrl"
                    # sp_path2 = "Shared%20Documents/EAGS%20Quote%20GeneratorUpdater/"

                    
                    # r = s.getfile(site + sp_path1+sp_path2+"Files('{1}')/$value".format(
                    #                                 file), filename=download_file_path)
            def on_closing():
                try:
                    top.attributes('-topmost', True)
                    messagebox.showeinfo("Info", f"Please wait for app update to complete",parent=top)
                    top.attributes('-topmost', False)
                    return
                except Exception as e:
                    raise e
        
        
            top.protocol("WM_DELETE_WINDOW", on_closing)        
            
            if updatedFilename:
                #renaming curr exe
                os.rename(curr_location.replace('.py','.exe'), curr_directory+"\\EAGS_Quote_Generator_Old.exe")
                # os.rename(curr_location, curr_directory+"\\EAGS_Quote_Generator_Old.exe")
                #rename exe version updatedFilename to EAGS_Quote_Generator.exe
                os.rename(curr_directory+"\\"+updatedFilename, curr_directory+"\\EAGS_Quote_Generator.exe")
                #trigger to open new EAGS_Quote_generator.exe
                os.startfile(curr_directory+"\\EAGS_Quote_Generator.exe")
                #sys.exit() current exe
                return True
            else:
                return False

            

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

















   


# site = 'https://ealloys.sharepoint.com'

# # username = "svc_tableauonline@biourja.com"
# username = "priyanshi.jhawar@ealloys.com"
# # password = "L!,'W%^9#@}rzf6NGyZKwz"
# password = "PJWelcome1$"






