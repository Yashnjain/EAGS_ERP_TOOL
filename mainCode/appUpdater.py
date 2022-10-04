import glob
import os, shutil, traceback
from mail import send_mail
import numpy as np





UPDATE_LOC = r'C:\Users\imam.khan\Documents\EAGS_AppUpdate'
# UPDATE_LOC = r'C:\Users\imam.khan\OneDrive - BioUrja Trading LLC\Documents\EAGS\FinalCodePreRequisites\FinalCodePrep\dist\EAGS_Quote_Generator'




def recursive_overwrite(src, dest, ignore=None):
    if os.path.isdir(src):
        if not os.path.isdir(dest):
            os.makedirs(dest)
        files = os.listdir(src)
        if ignore is not None:
            ignored = ignore(src, files)
        else:
            ignored = set()
        for f in files:
            if f not in ignored:
                recursive_overwrite(os.path.join(src, f), 
                                    os.path.join(dest, f), 
                                    ignore)
    else:
        shutil.copyfile(src, dest)





def appUpdater(curr_version, curr_location, curr_directory):
    """Updates app version based on app version present inside exe and compare it with version present in name of exe parked in update location

    Args:
        curr_version (str): Fetches app curentversion from main app to be updated
        curr_location (str): current location of app to be Updated
        curr_directory (str): current directory of app to be Updated

    Raises:
        e: No version number was present in exe parked for updating other exe's    And Other Code Exceptions 
    """
    try:
        updatedFile = glob.glob(UPDATE_LOC+"\\*.exe")[0]

        updatedVersion = updatedFile.split("\\")[-1].split("_")[-1].split(".exe")[0]
        updatedFilename = updatedFile.split("\\")[-1]
        if updatedVersion == 'Generator':
            raise Exception("No version number was present in exe parked for updating other exe's")
        if updatedVersion == curr_version:
            return False
        else:
            #renaming curr exe
            os.rename(curr_location.replace('.py','.exe'), curr_directory+"\\EAGS_Quote_Generator_Old.exe")
            #delete EAGS_Quote_Generator_old.py on startup
            #checkUpdater function called
            #rename orignal file EAGS_Quote_Generator_old.py
            #download udpated version in curr location
            src = UPDATE_LOC
            dest = curr_directory
            recursive_overwrite(src, dest)

            
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
        # directories_created = [tempDir]
        # for directory in directories_created:
        #     path3 = os.path.join(os.getcwd(),directory)  
        #     try:
        #         os.makedirs(path3, exist_ok = True)
        #         print("Directory '%s' created successfully" % directory)
        #     except OSError as error:
        #         print("Directory '%s' can not be created" % directory) 

        # cap = tkcap.CAP(root)     # master is an instance of tkinter.Tk
        # imageV1path = os.path.join(tempDir,f'{error_id}V1.png')
        # cap.capture(f'{error_id}V1.png')
        # cap.capture(imageV1path)
        # imageV1path = os.getcwd()+'\\'+f'{error_id}V1.png'
        

        
        # dsp_msg = f"Error: {error_id}\nPlease send a screenshot of this error message along with the app window to devsupport@biourja.com"
        dsp_msg = f"A Report for this issue has been sent to IT India Team with Error ID: {error_id}, it will be resolved soon\nPlease capture this error id for future reference"
        # showerror(f"Error", message=dsp_msg)
        # imageV2path = os.path.join(tempDir,f'{error_id}V2.png')
        # cap = tkcap.CAP(root)     # master is an instance of tkinter.Tk
        # cap.capture(imageV2path)       # Capture and Save the screenshot of the tkiner window
        # imageV2path = os.getcwd()+'\\'+f'{error_id}V2.png'
        
        
        send_mail(receiver_email='imam.khan@biourja.com, yashn.jain@biourja.com, devsupport@biourja.com', mail_subject=f"EAGS APP ERROR FOUND in Updater {error_id}", 
        mail_body=f"<strong>User: {curr_directory} Error ID: {error_id}</strong>{nl}{msg}")
        
        return False
    
    