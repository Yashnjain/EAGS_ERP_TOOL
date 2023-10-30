import os,logging
from office365.runtime.auth.user_credential import UserCredential
from office365.runtime.auth.authentication_context import AuthenticationContext
from office365.sharepoint.client_context import ClientContext
from datetime import datetime,date


def shpUploader(localpath, filename):
    try:

        baseurl = 'https://ealloys.sharepoint.com'

        basesite = '/sales' # every share point has a home.

        siteurl = baseurl + basesite
        
        # filename= localpath.split("\\")[-1]
        

        # localpath = r'C:\temp\Almez_000001.pdf'

        remotepath = 'Shared%20Documents/Quotations' # existing folder path under sharepoint site.

        username="itdevsupport@ealloys.com"

        # password="PJWelcome1$"
        password="4Gi.osDYQjM4C92x9iXnkW9w"

    

        ctx_auth = AuthenticationContext(baseurl)

        ctx_auth.acquire_token_for_user(username, password)

        ctx = ClientContext(siteurl, ctx_auth) # make sure you auth to the siteurl.



        with open(localpath, 'rb') as content_file:

            file_content = content_file.read()



        dir, name = os.path.split(remotepath)
        current_month = datetime.strftime(date.today(), "%B")
        current_year = datetime.strftime(date.today(), "%Y")
        folder_name = f"/{current_month}_{current_year}"
        folder_path = remotepath + folder_name
        target_folder = ctx.web.folders.add(folder_path)
        ctx.execute_query()


        file = ctx.web.get_folder_by_server_relative_url(folder_path).upload_file(filename, file_content).execute_query()

        logging.info("File uploaded to sharepoint")

    except Exception as e:

        print(e)

        logging.info(f"error in upload to sharepoint {e}")

        print("Done")
    





