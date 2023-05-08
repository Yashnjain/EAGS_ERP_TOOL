Pyinstaller -F --paths "C:\Users\imam.khan\Documents\EAGS_New_Test\venv\Lib\site-packages" --hidden-import tkPDFViewer --icon=biourjaLogo.ico  --add-data "biourjaLogo.png;." --add-data "sound1.png;." --add-data "Entry1.png;." --add-data "Entry1New.png;." --add-data "Entry2.png;." --add-data "Entry2New.png;." --add-data "Entry3.png;." --add-data "Entry3New.png;." --add-data "Entry4.png;." --add-data "Entry4New.png;." --add-data "center.png;." --add-data "home(2).png;." --add-data "home(4).png;." --add-data "addRowS.png;." --add-data "addRow2S.png;." --add-data "deleteRowS.png;." --add-data "deleteRow2S.png;." --add-data "previewButtonS.png;." --add-data "previewButton2S.png;." --add-data "submitButtonS.png;." --add-data "submitButton2S.png;." --add-data "reportGenerator1.png;." --add-data "reportGenerator2.png;." --add-data "off.png;." --add-data "on.png;." --add-data "pdfCreator.xlsm;." --add-data "c:\users\imam.khan\appdata\roaming\python\python38\site-packages\customtkinter;customtkinter\" --hidden-import tkcalendar --hidden-import "snowflake-connector-python" --hidden-import "xlwings" --hidden-import "babel.numbers" EAGS_Quote_Generator.py --onedir -w




Update these lines in pandas table core.py
# self.parentframe.master.bind_all("<KP_8>", self.handle_arrow_keys)
# self.parentframe.master.bind_all("<Return>", self.handle_arrow_keys)
# self.parentframe.master.bind_all("<Tab>", self.handle_arrow_keys)
self.master.bind("<KP_8>", self.handle_arrow_keys)
self.master.bind("<Return>", self.handle_arrow_keys)
self.master.bind("<Tab>", self.handle_arrow_keys)

Add try except in these lines in core.py around 2299
# if rowover >= self.rows or self.startrow > self.rows:
#     return
# else:
#     self.endrow = rowover
try:
    if rowover >= self.rows or self.startrow > self.rows:
        return
    else:
        self.endrow = rowover
except:
    return

#do columns
# if colover > self.cols or self.startcol > self.cols:
#     return
# else:
#     self.endcol = colover
#     if self.endcol < self.startcol:
#         self.multiplecollist=list(range(self.endcol, self.startcol+1))
#     else:
#         self.multiplecollist=list(range(self.startcol, self.endcol+1))
try:
    if colover > self.cols or self.startcol > self.cols:
        return
    else:
        self.endcol = colover
        if self.endcol < self.startcol:
            self.multiplecollist=list(range(self.endcol, self.startcol+1))
        else:
            self.multiplecollist=list(range(self.startcol, self.endcol+1))
except:
    return


Update these lines in tkPDFViewer
# def pdf_view(self,master,width=1200,height=600,pdf_location="",bar=True,load="after"):
def pdf_view(self,master,width=1200,height=600,pdf_location="",bar=True,load="after", zoomDPI=72):

# pix = page.getPixmap()
pix = page.get_pixmap(dpi=zoomDPI)

# img = pix1.getImageData("ppm")
img = pix1.tobytes("ppm")

Add this line in tkpdfviewer
self.text.delete('1.0', END)
before
self.text.image_create(END,image=i)


