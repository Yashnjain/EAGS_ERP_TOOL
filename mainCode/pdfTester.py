# from tkPDFViewer import tkPDFViewer as pdf
# from tkinter import Tk, Button


# class ShowPdf(pdf.ShowPdf):
#     def goto(self, page):
#         try:
#             self.text.see(self.img_object_li[page - 1])
#         except IndexError:
#             if self.img_object_li:
#                 self.text.see(self.img_object_li[-1])

# root = Tk()

# pdfviewer = ShowPdf()
# # Add your pdf location and width and height.
# pdfframe = pdfviewer.pdf_view(root, pdf_location=r"C:/Users/imam.khan/OneDrive - BioUrja Trading LLC/Documents/EAGS/demoTest/5262022018.pdf", width=80, height=50)
# pdfframe.pack()


# Button(root, text="Go to page 3", command=lambda: pdfviewer.goto(1)).pack()
# root.mainloop()


from tkinter import *
from pandastable import Table, TableModel

class TestApp(Frame):
    def __init__(self, parent=None):
        self.parent = parent
        Frame.__init__(self)
        self.main = self.master
        self.main.geometry('600x400+200+100')
        self.main.title('Table app')
        f = Frame(self.main)
        f.pack(fill=BOTH,expand=1)
        df = TableModel.getSampleData()
        pt = Table(f, dataframe=df, showtoolbar=True, showstatusbar=True)
        for col_idx in range(len(df.columns)):
            col_name = pt.model.df.iloc[:, col_idx].name
            if pt.model.df[col_name].dtypes != 'object':
                pt.setColorByMask(
                    col_name, 
                    pt.model.df.iloc[:, col_idx] == pt.model.df.iloc[:, col_idx].max(), 
                    'lightgreen'
                )
        pt.show()
        return

app = TestApp()
app.mainloop()