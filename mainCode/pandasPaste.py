import pandas as pd
import xlwings as xw



def pandasPaste():
    wb = xw.Book()
    wb.app.visible = False
    ws = wb.sheets[0]
    ws.api.Paste()
    ws.range("A1").expand('table').copy()

    df = pd.read_clipboard()
    # print(df)
    df.to_clipboard()

    wb.app.quit()
    return 