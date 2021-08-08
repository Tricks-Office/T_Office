import win32com.client
import pandas as pd

hwp = win32com.client.gencache.EnsureDispatch('HWPFrame.HwpObject')
hwp.Run("FileNew")
hwp.RegisterModule('FilePathCheckDLL','SecurityModule')
hwp.Open("D:/Python/TOffice/Master/Sample.hwp")

df = pd.read_excel("Master/Dictionary.xlsx")

hwp.Run("MoveDocBegin")

for x in df.index:
    try:
        hwp.InitScan()
        while True:
            textdata = hwp.GetText()
            if textdata[0] == 1:
                break
            else:
                hwp.MovePos(201)
                text = textdata[1].strip()
                re_text=text.replace(df.Before[x],df.Change[x])
                if not re_text:
                    pass
                else:
                    hwp.Run("Select")
                    hwp.Run("MoveLineEnd")
                    hwp.HAction.GetDefault("InsertText",hwp.HParameterSet.HInsertText.HSet)
                    hwp.HParameterSet.HInsertText.Text = re_text
                    hwp.HAction.Execute("InsertText",hwp.HParameterSet.HInsertText.HSet)
                    hwp.Run("Cancel")
    finally:
        hwp.ReleaseScan()

hwp.SaveAs("D:/Python/TOffice/Master/Result.hwp")
hwp.Quit()