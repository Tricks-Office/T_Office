import win32com.client
import pandas as pd

hwp = win32com.client.gencache.EnsureDispatch('HWPFrame.HwpObject')
df = pd.read_excel("Master/Dictionary.xlsx")

hwp.RegisterModule('FilePathCheckDLL','SecurityModule')

hwp.Run("FileNew")
hwp.Open("D:/Python/TOffice/Master/Sample.hwp")

hwp.MovePos(2)
hwp.MovePos(3)

hwp.MovePos("moveTopOfFile")
hwp.MovePos("moveBottomOfFile")

hwp.Run("MoveWordBegin")
hwp.Run("MoveWordEnd")

hwp.MovePos(2)
hwp.Run("Select")
hwp.Run("MoveLineEnd")

hwp.SaveAs("D:/Python/TOffice/Master/Result.hwp")
hwp.Quit()