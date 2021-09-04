import win32com.client
import pandas as pd

hwp = win32com.client.gencache.EnsureDispatch('HWPFrame.HwpObject')
df = pd.read_excel("D:/Python/TOffice/Master/Dictionary.xlsx")

hwp.RegisterModule('FilePathCheckDLL','SecurityModule')
# hwp.Run("FileNew")
hwp.Open("D:/Python/TOffice/Master/Sample.hwp")

for x in df.index:
    hwp.HAction.GetDefault("AllReplace", hwp.HParameterSet.HFindReplace.HSet)
    hwp.HParameterSet.HFindReplace.FindString = df.Before[x]
    hwp.HParameterSet.HFindReplace.ReplaceString = df.Change[x]
    hwp.HParameterSet.HFindReplace.IgnoreMessage = 1
    hwp.HAction.Execute("AllReplace", hwp.HParameterSet.HFindReplace.HSet)

hwp.SaveAs("D:/Python/TOffice/Master/Result.hwp")
hwp.Quit()