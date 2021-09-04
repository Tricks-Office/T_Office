import os
import re
import win32com.client
from docx2pdf import convert
from fpdf import FPDF

TDPath = "D:\\Data\\ToDo"
RPath = "D:\\Data\\Result"

#PPT to PDF
powerpoint = win32com.client.Dispatch("Powerpoint.Application")
files = [f for f in os.listdir(TDPath) if re.match('.*[.]ppt', f)]
for file in files:
    #PPT 파일을 PDF로 바꾸는 로직
    deck = powerpoint.Presentations.Open(os.path.join(TDPath,file))
    pre, ext = os.path.splitext(file)
    deck.SaveAs(os.path.join(RPath, pre + ".pdf"), 32)  # formatType = 32 for ppt to pdf
    deck.Close()
powerpoint.Quit()


# Word to PDF
files = [f for f in os.listdir(TDPath) if re.match('.*[.]doc', f)]
for file in files:
    # Word 파일을 PDF로 바꾸는 로직
    pre, ext = os.path.splitext(file)
    convert(os.path.join(TDPath,file), os.path.join(RPath, pre + ".pdf"))

# HWP to PDF
hwp = win32com.client.gencache.EnsureDispatch('HWPFrame.HwpObject')
hwp.RegisterModule('FilePathCheckDLL','SecurityModule')
files = [f for f in os.listdir(TDPath) if re.match('.*[.]hwp', f)]
for file in files:
    # HWP 파일을 PDF로 바꾸는 로직
    hwp.Open(os.path.join(TDPath,file))
    hwp.HAction.GetDefault("FileSaveAs_S", hwp.HParameterSet.HFileOpenSave.HSet)
    pre, ext = os.path.splitext(file)
    hwp.HParameterSet.HFileOpenSave.filename = os.path.join(RPath, pre + ".pdf")
    hwp.HParameterSet.HFileOpenSave.Format = "PDF"
    hwp.HAction.Execute("FileSaveAs_S", hwp.HParameterSet.HFileOpenSave.HSet)
hwp.Quit()
#
# Image to PDF
files = [f for f in os.listdir(TDPath) if re.match('.*([.]jpg|[.]png|[.]gif)', f)]
pdf = FPDF('L')
for file in files:
    # Image 파일을 PDF로 바꾸는 로직
    pdf.add_page()
    pdf.image(os.path.join(TDPath,file), 0, 0, 330)
pdf.output(os.path.join(RPath, "IMG2PDF.pdf"), "F")
