from tkinter import *
from tkinter import filedialog
from tkinter import font
import pandas as pd
import os
from tkinter import messagebox
import PPT_F
import Com_F
import re

# Master Data 저장하기
def saveMaster():
    writer = pd.ExcelWriter('Master/Master.xlsx', engine='xlsxwriter')
    df.to_excel(writer, sheet_name='Master', index=False)
    dff.to_excel(writer, sheet_name='FileName', index=False)
    writer.save()
    return

# 폴더 경로 변경
def onClick(i):
    folder_selected = filedialog.askdirectory()
    lbPath[i].config(text=folder_selected)
    df.Detail[i] = folder_selected
    saveMaster()
    return

# 파일경로 변경
def onClickf(i):
    folder_selected = filedialog.askopenfile()
    lbPathf[i].config(text=folder_selected.name)
    dff.Detail[i] = folder_selected.name
    saveMaster()
    return

# Dictionary 기준, Keyword 변경
def PPT_KC_Action():
    files = [f for f in os.listdir(df.Detail[0]) if re.match('.*[.]ppt', f)]
    for file in files:
        PPT_F.PPT_KC(file,df.Detail[0],df.Detail[1],dff.Detail[0])
    messagebox.showinfo(title="Finish", message="Job Finished!")
    return

# PPT Format 통일 (python-pptx)
def UFontAc():
    files = [f for f in os.listdir(df.Detail[0]) if re.match('.*[.]ppt', f)]
    for file in files:
        PPT_F.AutoFont(file, df.Detail[0], df.Detail[1],
                          select.get(select.curselection()))
    messagebox.showinfo(title="Finish", message="Job Finished!")
    return 

# PPT Format 통일 (win 32)
def U2FontAc():
    files = [f for f in os.listdir(df.Detail[0]) if re.match('.*[.]ppt', f)]
    for file in files:
        PPT_F.AutoFont2(file, df.Detail[0], df.Detail[1],
                          select.get(select.curselection()))
    messagebox.showinfo(title="Finish", message="Job Finished!")
    return

# 영어로 구글 번역 (google-trans)
def GTrans2EnAc():
    files = [f for f in os.listdir(df.Detail[0]) if re.match('.*[.]ppt', f)]
    for file in files:
        PPT_F.GTrans(file, df.Detail[0], df.Detail[1], 'en')
    messagebox.showinfo(title="Finish", message="Job Finished!")
    return

# 한글로 구글 번역 (google-trans)
def GTrans2KoAc():
    files = [f for f in os.listdir(df.Detail[0]) if re.match('.*[.]ppt', f)]
    for file in files:
        PPT_F.GTrans(file, df.Detail[0], df.Detail[1], 'ko')
    messagebox.showinfo(title="Finish", message="Job Finished!")
    return

# Folder에서 File List 만들기
def Folder_FileList():
    Com_F.FileList(df.Detail[2],df.Detail[1])
    messagebox.showinfo(title="Finish", message="Job Finished!")
    return


# 창 호출 & 초기값 설정
win = Tk()
win.geometry("600x700")
win.title('Tricks_Office')

df= pd.read_excel("Master/Master.xlsx",sheet_name="Master")
dff= pd.read_excel("Master/Master.xlsx",sheet_name="FileName")

fonts=list(font.families())

# fram1(폴더경로) 설정
frame1 = Frame(win,pady=10)
frame1.pack()
lbName = []
lbPath = []
btnPath =[]

for i in df.index:
    lbName.append(Label(frame1, text=df.Item[i], width=10))
    lbName[i].grid(row=i, column=0, sticky=W,padx=5, pady=5)
    lbPath.append(Label(frame1, text=df.Detail[i], width=50,height=2,relief="groove"))
    lbPath[i].grid(row=i, column=1, sticky=W)
    btnPath.append(Button(frame1, text="Change Path", width=10,
                          command=lambda i=i: onClick(i)))
    btnPath[i].grid(row=i, column=2, sticky=W)

# frame1h(파일 경로) 설정
frame1h = Frame(win,pady=10)
frame1h.pack()
lbNamef = []
lbPathf = []
btnPathf =[]

for i in dff.index:
    lbNamef.append(Label(frame1h, text=dff.Item[i], width=10))
    lbNamef[i].grid(row=i, column=0, sticky=W,padx=5)
    lbPathf.append(Label(frame1h, text=dff.Detail[i], width=50,
                         height=2,relief="groove"))
    lbPathf[i].grid(row=i, column=1, sticky=W)
    btnPathf.append(Button(frame1h, text="Change Path", width=10,
                           command=lambda i=i: onClickf(i)))
    btnPathf[i].grid(row=i, column=2, sticky=W)

# framePPT 실행 버튼 설정 (PPT전용)
framePPT = Frame(win, padx=10, pady=10,bd=1,relief=SOLID)
framePPT.pack()

# PPT전용 표기
Label(framePPT,text = "PPT 전용").grid(row=0,column=1)

# Keywork Change 버튼
btnPPT_KC = Button(framePPT,text="Find and Replace keywords",width = 21,
                   height=2, command=PPT_KC_Action)
btnPPT_KC.grid(row=1, column=1, sticky=W,padx=5)
lbPPT_KC = Label(framePPT, text="Dictionary 정보를 기반으로, Keywords를 바꿔주는 기능입니다. \n"
                              "Dictionary 파일 경로가 정확하게 지정되어 있어야 합니다.",
                 width=50,height=2,relief="groove")
lbPPT_KC.grid(row=1, column=2, sticky=W)

# Font 통일1
UFontbtn = Button(framePPT, text="Unify to Same Font",width = 21,height=2,
                  command=UFontAc)
UFontbtn.grid(row=3, column=1,padx=5)
UFontlbl = Label(framePPT, text="아래 리스트에서 폰트가 선택되어 있어야 합니다.\n" +
                             "속도가 매우 빠르지만, 한글 폰트가 안바뀔수 있습니다.",
                 width=50,height=2,relief="groove")
UFontlbl.grid(row=3, column=2)

# Font 통일2
U2Fontbtn = Button(framePPT, text="Unify to Same Font 2",width = 21,height=2,
                   command=U2FontAc)
U2Fontbtn.grid(row=4, column=1,padx=5)
U2Fontlbl = Label(framePPT, text="아래 리스트에서 폰트가 선택되어 있어야 합니다.\n" +
                              "속도는 좀 느립니다.",width=50,height=2,relief="groove")
U2Fontlbl.grid(row=4, column=2)

# Google Trans 2 En
GTrans2Enbtn = Button(framePPT, text="Google Translate 2 En",width = 21,
                      height=2, command=GTrans2EnAc)
GTrans2Enbtn.grid(row=5, column=1,padx=5)
GTrans2Enlbl = Label(framePPT, text="영어로 번역하며, 번역 가능 횟수에 제한이 있을 수 있습니다.",
                     width=50,height=2,relief="groove")
GTrans2Enlbl.grid(row=5, column=2)

# Google Trans 2 Ko
GTrans2Kobtn = Button(framePPT, text="Google Translate 2 Ko",width = 21,
                      height=2,command=GTrans2KoAc)
GTrans2Kobtn.grid(row=6, column=1,padx=5)
GTrans2Kolbl = Label(framePPT,
                     text="한글로 번역하며, 번역 가능 횟수에 제한이 있을 수 있습니다.",
                     width=50,height=2,relief="groove")
GTrans2Kolbl.grid(row=6, column=2)

# frameCom 실행 버튼 설정 (공통)
frameCom = Frame(win, padx=10, pady=10,bd=1,relief=SOLID)
frameCom.pack()

# PPT전용 표기
Label(frameCom,text = "공통").grid(row=0,column=0)

# File List
DFbtn = Button(frameCom, text="Directory File List",width = 21,
                      height=2,command=Folder_FileList)
DFbtn.grid(row=1, column=0,padx=5)
DFlbl = Label(frameCom, text="Target 디렉토리의 파일 리스트를 만들어 결과 폴더에 저장.\n"
                          "Target 디렉토리는 반드시 설정되어 있어야 합니다.",
                     width=50,height=2,relief="groove")
DFlbl.grid(row=1, column=1)

# frame3 Font 리스트 상자 만들기
frame3 = Frame(win,padx=5, pady=10,)       # select of names
frame3.pack()
scroll = Scrollbar(frame3, orient=VERTICAL)
select = Listbox(frame3, yscrollcommand=scroll.set, height=16,width = 60)
scroll.config (command=select.yview)
scroll.pack(side=RIGHT, fill=Y)
select.pack(side=LEFT,  fill=BOTH, expand=1)
select.delete(0,END)
for item in fonts :
    select.insert (END, item)
select.select_set(0)

win.mainloop()

