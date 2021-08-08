from docx import Document
import pandas as pd

doc = Document("Master/Sample.docx")
df = pd.read_excel("Master/Dictionary.xlsx")

for x in df.index:
    for paragraph in doc.paragraphs:
        for run in paragraph.runs:
            cur_text = run.text
            new_text = cur_text.replace(df.Before[x], df.Change[x])
            run.text = new_text

doc.save("Master/Result.docx")
