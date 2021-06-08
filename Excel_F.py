import pandas as pd

def Excel_M(file,SPath,df):
    l_row=0
    df1=pd.read_excel(SPath + "/" + file, header=None)
    for x in df.index:
        if not pd.isna(df1.values[x,df1.columns.size-1]) :
            l_row = x
            break

    df1=pd.read_excel(SPath + "/" + file, skiprows = l_row)

    frames = pd.concat([df1, df]).drop_duplicates()
    return frames
