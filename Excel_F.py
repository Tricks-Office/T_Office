import pandas as pd

# df1 = pd.DataFrame({'key': ['b', 'b', 'a', 'c', 'a', 'a', 'b'],
#              'dat a1': range(7)})
#
# print(df1['dat a1'])

df=pd.read_excel("Test.xlsx")
df1=pd.read_excel("Test2.xlsx")
# frames = pd.concat([df1, df])
frames = df1.append(df,sort=False)
frames.to_excel("Result.xlsx", index=False)
print(frames)
# print(df['수령인 주소(전체)'])


#
# df1=pd.read_excel("Compare1.xlsx")
# df2=pd.read_excel("Compare2.xlsx")
#
#
#
# ds1 = set([tuple(line) for line in df1.values])
# ds2 = set([tuple(line) for line in df2.values])
#
# frame=pd.DataFrame(list(ds1.difference(ds2)))
# print(frame)


# frames = pd.concat([df, df1])
# frames.to_excel("Result.xlsx", index=False)
# print(frames)


# l_row=0
# df=pd.read_excel("test.xlsx", header=None)
# for x in df.index:
#     z=df.columns.size-1
#     y = df.values[x,df.columns.size-1]
#     if not pd.isna(df.values[x,df.columns.size-1]) :
#         l_row = x
#         break
#
# df=pd.read_excel("test.xlsx", skiprows = l_row)
#
# print(df)
