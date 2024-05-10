import pandas as pd

df = pd.read_excel("4hs-code.xlsx")

df.columns.values[0] = 'D'
df.columns.values[1] = 'SKU'
df.columns.values[2] = 'F'
df.columns.values[3] = 'DESCRIPTION'
df.columns.values[4] = 'COLOR'
df.columns.values[5] = 'FD'
df.columns.values[6] = 'FABRIC'
df.columns.values[7] = 'WK'
df.columns.values[8] = 'CODE'

# 排列规则
desired_columns = ['SKU', 'DESCRIPTION', 'FABRIC', 'CODE', 'COLOR', 'WK']
# 使用loc，选中所有行，和对应的列。列可以是数组
subset = df.loc[:, desired_columns]


print(subset.iloc[2:7, 0:5])