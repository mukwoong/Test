import pandas as pd

def fill_blank_cells(df):
    for i in range(1, len(df)):
        # 填充第一列空白-根据上一个值
        if pd.isnull(df.iloc[i, 0]):
            df.iat[i, 0] = df.iat[i-1, 0]
    return df

df = pd.read_excel("0all-sku.xlsx")
df = fill_blank_cells(df)
subset = df.iloc[:, 0]
print(subset[30: 40])