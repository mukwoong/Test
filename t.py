import pandas as pd

def fill_blank_cells(df):
    for i in range(1, len(df)):
        # 填充第一列空白-根据上一个值
        if pd.isnull(df.iloc[i, 0]):
            df.iat[i, 0] = df.iat[i-1, 0]
    return df

df = pd.read_excel("0all-sku.xlsx")
# 提前处理空白项目
df = fill_blank_cells(df)

# 示例1
subset = df[df.iloc[:, 0] == "SU24215-1"]

# 尺码(5-13列)
size = subset.iloc[:, 5:14]
# 价格(14列)
price = subset.iloc[0, 14]
# po号
po = subset.iloc[:, -2]


print(po)