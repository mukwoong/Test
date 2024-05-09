import pandas as pd

def fill_blank_cells(df):
    for i in range(1, len(df)):
        # 填充第一列空白-根据上一个值
        if pd.isnull(df.iloc[i, 0]):
            df.iat[i, 0] = df.iat[i-1, 0]
    return df

df = pd.read_excel("1ws.xlsx")
# 提前处理空白项目
df = fill_blank_cells(df)

# 示例1
subset = df[df.iloc[:, 0] == "SP24322-1"]
def get_size_sum():
    # 尺码(4-12列)S
    size = subset.iloc[:, 3:12]
    # 这里不能直接将nan转成int，后面相加之后的float才可以
    size = size.astype(float)
    col_sum = size.sum(axis=0)
    # tolist()将df列转成列表S
    col_sum_list = col_sum.astype(int).tolist()

    print(col_sum_list)
# 价格(14列)
#price = subset.iloc[0, 14]

def get_po():
    # po号
    po = subset.iloc[:, -3]
    print(po)
    result_dict = {'SU24215-1': {}}

    for index, row in subset.iterrows():
        # [row.iloc[-2]]是key如"web",-3是值
        result_dict['SU24215-1'][row.iloc[-2]] = row.iloc[-3]

    print(result_dict)

