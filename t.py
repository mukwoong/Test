import pandas as pd
import re
from collections import defaultdict

# packing file
# 获取sku集合，尺码数量

def get_size_sum(file, skus):
    df = pd.read_excel(file)
    for i in range(1, len(df)):
        # 填充第一列空白-根据上一个值
        if pd.isnull(df.iloc[i, 0]):
            df.iat[i, 0] = df.iat[i-1, 0]


    sku_dict = {}
    for sku in skus:
        subset = df[df.iloc[:, 0] == sku]
        # 尺码(4-12列)
        size = subset.iloc[:, 3:12]
        # 这里不能直接将nan转成int，后面相加之后的float才可以
        size = size.astype(float)
        col_sum = size.sum(axis=0)
        # tolist()将df列转成列表S
        col_sum_list = col_sum.astype(int).tolist()
        sku_dict[sku] = col_sum_list
    return sku_dict

# 去重的SKU号
def extract_skus_from_excel(file_path):
    df = pd.read_excel(file_path)
    column_values = df.iloc[:, 0].tolist()
    skus = set()
    for value in column_values:
        if isinstance(value, str) and re.match(r'^[A-Za-z0-9-]+', value):
            skus.add(value)

    return skus


uniqe_ws_skus = extract_skus_from_excel("1ws.xlsx")
uniqe_web_skus = extract_skus_from_excel("2web.xlsx")
print(uniqe_ws_skus)

# 尺码数量前提前处理空白项目
ws_size = get_size_sum("1ws.xlsx", uniqe_ws_skus)

web_size = get_size_sum("2web.xlsx", uniqe_web_skus)

# 如果有两个文件，或者3个文件。求和
ws = pd.DataFrame(ws_size)
web = pd.DataFrame(web_size)
df = pd.concat([ws, web],axis=1)
result = df.T.groupby(level=0).sum()

# print(result)


