import pandas as pd
import re
from collections import defaultdict

# packing file
# 获取sku集合，尺码数量

def get_size_sum(file):
    df = pd.read_excel(file)
    column_values = df.iloc[:, 0].tolist()
    skus = set()
    for value in column_values:
        if isinstance(value, str) and re.match(r'^[A-Za-z0-9-]+', value):
            skus.add(value)


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



