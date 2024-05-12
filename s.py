import pandas as pd
import re

# 0all-sku.file
# 获取价格和PO号
def get_price(skus):
    price_list = {}
    df = pd.read_excel("0all-sku.xlsx")
    df.iloc[:, 0] =df.iloc[:, 0].str.strip()
    for sku in skus:
        subset = df[df.iloc[:, 0] == sku]
        # **********旧版[0, 13] 新版应该写[0, 14]***********
        price = subset.iloc[0, 13]
        price_list[sku] = price

    return price_list

def get_po(skus):
    df = pd.read_excel("0all-sku.xlsx")
    df.iloc[:, 0] =df.iloc[:, 0].str.strip()
    for i in range(1, len(df)):
        # 填充第一列空白-根据上一个值
        if pd.isnull(df.iloc[i, 0]):
            df.iat[i, 0] = df.iat[i-1, 0]

    result_dict = {}
    for sku in skus:
        result_dict[sku] = {}

        subset = df[df.iloc[:, 0] == sku]
        po = subset.iloc[:, -3]
        
        for index, row in subset.iterrows():
            # {'su888':{'web':77,'ws',66,'rt',55}}
            result_dict[sku][row.iloc[-2]] = row.iloc[-3]  
        
    return result_dict


def get_po_no_name(skus):
    df = pd.read_excel("0all-sku.xlsx")
    df.iloc[:, 0] =df.iloc[:, 0].str.strip()
    for i in range(1, len(df)):
        # 填充第一列空白-根据上一个值
        if pd.isnull(df.iloc[i, 0]):
            df.iat[i, 0] = df.iat[i-1, 0]

    result_dic = {}
    for sku in skus:
        subset = df[df.iloc[:, 0] == sku]
        if len(subset) > 0:
            po_ws = subset.iloc[0, -2]
            po_web = subset.iloc[1, -2]
            result_dic[sku] = {}
            result_dic[sku]['ws'] = po_ws
            result_dic[sku]['web'] = po_web

    return result_dic

def extract_skus_from_excel(file_path):
    df = pd.read_excel(file_path)
    column_values = df.iloc[:, 0].tolist()
    skus = set()
    for value in column_values:
        if isinstance(value, str) and re.match(r'^[A-Za-z0-9-]+', value):
            skus.add(value)

    return skus



