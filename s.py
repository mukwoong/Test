import pandas as pd
import re

# 0all-sku.file
# 获取价格和PO号
def get_price(skus):
    df = pd.read_excel("0all-sku.xlsx")
    for sku in skus:
        subset = df[df.iloc[:, 0] == sku]
        # 价格(14列)
        price = subset.iloc[0, 14]

def get_po(skus):
    df = pd.read_excel("0all-sku.xlsx")
    for i in range(1, len(df)):
        # 填充第一列空白-根据上一个值
        if pd.isnull(df.iloc[i, 0]):
            df.iat[i, 0] = df.iat[i-1, 0]
    for sku in skus:
        subset = df[df.iloc[:, 0] == sku]
        # po号
        po = subset.iloc[:, -3]
        result_dict = {sku: {}}

        for index, row in subset.iterrows():
            # [row.iloc[-2]]是key如"web",-3是值
            result_dict[sku][row.iloc[-2]] = row.iloc[-3]
        return result_dict


def extract_skus_from_excel(file_path):
    df = pd.read_excel(file_path)
    column_values = df.iloc[:, 0].tolist()
    skus = set()
    for value in column_values:
        if isinstance(value, str) and re.match(r'^[A-Za-z0-9-]+', value):
            skus.add(value)

    return skus


# uniqe_ws_skus = extract_skus_from_excel("1ws.xlsx")
# print(uniqe_ws_skus)

uniqe_ws_skus = {'SU24637-1', 'SU24546-1', 'SU24619-2', 'SU24210-2', 'SU24210-1'}

result = get_po(uniqe_ws_skus)
data_list = [ast.literal_eval(item) for item in data_str.split('\n')]

# 打印每个键对应的 'RT' 值
for item in data_list:
    for key, value in item.items():
        if 'RT' in value:
            print(f"Key: {key}, RT Value: {value['RT']}")

#如果多个
# merged_set = ws_skus.union(web_skus)
