from t import get_size_sum
import pandas as pd
from s import get_po
from s import get_po_no_name
from i import get_invoice

def multiple():
    web = get_size_sum('2web.xlsx')
    ws = get_size_sum('1ws.xlsx')
    # 以上都是字典，转成df是更好的计算和展示
    web_df = pd.DataFrame(web)
    ws_df = pd.DataFrame(ws)
    sum_df = pd.concat([ws_df, web_df],axis=1)
    result = sum_df.T.groupby(level=0).sum()
    # 输出尺码和
    print(result)

    # 方法2是使用union，merged_set = ws_skus.union(web_skus)
    merged_sku = set(ws.keys()) | set(web.keys())

    po_reslut = get_po_no_name(merged_sku)
    po_df = pd.DataFrame(po_reslut).T
    print(po_df)

    get_invoice(merged_sku)

def single():
    # ***********需改文件名*************
    ws = get_size_sum('1ws.xlsx')
    # *********************************
    ws_df = pd.DataFrame(ws).T
    # 尺码和
    print(ws_df)
    sku_set = set(ws.keys())

    result = get_po_no_name(sku_set)
    df = pd.DataFrame(result).T
    print(df)

    get_invoice(sku_set)

#single()
#如果多个
multiple()


