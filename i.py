import pandas as pd
from s import get_price


def get_invoice(skus):
    df = pd.read_excel("4hs-code.xlsx")

    #*****************原excel排列规则***************
    df.columns.values[0] = 'D'
    df.columns.values[1] = 'SKU'
    df.columns.values[2] = 'F'
    df.columns.values[3] = 'DESCRIPTION'
    df.columns.values[4] = 'COLOR'
    df.columns.values[5] = 'FD'
    df.columns.values[6] = 'FABRIC'
    df.columns.values[7] = 'WK'
    df.columns.values[8] = 'CODE'

    #*******************检索第一列SKU******************
    filtered_df = df[df.iloc[:, 1].isin(skus)]

    price = get_price(skus)
    price_df = pd.DataFrame([price]).T
    price_df.reset_index(inplace=True)
    
    price_df.columns = ['SKU', 'PRICE']

    merged_df = pd.merge(filtered_df,price_df, on='SKU')
    # *************排列规则********************
    desired_columns = ['SKU', 'DESCRIPTION', 'FABRIC', 'CODE', 'COLOR', 'WK', 'PRICE']
    # 使用loc，选中所有行，和对应的列。列可以是数组
    subset = merged_df.loc[:, desired_columns]

    subset.to_excel('invoice_output.xlsx', index=False)