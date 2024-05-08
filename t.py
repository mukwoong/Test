import pandas as pd


df = pd.read_excel("0all-sku.xlsx")
subset = df.iloc[:, 0]
print(subset)