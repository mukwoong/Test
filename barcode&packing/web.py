import pandas as pd
import re

# Function to get SKUs from a given file
def get_sku(file_name):
    df = pd.read_excel(file_name)
    column_values = df.iloc[:, 0].tolist()
    sku = [value for value in column_values if isinstance(value, str) and re.match(r'^[A-Za-z0-9-]+$', value)]

    print("sku:", sku)
    return sku

# Function to get web data
def get_web(web_file, all_sku_df):
    web = []
    colors = []
    skus = get_sku(web_file)
    for sku in skus:
        matching_rows = all_sku_df.index[all_sku_df.iloc[:, 0].str.upper() == sku.upper()]

        if len(matching_rows) > 0:
            index = matching_rows[0]
            color_ = all_sku_df.iloc[index, 3]
            web_index = index + 1
            web_num = all_sku_df.iloc[web_index, 17]

            colors.append(color_)
            web.append(web_num)
        else:
            web.append("no such web.")
            colors.append("")  # Add a placeholder for color

    return web, colors

# Function to get PO data
def get_po(po_file, all_sku_df):
    po = []
    colors = []
    skus = get_sku(po_file)
    for sku in skus:
        matching_rows = all_sku_df.index[all_sku_df.iloc[:, 0].str.upper() == sku.upper()]

        if len(matching_rows) > 0:
            index = matching_rows[0]
            po_num = all_sku_df.iloc[index, 17]
            color = all_sku_df.iloc[index, 3]
            po.append(po_num)
            colors.append(color)
        else:
            po.append("no such po.")
            colors.append("")  # Add a placeholder for color

    return po, colors

# Main function
def main():
    file1 = 'web.xlsx'
    file2 = 'po.xlsx'
    all_sku_file = 'all-sku.xlsx'

    # Read all-sku file
    df_all_sku = pd.read_excel(all_sku_file)
    df_all_sku.iloc[:, 0] = df_all_sku.iloc[:, 0].fillna('')
    df_all_sku.iloc[:, 0] = df_all_sku.iloc[:, 0].str.strip()

    # Get web and PO data
    result_web = get_web(file1, df_all_sku)
    result_po = get_po(file2, df_all_sku)

    # Determine the maximum length of results
    max_length = max(len(result_web[0]), len(result_po[0]))

    # Pad results to ensure they have the same length
    result_web_padded = [result_web[0] + [''] * (max_length - len(result_web[0])), 
                         result_web[1] + [''] * (max_length - len(result_web[1]))]
    result_po_padded = [result_po[0] + [''] * (max_length - len(result_po[0])), 
                        result_po[1] + [''] * (max_length - len(result_po[1]))]

    # Combine results into a DataFrame
    df_result = pd.DataFrame({
        'Web': result_web_padded[0],
        'Web SKU': get_sku(file1) + [''] * (max_length - len(get_sku(file1))),
        'Web Color': result_web_padded[1],
        'PO': result_po_padded[0],
        'PO SKU': get_sku(file2) + [''] * (max_length - len(get_sku(file2))),
        'PO Color': result_po_padded[1]
    })

    # Write the results to a single Excel file
    df_result.to_excel('result_combined.xlsx', index=False)

if __name__ == "__main__":
    main()
