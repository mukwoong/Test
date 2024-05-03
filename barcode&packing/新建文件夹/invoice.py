import pandas as pd
import re

# Function to get SKU from a given file
def get_sku(file_name):
    df = pd.read_excel(file_name)
    column_values = df.iloc[:, 0].tolist()
    sku = [value for value in column_values if isinstance(value, str) and re.match(r'^[A-Za-z0-9-]+$', value)]
    return sku

# Function to get price for a given SKU
def get_price(sku):
    df_all_sku = pd.read_excel('all-sku.xlsx')
    df_all_sku.iloc[:, 0] = df_all_sku.iloc[:, 0].fillna('')
    df_all_sku.iloc[:, 0] = df_all_sku.iloc[:, 0].str.strip()
    matching_rows = df_all_sku.index[df_all_sku.iloc[:, 0].str.upper() == sku.upper()]

    if len(matching_rows) > 0:
        price = df_all_sku.iloc[matching_rows[0], 13]
        return price
    else:
        return ""  # Return a placeholder for price
    
def process_excel(target_value, file_name):
    start_col = 4
    end_col = 12

    # Read Excel file
    df = pd.read_excel(file_name)
    result_rows = []

    # Locate the row where Column1 value matches the target value
    matching_row_index = df.index[df.iloc[:, 0] == target_value].tolist()

    if matching_row_index:
        # If there's a match, iterate through subsequent rows
        for row_index in matching_row_index:
            result_rows.append(row_index)
            # Check subsequent rows for 'None'
            next_row_index = row_index + 1
            while next_row_index < len(df) and pd.isna(df.iloc[next_row_index, 0]):
                result_rows.append(next_row_index)
                next_row_index += 1

    # Process columns start_col to end_col and accumulate totals
    column_totals = []
    for col in range(start_col - 1, end_col):  # Convert to zero-based indexing
        total_column = 0
        # Fill NaN values with 0 in the current column
        df.iloc[:, col] = df.iloc[:, col].fillna(0)
        # Accumulate values from the current column
        for row_index in result_rows:
            total_column += df.iloc[row_index, col]
        column_totals.append(total_column)

    return column_totals

def process_target_values(target_values, file_name):
    target_totals = {}  # Dictionary to store column totals for each target value
    for target_value in target_values:
        result_list = process_excel(target_value, file_name)
        target_totals[target_value] = result_list  # Store column totals for each target value

    return target_totals

def merge_target_totals(target_totals1, target_totals2):
    merged_totals = {}
    all_targets = set(list(target_totals1.keys()) + list(target_totals2.keys()))

    for target in all_targets:
        totals1 = target_totals1.get(target, [])
        totals2 = target_totals2.get(target, [])
        merged_totals[target] = [sum(pair) for pair in zip(totals1, totals2)]

    return merged_totals

def main():
    file1 = 'po.xlsx'
    file2 = 'web.xlsx'

    # Get SKUs from both files
    skus = get_sku(file1)
    skus.extend(get_sku(file2))

    output_data = []
    unique_skus = set()  # Track unique SKUs

    # Read data from 'hs-code.xlsx'
    df = pd.read_excel('hs-code.xlsx')
    df.iloc[:, 1] = df.iloc[:, 1].fillna('')
    df.iloc[:, 1] = df.iloc[:, 1].str.strip()

    # Iterate through SKUs and retrieve corresponding information
    for sku in skus:
        target_value = sku.strip().upper()
        matching_rows = df.index[df.iloc[:, 1].str.upper() == target_value]
        
        if len(matching_rows) > 0:
            index = matching_rows[0]
            style = df.iloc[index, 3]
            color = df.iloc[index, 4]
            fabric = df.iloc[index, 6]
            wk = df.iloc[index, 7]
            code = str(df.iloc[index, 8])  # Convert code to string
            
            if sku not in unique_skus:  # Check if SKU is unique
                price = get_price(sku)
                print(sku, style, fabric, code, color, wk, price, sep=",")  # Include price in print statement
                output_data.append([sku, style, fabric, code, color, wk, price])
                unique_skus.add(sku)  # Add SKU to set of unique SKUs

    # Get merged totals from process_excel() in code2
    target_values1 = set(get_sku(file1))
    target_values2 = set(get_sku(file2))

    target_totals1 = process_target_values(target_values1, file1)
    target_totals2 = process_target_values(target_values2, file2)
    merged_totals = merge_target_totals(target_totals1, target_totals2)

    # Add merged totals to output data
    for sku_row in output_data:
        sku = sku_row[0]
        merged_total = merged_totals.get(sku, [])
        sku_row.append(merged_total)

    # Write the output data to a new Excel file
    output_df = pd.DataFrame(output_data, columns=["SKU", "Style", "Fabric", "Code", "Color", "WK", "Price", "Merged Totals"])
    output_df.to_excel("invoice_output.xlsx", index=False)

if __name__ == "__main__":
    main()
