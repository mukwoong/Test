import pandas as pd
import re

# Define constants for filenames
EXCEL_FILENAME = "selected_data.xlsx"
NUM_FILENAME = 'sum_data.xlsx'

# Function to extract valid bar SKUs
def packed_sku_po():
    df = pd.read_excel('bar-sku.xlsx')
    column_values = df.iloc[:, 0].tolist()
    bar_sku = [value for value in column_values if isinstance(value, str) and re.match(r'^[A-Za-z0-9-]+$', value)]
    return bar_sku

# Function to add data to excel
def add_to_excel(bar_skus):
    sku_new_rows = {}
    selected_data = pd.DataFrame()
    df = pd.read_excel('barcode.xlsx')
    appended_data = []

    for sku in bar_skus:
        data = df[df.iloc[:, 1] == sku].copy()
        data.iloc[:, 4] = data.iloc[:, 4].astype(str)
        data = data.iloc[:-1, [1, 5, 7, 3, 4]]  # Ensure to create a copy
        appended_data.append(data)
        new_rows = len(data)
        sku_new_rows[sku] = new_rows

    appended_data = pd.concat(appended_data, ignore_index=True)
    selected_data = pd.concat([selected_data, appended_data], ignore_index=True)
    return selected_data, sku_new_rows

# Function to process excel
def process_excel(target_value, sku_new_rows):
    start_col = 4
    end_col = start_col + sku_new_rows.get(target_value, 0)
    df = pd.read_excel('bar-sku.xlsx')
    result_rows = []

    matching_row_index = df.index[df.iloc[:, 0] == target_value].tolist()

    if matching_row_index:
        for row_index in matching_row_index:
            result_rows.append(row_index)
            next_row_index = row_index + 1
            while next_row_index < len(df) and pd.isna(df.iloc[next_row_index, 0]):
                result_rows.append(next_row_index)
                next_row_index += 1

    column_totals = []
    for col in range(start_col, end_col):
        total_column = 0
        df.iloc[:, col] = df.iloc[:, col].fillna(0)
        for row_index in result_rows:
            total_column += df.iloc[row_index, col]
        column_totals.append(total_column)

    return column_totals

# Main part of the script
if __name__ == "__main__":
    bar_skus = packed_sku_po()
    code_data, sku_new_rows = add_to_excel(bar_skus)
    
    all_results = pd.DataFrame()
    
    for target_value in bar_skus:
        result_list = process_excel(target_value, sku_new_rows)
        row_result = pd.DataFrame(result_list)
        all_results = pd.concat([all_results, row_result], axis=0, ignore_index=True)
    
    # Write results to excel files
    all_results.to_excel(NUM_FILENAME, startcol=6, index=False, header=True)
    code_data.to_excel(EXCEL_FILENAME, index=False)
