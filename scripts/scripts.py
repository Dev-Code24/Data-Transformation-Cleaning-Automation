import pandas as pd
from openpyxl import load_workbook
from openpyxl.styles import PatternFill
import math

def extract_data(file_path, file_type='excel'):
    """Extracts data from an Excel (.xlsx) file."""
    if file_type == 'excel':
        return pd.read_excel(file_path)
    else:
        raise ValueError("Unsupported file type")


def load_data(df, output_path):
    """Saves the transformed data to an Excel (.xlsx) file."""
    df.to_excel(output_path, index=False)

def transform_data(df):
    df.columns = [col.lower().replace(" ", "_") for col in df.columns]
    df.drop_duplicates(inplace=True)
    
    df['age'].fillna(math.ceil(df['age'].mean()), inplace=True)
    df['score'].fillna(df['score'].mean(), inplace=True)
    return df


def generate_report(df, output_path, threshold_value=50.0):

    with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='Report', index=False)
    
    # Load the workbook and select the sheet to apply formatting
    wb = load_workbook(output_path)
    ws = wb['Report']
      
    red_fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
    
    columns_to_check = ["age", "score", "purchase_amount"]  
    
    col_indexes = [cell.column for cell in ws[1] if cell.value in columns_to_check]
    
    for row in ws.iter_rows(min_row=2, max_row=ws.max_row):
        for col_index in col_indexes:
            cell = row[col_index - 1]  # Adjust for zero-based index
            if isinstance(cell.value, (int, float)) and cell.value < threshold_value:
                cell.fill = red_fill
    
    wb.save(output_path)
    print("Report generated with conditional formatting applied to selected columns.")
