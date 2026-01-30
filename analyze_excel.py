import pandas as pd
import os
import sys

# Redirect stdout to a file
sys.stdout = open('analysis_result.txt', 'w', encoding='utf-8')

file_path = 'Target_08월 월간진도보고서.xlsx'
if not os.path.exists(file_path):
    print(f"File not found: {file_path}")
else:
    try:
        # Read the Excel file to understand sheets and columns
        xl = pd.ExcelFile(file_path)
        print(f"Sheet names: {xl.sheet_names}")
        
        for sheet in xl.sheet_names:
            print(f"\n--- Sheet: {sheet} ---")
            try:
                df = pd.read_excel(file_path, sheet_name=sheet, nrows=5)
                print("Columns:")
                print(df.columns.tolist())
                print("First 5 rows:")
                print(df.head())
            except Exception as e_sheet:
                 print(f"Error reading sheet {sheet}: {e_sheet}")

    except Exception as e:
        print(f"Error reading excel: {e}")

sys.stdout.close()
