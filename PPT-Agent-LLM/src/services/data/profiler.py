import json
import pandas as pd

def build_data_profile(excel_path: str, sheet_name: str = "Press Release", sample_rows: int = 3) -> dict:
    df = pd.read_excel(excel_path, sheet_name=sheet_name)
    df.columns = df.columns.str.strip()
    
    profile = {
        "columns": list(df.columns),
        "samples": []
    }
    
    for _, row in df.head(sample_rows).iterrows():
        profile["samples"].append({col: str(row[col]) for col in df.columns})
    
    return profile

if __name__ == "__main__":
    import glob
    import os
    
    excel_files = glob.glob("*.xlsx")
    if excel_files:
        excel_path = excel_files[0]
        print(f"Using Excel: {excel_path}")
        profile = build_data_profile(excel_path, "Press Release", sample_rows=3)
        print(json.dumps(profile, indent=2))
    else:
        print("No Excel file found in current directory.")
