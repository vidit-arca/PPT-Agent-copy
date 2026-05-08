import pandas as pd
from src.core.agent import read_excel_files_from_directory

df = read_excel_files_from_directory("/Users/apple/Desktop/PPT-Agent copy/2026-03-02_to_2026-03-08")
print(df[df["SourceFile"] == "Listed Companies"][["SubCategory", "PDF_URL", "Subdomain"]])
