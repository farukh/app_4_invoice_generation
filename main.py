import pandas as pd
import glob

filepaths = glob.glob("invoices/*.xlsx")
print(filepaths)

for file in filepaths:
    df = pd.read_excel(file,sheet_name='Sheet 1') #install openpyxl package
    print(file)
    print(df)