import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path

filepaths = glob.glob("invoices/*.xlsx")
print(filepaths)

for filepath in filepaths:
    df = pd.read_excel(filepath,sheet_name='Sheet 1') #install openpyxl package
    print(filepath)
    print(df)
    filename = Path(filepath).stem #extract filename from path
    invoice_nr, date = filename.split("-")
    pdf = FPDF(orientation='P', unit='mm', format='A4')
    pdf.add_page()
    pdf.set_font(family='Times',size=12,style='B')
    pdf.cell(w=50,h=8,txt=f'Invoice nr.{invoice_nr}',ln=1)
    pdf.cell(w=50,h=8,txt=f'Date: {date}')

    pdf.output(f"PDFs/{filename}.pdf")