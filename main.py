import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path

filepaths = glob.glob("invoices/*.xlsx")
print(filepaths)


for filepath in filepaths:
    print(filepath)
    filename = Path(filepath).stem #extract filename from path
    invoice_nr, date = filename.split("-")
    pdf = FPDF(orientation='P', unit='mm', format='A4')
    pdf.add_page()
    pdf.set_font(family='Times',size=12,style='B')
    pdf.cell(w=50,h=8,txt=f'Invoice nr.{invoice_nr}', ln=1)
    pdf.cell(w=50,h=8,txt=f'Date: {date}', ln=2)
    df = pd.read_excel(filepath,sheet_name='Sheet 1') #install openpyxl package

    columns = list(df.columns)
    print(columns)
    columns = [item.replace('_',' ').title() for item in columns]

    pdf.set_font(family='Times', size=10,style='B')
    pdf.set_text_color(80, 80, 80)
    pdf.cell(w=30, h=10, txt=f"{columns[0]}", border=1)
    pdf.cell(w=60, h=10, txt=f"{columns[1]}", border=1)
    pdf.cell(w=40, h=10, txt=f"{columns[2]}", border=1)
    pdf.cell(w=30, h=10, txt=f"{columns[3]}", border=1)
    pdf.cell(w=30, h=10, txt=f"{columns[4]}", ln=1, border=1)

    for index, row in df.iterrows():
        pdf.set_font(family='Times', size=10, )
        pdf.set_text_color(80, 80, 80)
        pdf.cell(w=30, h=10, txt=f"{row['product_id']}", border=1)
        pdf.cell(w=60, h=10, txt=f"{row['product_name']}", border=1)
        pdf.cell(w=40, h=10, txt=f"{row['amount_purchased']}", border=1)
        pdf.cell(w=30, h=10, txt=f"{row['price_per_unit']}", border=1)
        pdf.cell(w=30, h=10, txt=f"{row['total_price']}", ln=1, border=1)

# Sum of Total Price Column
    sum_of_total = sum(df['total_price'])

    pdf.set_font(family='Times', size=10, style='B')
    pdf.set_text_color(80, 80, 80)
    pdf.cell(w=30, h=10, txt=f"", border=1)
    pdf.cell(w=60, h=10, txt=f"", border=1)
    pdf.cell(w=40, h=10, txt=f"", border=1)
    pdf.cell(w=30, h=10, txt=f"Total", border=1)
    pdf.cell(w=30, h=10, txt=f"{sum_of_total}", border=1, ln=1)

    pdf.set_font(family='Times', size=12, style='B')
    pdf.cell(w=30, h=10, txt=f"The total price is : {sum_of_total}",  ln=1)
    pdf.set_font(family='Times', size=12, style='B')

    pdf.cell(w=40, h=8, txt=f"PythonHow ")
    pdf.image('pythonhow.png',w=10)


    pdf.output(f"PDFs/{filename}.pdf")

