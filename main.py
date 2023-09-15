import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path


file_paths = glob.glob('Invoices/*.xlsx')
print(file_paths)

for file in file_paths:
    df = pd.read_excel(file, sheet_name='Sheet 1')
    pdf = FPDF(orientation="P", unit="mm", format="A4")
    pdf.add_page()

    #get invoice number and date
    file_name = Path(file).stem


    invoice_num, date = file_name.split("-")
    print(invoice_num, date)

    pdf.set_font(family="Times", size=16, style="B")
    pdf.cell(w=50, h=8, txt=f"Invoice nr.{invoice_num}", ln=1)

    #add date
    pdf.set_font(family="Times", size=14, style="B")
    pdf.cell(w=50, h=8, txt=f"Date: {date}", ln=1)

    pdf.output(f"PDFs/{file_name}.pdf")
