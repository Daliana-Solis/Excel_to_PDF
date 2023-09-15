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

    file_name = Path(file).stem
    invoice_num = file_name.split("-")[0]
    print(invoice_num)
    pdf.set_font(family="Times", size=16, style="B")
    pdf.cell(w=50, h=8, txt=f"Invoice nr.{invoice_num}")
    pdf.output(f"PDFs/{file_name}.pdf")
