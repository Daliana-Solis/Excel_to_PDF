import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path


file_paths = glob.glob('Invoices/*.xlsx')
print(file_paths)

for file in file_paths:
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

    pdf.ln(20)

    #read data
    df = pd.read_excel(file, sheet_name='Sheet 1')

    #extract table column headers and add to PDF and remove _ in each title
    header_col = df.columns
    headers = [colm_name.replace("_", " ").title() for colm_name in header_col]


    pdf.set_font(family="Times", size=10, style="B")
    pdf.set_text_color(80, 80, 80)
    pdf.cell(w=30, h=8, txt=headers[0], border=1)
    pdf.cell(w=50, h=8, txt=headers[1], border=1)
    pdf.cell(w=30, h=8, txt=headers[2], border=1)
    pdf.cell(w=30, h=8, txt=headers[3], border=1)
    pdf.cell(w=30, h=8, txt=headers[4], border=1, ln=1)

    #iterate through every data row and add to PDF
    for index, row in df.iterrows():
        pdf.set_font(family="Times", size=10)
        pdf.set_text_color(80,80,80)

        #columns data
        pdf.cell(w=30, h=8, txt=str(row["product_id"]), border=1)
        pdf.cell(w=50, h=8, txt=str(row["product_name"]),border=1)
        pdf.cell(w=30, h=8, txt=str(row["amount_purchased"]), border=1)
        pdf.cell(w=30, h=8, txt=str(row["price_per_unit"]), border=1)
        pdf.cell(w=30, h=8, txt=str(row["total_price"]), border=1, ln=1)



    pdf.output(f"PDFs/{file_name}.pdf")
