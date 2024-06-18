import pandas as pd
import glob
from fpdf import FPDF
from pathlib import Path

filepaths = glob.glob("Invoices/*.xlsx")

for filepath in filepaths:
    filename = Path(filepath).stem
    invoice_nr, date = filename.split("-")

    pdf = FPDF(orientation="P", unit="mm", format="a4")
    pdf.add_page()

    pdf.set_font(family="Times", size=16, style="B")
    pdf.cell(w=10, h=8, txt=f"Invoice nr.{invoice_nr}", ln=1)

    pdf.set_font(family="Times", size=16, style="B", )
    pdf.cell(w=10, h=8, txt=f"Date{date}", ln=1)

    df = pd.read_excel(filepath, sheet_name="Sheet 1")
    column = df.columns
    heading = [item.replace("_", " ").title() for item in column]
    pdf.set_font(family="Times", size=10, style="b")
    pdf.cell(w=30, h=8, txt=heading[0], border=1)
    pdf.cell(w=45, h=8, txt=heading[1], border=1)
    pdf.cell(w=35, h=8, txt=heading[2], border=1)
    pdf.cell(w=30, h=8, txt=heading[3], border=1)
    pdf.cell(w=20, h=8, txt=heading[4], border=1, ln=1)

    for index, row in df.iterrows():
        pdf.set_font(family="Times", size=10)
        pdf.cell(w=30, h=8, txt=str(row["product_id"]), border=1)
        pdf.cell(w=45, h=8, txt=str(row["product_name"]),border=1)
        pdf.cell(w=35, h=8, txt=str(row["amount_purchased"]),border=1)
        pdf.cell(w=30, h=8, txt=str(row["price_per_unit"]),border=1)
        pdf.cell(w=20, h=8, txt=str(row["total_price"]),border=1, ln=1)

    pdf.output(f"PDFs/{filename}.pdf")
