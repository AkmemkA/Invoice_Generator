import pandas as pd
import glob
from fpdf import FPDF

filepaths = glob.glob("Invoices/*xlsx")

for path in filepaths:
    df = pd.read_excel(path, sheet_name="Sheet 1")

    # Create PDF file, add page
    pdf = FPDF(orientation="P", unit="mm", format="A4")
    pdf.add_page()
    pdf.set_font(family="Times", style="B", size=16)
    pdf.set_text_color(0, 0, 0)
    invoice_number = path.split("-")[0].strip("Invoices/")
    date_created = path.split("-")[1].strip("Invoices/").strip(".xlsx")

    # Insert Date
    pdf.cell(w=50, h=8, txt=f"Invoice # {invoice_number}", align="L", ln=1)
    pdf.set_font(family="Times", size=16, style="B")
    pdf.cell(w=50, h=8, txt=f"Date: {date_created}", align="L", ln=1)

    # Add a header
    columns_ids = df.columns
    columns_ids = [item.replace("_", " ").title() for item in columns_ids]
    pdf.set_font(family="Times", size=10, style="B")
    pdf.cell(w=30, h=8, txt=columns_ids[0], border=1)
    pdf.cell(w=60, h=8, txt=columns_ids[1], border=1)
    pdf.cell(w=40, h=8, txt=columns_ids[2], border=1)
    pdf.cell(w=30, h=8, txt=columns_ids[3], border=1)
    pdf.cell(w=30, h=8, txt=columns_ids[4], border=1, ln=1)

    # Add rows with data
    for index, row in df.iterrows():
        pdf.set_font(family="Times", size=10)
        pdf.cell(w=30, h=8, txt=str(row["product_id"]), border=1)
        pdf.cell(w=60, h=8, txt=str(row["product_name"]), border=1)
        pdf.cell(w=40, h=8, txt=str(row["amount_purchased"]), border=1)
        pdf.cell(w=30, h=8, txt=str(row["price_per_unit"]), border=1)
        pdf.cell(w=30, h=8, txt=str(row["total_price"]), border=1, ln=1)

    # Calculate sum of totals and add to the table
    sum_totals = df["total_price"].sum()
    pdf.set_font(family="Times", size=10)
    pdf.cell(w=30, h=8, txt="", border=1)
    pdf.cell(w=60, h=8, txt="", border=1)
    pdf.cell(w=40, h=8, txt="", border=1)
    pdf.cell(w=30, h=8, txt="", border=1)
    pdf.cell(w=30, h=8, txt=str(sum_totals), border=1, ln=1)

    # Add summary under the table
    pdf.set_font(family="Times", size=14, style="B")
    pdf.cell(w=30, h=8, txt="", ln=1)
    pdf.cell(w=30, h=8, txt=f"The total of the invoice is: {str(sum_totals)} euro.", ln=1)



    pdf.output(f"PDFs/Invoice {invoice_number}-{date_created}.pdf")


