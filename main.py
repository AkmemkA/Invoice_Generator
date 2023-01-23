import pandas as pd
import glob
from fpdf import FPDF

filepaths = glob.glob("Invoices/*xlsx")

for path in filepaths:
    # Create PDF file, add page
    pdf = FPDF(orientation="P", unit="mm", format="A4")
    pdf.add_page()
    # Write Invoice number and date
    invoice_number = path.split("-")[0].strip("Invoices/")
    date_created = path.split("-")[1].strip("Invoices/").strip(".xlsx")
    pdf.set_font(family="Times", style="B", size=16)
    pdf.cell(w=50, h=8, txt=f"Invoice # {invoice_number}", align="L", ln=1)
    pdf.set_font(family="Times", style="B", size=16)
    pdf.cell(w=50, h=8, txt=f"Date: {date_created}", align="L", ln=1)
    # Create data frame
    df = pd.read_excel(path, sheet_name="Sheet 1")
    # Get titles of each column
    titles = df.columns
    titles = [item.replace('_', ' ').title() for item in titles]
    # Insert titles into a table
    pdf.set_font(family="Times", size=12, style="B")
    pdf.cell(w=30, h=8, txt=titles[0], border=1)
    pdf.cell(w=60, h=8, txt=titles[1], border=1)
    pdf.cell(w=40, h=8, txt=titles[2], border=1)
    pdf.cell(w=30, h=8, txt=titles[3], border=1)
    pdf.cell(w=30, h=8, txt=titles[4], border=1, ln=1)
    # Insert values for each category
    for index, row in df.iterrows():
        pdf.set_font(family="Times", size=12)
        pdf.cell(w=30, h=8, txt=str(row["product_id"]), border=1)
        pdf.cell(w=60, h=8, txt=str(row["product_name"]), border=1)
        pdf.cell(w=40, h=8, txt=str(row["amount_purchased"]), border=1)
        pdf.cell(w=30, h=8, txt=str(row["price_per_unit"]), border=1)
        pdf.cell(w=30, h=8, txt=str(row["total_price"]), border=1, ln=1)

    pdf.output(f"PDFs/Invoice {invoice_number}-{date_created}.pdf")




    """
    # Add footer
    pdf.ln(260)
    pdf.set_font(family="Times", size=10)
    pdf.set_text_color(100, 100, 100)
    pdf.cell(w=0, h=12, txt=row['Topic'], align="R", ln=1)
    # Create lines on each page
    for y in range(20, 290, 10):
        pdf.line(10, y, 200, y)
    # Add empty pages with footers
    for page in range(1, int(row['Pages'])):
        pdf.add_page()
        pdf.ln(272)
        pdf.set_font(family="Times", size=10)
        pdf.set_text_color(100, 100, 100)
        pdf.cell(w=0, h=12, txt=row['Topic'], align="R", ln=1)
        # Create lines on each page
        for y in range(20, 290, 10):
            pdf.line(10, y, 200, y)
            """

