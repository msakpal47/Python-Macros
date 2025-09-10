import os
import pandas as pd
import textwrap

from reportlab.lib.pagesizes import A4
from reportlab.pdfgen import canvas
from reportlab.lib.units import mm
from datetime import datetime
from reportlab.platypus import Paragraph
from reportlab.lib.styles import getSampleStyleSheet

def generate_contract_notes_by_sheet(input_path, base_output_dir):
    # Read all sheets from the Excel file
    xl = pd.read_excel(input_path, sheet_name=None, engine="openpyxl")

    for sheet_name, df in xl.items():
        output_dir = os.path.join(base_output_dir, sheet_name)
        os.makedirs(output_dir, exist_ok=True)

        for index, row in df.iterrows():
            print(row)
            # Extract data with proper formatting
            name = str(row[0]).strip()
            address = str(row[1]).strip()
            trade_date = row[2]
            contract_number = str(row[3]).strip()
            fund_name = str(row[4]).strip()
            shares = "{:,.6f}".format(float(row[5])) if pd.notna(row[5]) else "0.000000"
            nav = "{:,.2f}".format(float(row[6])) if pd.notna(row[6]) else "0.00"
            total_amount = "{:,.2f}".format(float(row[7])) if pd.notna(row[7]) else "0.00"
            contract_note_date = row[8]

            # Prepare file name
            # filename = f"Contract_Note_{contract_number}.pdf"
            filename = f"{sheet_name}_{name.replace(' ', '_')}.pdf"
            filepath = os.path.join(output_dir, filename)

            # Create PDF with proper margins
            c = canvas.Canvas(filepath, pagesize=A4)
            width, height = A4
            left_margin = 20 * mm
            right_margin = width - 20 * mm
            top_margin = height - 25 * mm

            # Header - Centered and properly spaced
            c.setFont("Times-Bold", 12)
            header = f"{sheet_name.capitalize()} Asset Management Incorporated VCC Sub-Fund"
            c.drawCentredString(width/2, top_margin, header)
            
            c.setFont("Times-Roman", 10)
            address_line = f"{address}"
            c.drawCentredString(width/2, top_margin - 20, address_line)

            # Trade information - Right aligned
            c.setFont("Times-Roman", 10)
            c.drawRightString(width - 147, top_margin - 50, "Trade Date:")
            c.drawString(width - 115, top_margin - 50, trade_date.strftime("%d %b %Y"))
            
            c.drawRightString(width - 120, top_margin - 65, "Contract Number:")
            c.drawString(width - 115, top_margin - 65, contract_number)

            # Recipient information - Left aligned with proper line breaks
            c.setFont("Times-Roman", 10)
            c.drawString(left_margin, top_margin - 100, "Issued in the name of:")
            
            c.setFont("Times-Bold", 10)
            # Handle long names/addresses with line breaks

            c.drawString(left_margin, top_margin - 115, name)
            c.drawString(left_margin, top_margin - 130, address)

            # Fund information
            c.setFont("Times-Roman", 10)
            fund_text = "In accordance with your instructions, we confirm having issued the following Units in:"
            fund_lines = [fund_text[i:i+100] for i in range(0, len(fund_text), 100)]
            for i, line in enumerate(fund_lines):
                c.drawString(left_margin, top_margin - 160 - (i * 15), line)
            
            c.setFont("Times-Bold", 10)
            fund_name_lines = [fund_name[i:i+100] for i in range(0, len(fund_name), 100)]
            for i, line in enumerate(fund_name_lines):
                c.drawString(left_margin, top_margin - 175 - (len(fund_lines)-1)*15 - (i * 15), line)

            # Share information in tabular format
            table_start = top_margin - 210 - (len(fund_name_lines)-1*15)
            
            # Draw table headers
            c.setFont("Times-Bold", 10)
            c.drawString(left_margin, table_start, "Particulars")
            c.drawString(left_margin + 80 * mm, table_start, "Details")
            
            # Draw horizontal line
            c.line(left_margin, table_start - 5, right_margin, table_start - 5)
            
            # Table rows
            c.setFont("Times-Roman", 10)
            c.drawString(left_margin, table_start - 25, "Number of shares")
            c.drawString(left_margin + 80 * mm, table_start - 25, shares)
            # c.line(left_margin + 80 * mm, table_start - 27, right_margin, table_start - 27)
            
            c.drawString(left_margin, table_start - 45, "N.A.V")
            c.drawString(left_margin + 80 * mm, table_start - 45, nav)
            # c.line(left_margin + 80 * mm, table_start - 47, right_margin, table_start - 47)
            
            c.setFont("Times-Bold", 10)
            c.drawString(left_margin, table_start - 65, "Total Amount")
            c.drawString(left_margin + 80 * mm, table_start - 65, total_amount)
            # c.line(left_margin + 80 * mm, table_start - 67, right_margin, table_start - 67)

            # Footer note with proper line breaks
                        

            footer_note = (
                "For any discrepancy in the particulars given above; "
                "please email us on ops@apexasset.ai by quoting \n "
                "the Contract Number."
            )
            footer_lines = []
            for para in footer_note.split("\n"):
                wrapped = textwrap.wrap(para, width=99,break_long_words=False,replace_whitespace=False)  # wrap at 100 chars
                footer_lines.extend(wrapped)

            # Draw each line
            for i, line in enumerate(footer_lines):
                c.setFont("Times-Roman", 10)
                c.drawString(left_margin, table_start - 100 - (i * 14), line)

            # footer_lines = [footer_note[i:i+100] for i in range(0, len(footer_note), 100)]
            # for i, line in enumerate(footer_lines):
            #     c.setFont("Times-Roman", 10)
            #     c.drawString(left_margin, table_start - 100 - (i * 14), line)

            # Contract note date
            c.setFont("Times-Roman", 10)
            c.drawString(left_margin, table_start - 130, f"Date of Contract Note: {contract_note_date.strftime("%d %b %Y")}")

            # Footer repetition
            c.setFont("Times-Bold", 10)
            c.drawCentredString(width/2, 40, f"{sheet_name.capitalize()} Asset Management Incorporated VCC Sub-Fund")
            c.setFont("Times-Roman", 11)
            c.drawCentredString(width/2, 25, f"{address}")

            c.save()

        print(f"PDFs generated for sheet '{sheet_name}' in: {output_dir}")

def process_directory(input_dir, output_base_dir):
    supported_extensions = ['.xlsx', '.xls', '.csv']

    for filename in os.listdir(input_dir):
        file_path = os.path.join(input_dir, filename)
        ext = os.path.splitext(filename)[1].lower()

        # Skip temp/hidden Excel files and folders
        if filename.startswith("~$") or not os.path.isfile(file_path):
            print(f"🔸 Skipped: {filename}")
            continue

        if ext in supported_extensions:
            print(f"🟢 Processing: {filename}")
            generate_contract_notes_by_sheet(file_path, output_base_dir)
        else:
            print(f"🔸 Skipped: {filename}")

if __name__ == "__main__":
    input_dir = r".\\"
    output_base_dir = r".\\Output"
    os.makedirs(output_base_dir, exist_ok=True)
    process_directory(input_dir, output_base_dir)
