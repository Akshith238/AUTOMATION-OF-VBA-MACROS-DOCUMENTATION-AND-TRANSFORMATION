from fpdf import FPDF

def generate_pdf(pdf_data, filename):
    pdf = FPDF()
    pdf.set_auto_page_break(auto=True, margin=15)
    
    # Add title page
    pdf.add_page()
    pdf.set_font("Arial", 'B', size=16)
    pdf.cell(200, 10, txt="Automated Macro Analysis Report", ln=True, align='C')
    pdf.ln(10)
    pdf.set_font("Arial", size=12)
    pdf.cell(200, 10, txt="Generated Report", ln=True, align='C')
    pdf.ln(10)

    # Add content
    pdf.multi_cell(0, 10, pdf_data)
    
    # Generate the PDF file with the specified filename
    pdf_path = f"{filename}.pdf"
    pdf.output(pdf_path)
    
    return pdf_path
