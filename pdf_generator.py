from fpdf import FPDF

def generate_pdf(content):
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Arial", size=12)
    
    # Split content into lines
    lines = content.split('\n')
    
    for line in lines:
        if line.startswith('# '):
            pdf.set_font("Arial", 'B', 16)
            pdf.cell(200, 10, txt=line[2:], ln=True)
        elif line.startswith('## '):
            pdf.set_font("Arial", 'B', 14)
            pdf.cell(200, 10, txt=line[3:], ln=True)
        else:
            pdf.set_font("Arial", size=12)
            pdf.multi_cell(0, 5, txt=line)
    
    output_path = "vba_documentation.pdf"
    pdf.output(output_path)
    return output_path