import streamlit as st
import pandas as pd
from reportlab.lib.pagesizes import letter
from reportlab.lib import colors
from reportlab.pdfgen import canvas
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.units import inch
from reportlab.platypus import Paragraph

# Load candidates
candidates_file = 'candidates.xlsx'  # Update this path as necessary
candidates_df = pd.read_excel(candidates_file)

# Check for leading/trailing spaces in column names
candidates_df.columns = candidates_df.columns.str.strip()

# Function to generate certificate
def generate_certificate(name, start_date, end_date, issue_date):
    pdf_file = f"Internship_Certificate_{name}.pdf"
    c = canvas.Canvas(pdf_file, pagesize=letter)
    width, height = letter


    # Add logos at the top center
    c.drawImage('Certificate Batch Untitled.png', 450, height - 100, 65, 50)  # Adjust the path and size as necessary
    
    # Add logos at the top center
    c.drawImage('image.png', 250, height - 120, 120, 60)  # Adjust the path and size as necessary

    # Add border
    c.setStrokeColor(colors.navy)
    c.setLineWidth(2)
    c.rect(0.5 * inch, 0.5 * inch, width - inch, height - inch)
    c.setStrokeColor(colors.plum)
    c.setLineWidth(1.5)
    c.rect(0.55 * inch, 0.55 * inch, width - 1.1 * inch, height - 1.1 * inch)

    # Add text
    styles = getSampleStyleSheet()
    c.setFont("Helvetica-Bold", 24)
    c.setFillColor(colors.navy)
    c.drawCentredString(width / 2.0, height - 150, "CERTIFICATE OF INTERNSHIP")

    c.setFont("Helvetica-Bold", 20)
    c.drawCentredString(width / 2.0, height - 180, "Financial Analyst")

    c.setFont("Helvetica", 16)
    c.drawCentredString(width / 2.0, height - 210, "This certifies that")

    c.setFont("Helvetica-Bold", 20)
    c.drawCentredString(width / 2.0, height - 240, name)

    c.setFont("Helvetica", 20)
    text = f"has successfully completed the Financial Analyst Internship program at PredictRAM\nfrom {start_date.strftime('%d-%m-%Y')} to {end_date.strftime('%d-%m-%Y')}."
    
    # Create Paragraph for text wrapping
    p = Paragraph(text, style=styles['Normal'])
    text_width = width - 1.5 * inch
    p.width = text_width
    p.height = 50  # Adjust height as needed
    p.wrapOn(c, text_width, 50)
    
    # Draw text in PDF
    text_y = height - 270
    p.drawOn(c, (width - text_width) / 2.0, text_y)
    text_y -= p.height

    c.setFont("Helvetica-Bold", 12)
    c.drawString(0.75 * inch, text_y, "Key Responsibilities and Achievements:")

    responsibilities = [
        "Conducted in-depth fundamental and technical analysis of stocks.",
        "Tracked and recorded market data, preparing forecasts on financial and economic events.",
        "Provided insights on upcoming economic events and trends.",
        "Mastered advanced software for predictive analysis.",
        "Developed research reports on national economic conditions and financial forecasts.",
        "Contributed to secondary financial research, enhancing team outputs."
    ]
    y = text_y - 20
    c.setFont("Helvetica", 11)
    for responsibility in responsibilities:
        c.drawString(0.75 * inch, y, f"- {responsibility}")
        y -= 25

    c.drawString(0.75 * inch, y, "Performance Summary:")
    y -= 40
    performance_summary = f"{name} demonstrated strong analytical skills, effectively contributed to team projects,\nand delivered valuable insights that supported the company’s objectives."
    
    # Create Paragraph for performance summary
    p = Paragraph(performance_summary, style=styles['Normal'])
    p.width = text_width
    p.height = 50  # Adjust height as needed
    p.wrapOn(c, text_width, 50)
    p.drawOn(c, 0.75 * inch, y)
    y -= p.height

    c.drawCentredString(width / 2.0, y, f"Issue Date: {issue_date.strftime('%d-%m-%Y')}")
    y -= 40

    # Footer with images and names
    c.drawImage('signature1.png', 0.75 * inch, y, 100, 60)  # Replace with actual path
    c.drawImage('signature2.png', width - 2.25 * inch, y, 120, 60)  # Replace with actual path
    y -= 40
    c.drawString(0.75 * inch, y, "Subir Singh")
    c.drawString(0.75 * inch, y - 15, "Director")
    c.drawString(width - 1.75 * inch, y, "Sheetal Maurya")
    c.drawString(width - 1.75 * inch, y - 15, "Asst. Prof")
    y -= 40
    c.drawImage('Supported By1.png', (width - 320) / 2.0, y, 320, 60)  # Replace with actual path

    c.showPage()
    c.save()
    return pdf_file

# Streamlit App
# Add a logo to the top header
st.image("png_2.3-removebg-preview.png", width=400)  # Replace "your_logo.png" with the path to your logo
st.title("Internship Certificate Generator")

# Input Form
st.header("Enter Internship Details")

with st.form("internship_form"):
    name = st.text_input("Candidate Name")
    start_date = st.date_input("Internship Start Date")
    end_date = st.date_input("Internship End Date")
    issue_date = st.date_input("Issue Date")
    submit_button = st.form_submit_button(label="Generate Certificate")

# Check if the candidate is in the list
if submit_button:
    if 'Candidate Name' in candidates_df.columns:
        if name in candidates_df['Candidate Name'].values:
            # Generate PDF Certificate
            pdf_file = generate_certificate(name, start_date, end_date, issue_date)
            
            with open(pdf_file, "rb") as file:
                st.download_button(
                    label="Download Certificate",
                    data=file,
                    file_name=pdf_file,
                    mime="application/pdf"
                )
        else:
            st.error("Candidate not found in the list.")
    else:
        st.error("'Candidate Name' column not found in the candidates file.")
