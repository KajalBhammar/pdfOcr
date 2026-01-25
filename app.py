import streamlit as st
import base64
import fitz  # PyMuPDF
from openpyxl import Workbook
import os
from datetime import datetime
from mistralai import Mistral
import tempfile
from dotenv import load_dotenv

# Load environment variables
load_dotenv()

# Initialize Mistral client
api_key = os.getenv("MISTRAL_API_KEY")
if not api_key:
    st.error("❌ MISTRAL_API_KEY not found in .env file")
    st.stop()

client = Mistral(api_key=api_key)

# Facility mapping list
facility_list = [
    "Alliance Health at Marina Bay, 2 Seaport, Quincy, MA",
    "Alliance Health at West Acres, 804 Pleasant St, Brockton, MA",
    "Sherrill House, 135 S Huntington Ave, Jamaica Plain, MA",
    "Alliance Health at Maples, 90 Taunton St, Wrentham, MA 02097",
    "Oak Knoll, 9 Ambetter Dr, Framingham, MA",
    "Sippican Rehab & Healthcare, 15 Mill St, Marion, MA",
    "Alliance Health at Braintree, 175 Grove St, Braintree, MA",
    "Harrington House Rehab & Healthcare, 160 Main St, Walpole, MA",
    "Bethany Healthcare Rest Home, 97 Bethany Rd, Framingham, MA",
    "Alliance Health at Marie Esther, 720 Boston, Marlborough, MA",
    "Shrewsbury Nursing & Rehab Center, 40 Julio Dr, Shrewsbury, MA 01545",
    "Alliance Health at Doolittle Unit 1, 16 Bird St, Foxboro, MA",
    "CareOne at Concord, 57 Old Rd to 9 Acre Corner, Concord, MA 01742",
    "The Commons at Lincoln, 3 Harvest Cir, Lincoln, MA",
    "Rivercrest Nursing and Wellness, 100 Newbury Ct, Concord, MA",
    "Brookhaven at Lexington Independent Living, 1010 Waltham St, Lexington, MA",
    "Woburn Rehabilitation & Nursing Center, 18 Frances St, Woburn, MA 01801",
    "CareOne at Wilmington, 750 Woburn St, Wilmington, MA 01887",
    "CareOne at Lexington, 178 Lowell St, Lexington, MA 02420",
    "Winchester Rehabilitation and Nursing Center, 223 Swanton St, Winchester, MA 01890",
    "Aberjona Rehabilitation & Nursing Center, 184 Swanton St, Winchester, MA 01890",
    "The Commons in Lincoln, 1 Harvest Cir, Lincoln, MA 01773",
    "CareOne at Essex Park, 265 Essex St, Beverly, MA 01915",
    "CareOne at Peabody, 199 Andover St, Peabody, MA 01960"
]

def match_facility(extracted_facility):
    """Match extracted facility against the facility list"""
    if not extracted_facility or extracted_facility.lower() == "not found":
        return "Not found"
    
    extracted_lower = extracted_facility.lower()
    
    for facility in facility_list:
        facility_lower = facility.lower()
        facility_parts = [part.strip() for part in facility_lower.split(',')]
        facility_name = facility_parts[0]
        
        if facility_name in extracted_lower:
            return facility
    
    for facility in facility_list:
        facility_lower = facility.lower()
        facility_parts = facility_lower.split(',')
        for part in facility_parts:
            part = part.strip()
            if len(part) > 5 and part in extracted_lower:
                return facility
    
    return "Not found"

def extract_from_pdf(pdf_file, progress_bar, status_text):
    """Extract data from PDF and return Excel file"""
    # Save PDF to temp file
    with tempfile.NamedTemporaryFile(delete=False, suffix='.pdf') as tmp:
        tmp.write(pdf_file.read())
        pdf_path = tmp.name
    
    try:
        # Convert PDF to images
        doc = fitz.open(pdf_path)
        total_pages = len(doc)
        images = [doc[page_num].get_pixmap(matrix=fitz.Matrix(2, 2)) for page_num in range(total_pages)]
        
        # Create Excel workbook
        wb = Workbook()
        ws = wb.active
        ws.title = "Extracted Data"
        
        # Add headers
        ws['A1'] = "Appt Date"
        ws['B1'] = "Name of Patient"
        ws['C1'] = "Facility Information_name_pdf"
        ws['D1'] = "Facility_name_final"
        
        row = 2
        
        # Process each page
        for page_num, pixmap in enumerate(images, 1):
            # Update progress
            progress = page_num / total_pages
            progress_bar.progress(progress)
            status_text.text(f"Processing page {page_num} of {total_pages}...")
            
            # Convert pixmap to PNG bytes
            png_bytes = pixmap.tobytes("png")
            image_base64 = base64.b64encode(png_bytes).decode("utf-8")
            
            # Extract data using Mistral
            inputs = [
                {
                    "role": "user",
                    "content": [
                        {"type": "text", "text": "Extract ONLY PRINTED/TYPED text from this document. Skip all handwritten text completely.\n\nExtract the following information:\n1. Appointment Date (Appt Date)\n2. Name of Patient\n3. Facility Information (Include BOTH facility name AND address together)\n\nIMPORTANT FOR PATIENT NAME: Copy the name EXACTLY as it appears in the PDF. Do NOT rearrange, reformat, or change the order. If it says 'DAVID FREITAS W', write 'DAVID FREITAS W' - do NOT change it to 'FREITAS, DAVID'. Extract it as-is without any modifications.\n\nReturn the data in this format:\nAppt Date: [date]\nName of Patient: [name - COPY EXACTLY AS APPEARS IN PDF, DO NOT CHANGE FORMAT]\nFacility Information: [facility name and address]\n\nIf any field is not found or is handwritten, write 'Not found'.\nFor Facility Information, extract and combine the facility name with its complete address on the same line.\nReturn all found printed text, do NOT skip pages."},
                        {
                            "type": "image_url",
                            "image_url": f"data:image/png;base64,{image_base64}"
                        }
                    ]
                }
            ]
            
            response = client.beta.conversations.start(
                agent_id="ag_019bef914c89717aa9731bda4c9d23df",
                inputs=inputs
            )
            
            # Parse response
            if hasattr(response.outputs[0], 'content'):
                response_text = response.outputs[0].content
            elif hasattr(response.outputs[0], 'text'):
                response_text = response.outputs[0].text
            elif hasattr(response.outputs[0], 'message') and hasattr(response.outputs[0].message, 'content'):
                response_text = response.outputs[0].message.content
            else:
                response_text = str(response.outputs[0])
            
            # Extract fields from response
            lines = response_text.split('\n')
            appt_date = ""
            patient_name = ""
            facility_info = ""
            
            for line in lines:
                if "Appt Date:" in line:
                    appt_date = line.split("Appt Date:")[-1].strip().replace("*", "")
                elif "Name of Patient:" in line:
                    patient_name = line.split("Name of Patient:")[-1].strip().replace("*", "")
                elif "Facility Information:" in line:
                    facility_info = line.split("Facility Information:")[-1].strip().replace("*", "")
            
            # Only add to Excel if at least one field is found
            has_patient_info = patient_name and patient_name.lower() != "not found"
            has_facility_info = facility_info and facility_info.lower() != "not found"
            
            if has_patient_info or has_facility_info:
                facility_final = match_facility(facility_info)
                
                ws[f'A{row}'] = appt_date
                ws[f'B{row}'] = patient_name
                ws[f'C{row}'] = facility_info
                ws[f'D{row}'] = facility_final
                row += 1
        
        # Save Excel file
        pdf_filename = os.path.splitext(pdf_file.name)[0]
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_file = f"{pdf_filename}_{timestamp}.xlsx"
        wb.save(output_file)
        
        # Close PDF
        doc.close()
        
        return output_file
    
    finally:
        # Clean up temp file
        os.unlink(pdf_path)

# Streamlit UI
st.set_page_config(page_title="PDF Data Extractor", layout="centered")
st.title("📄 PDF Data Extractor")

# File uploader
uploaded_file = st.file_uploader("Upload PDF", type="pdf")

if uploaded_file is not None:
    st.success(f"✅ File uploaded: {uploaded_file.name}")
    
    if st.button("🔄 Process PDF", use_container_width=True):
        # Progress bar and status
        progress_bar = st.progress(0)
        status_text = st.empty()
        
        try:
            # Process PDF
            output_file = extract_from_pdf(uploaded_file, progress_bar, status_text)
            
            # Complete
            progress_bar.progress(1.0)
            status_text.text("✅ Processing complete!")
            
            # Download button
            with open(output_file, "rb") as file:
                st.download_button(
                    label="📥 Download Excel File",
                    data=file,
                    file_name=os.path.basename(output_file),
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
            
            # Clean up
            os.unlink(output_file)
        
        except Exception as e:
            st.error(f"❌ Error: {str(e)}")
