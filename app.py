import streamlit as st
import base64
import fitz  # PyMuPDF
from openpyxl import Workbook
from openpyxl.styles import PatternFill
import os
from datetime import datetime
from mistralai import Mistral
import tempfile
from datetime import datetime as dt

# Initialize Mistral client with Streamlit secrets or environment variable
try:
    api_key = st.secrets["mistral_api_key"]
except KeyError:
    api_key = os.environ.get("MISTRAL_API_KEY")
    if not api_key:
        st.error("❌ API Key not found. Please set 'mistral_api_key' in Streamlit secrets or MISTRAL_API_KEY environment variable.")
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
        return ""
    
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
    
    return ""

def format_date_to_ddmmyyyy(date_str, default_date="", is_required=False):
    """Convert date string to dd/mm/yyyy format"""
    if not date_str or date_str.lower() == "not found":
        # Return empty string for all cases
        return ""
    
    date_str = date_str.strip()
    
    # Common date formats to try
    date_formats = [
        "%m/%d/%Y", "%m/%d/%y", "%m-%d-%Y", "%m-%d-%y",
        "%d/%m/%Y", "%d/%m/%y", "%d-%m-%Y", "%d-%m-%y",
        "%Y/%m/%d", "%Y-%m-%d",
        "%B %d, %Y", "%b %d, %Y",
        "%d %B %Y", "%d %b %Y",
        "%m/%d", "%d/%m"
    ]
    
    for fmt in date_formats:
        try:
            parsed_date = dt.strptime(date_str, fmt)
            return parsed_date.strftime("%d/%m/%Y")
        except ValueError:
            continue
    
    # If no format matches, return empty string
    return ""

def extract_from_pdf(pdf_file, progress_bar, status_text):
    """Extract data from PDF and return Excel file"""
    # Save PDF to temp file
    with tempfile.NamedTemporaryFile(delete=False, suffix='.pdf') as tmp:
        tmp.write(pdf_file.read())
        pdf_path = tmp.name
    
    doc = None
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
        ws['A1'] = "Name of the Phleb"
        ws['B1'] = "Date"
        ws['C1'] = "No of Patient"
        ws['D1'] = "Patient ID"
        ws['E1'] = "patient bod"
        ws['F1'] = "Name of Patient"
        ws['G1'] = "Patient Birthdate"
        ws['H1'] = "Facility  Information"
        ws['I1'] = "Patients ICD Code"
        ws['J1'] = "From"
        ws['K1'] = "To"
        ws['L1'] = "Miles"
        ws['M1'] = "To"
        ws['N1'] = "Miles"
        ws['O1'] = "To"
        ws['P1'] = "Miles"
        ws['Q1'] = "To"
        ws['R1'] = "Miles"
        ws['S1'] = "To"
        ws['T1'] = "Miles"
        ws['U1'] = "Total Miles"
        ws['V1'] = "Insurance Company"
        ws['W1'] = "Mem ID"
        ws['X1'] = "Group Mem ID"
        
        row = 2
        prev_date = ""
        prev_facility = ""
        
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
                        {"type": "text", "text": "Extract ONLY PRINTED/TYPED text from this document. Skip all handwritten text completely.\n\nExtract the following information:\n1. Appointment Date (Appt Date)\n2. Name of Patient\n3. Patient Birthdate (DOB)\n4. Facility Information (Include BOTH facility name AND address together)\n5. Primary Insurance Company\n6. Sub/Member No. (Member ID)\n\nIMPORTANT FOR PATIENT NAME: Copy the name EXACTLY as it appears in the PDF. Do NOT rearrange, reformat, or change the order. If it says 'DAVID FREITAS W', write 'DAVID FREITAS W' - do NOT change it to 'FREITAS, DAVID'. Extract it as-is without any modifications.\n\nReturn the data in this format:\nAppt Date: [date]\nName of Patient: [name - COPY EXACTLY AS APPEARS IN PDF, DO NOT CHANGE FORMAT]\nPatient Birthdate: [DOB in any date format found]\nFacility Information: [facility name and address]\nPrimary: [Insurance Company name]\nSub/Member No.: [Member ID]\n\nIf any field is not found or is handwritten, leave it blank.\nFor Facility Information, extract and combine the facility name with its complete address on the same line.\nReturn all found printed text, do NOT skip pages."},
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
            
            # Skip this page if it contains "DROP SHEET"
            if "DROP SHEET" in response_text.upper():
                # Write a completely blank row for DROP SHEET pages
                ws[f'A{row}'] = ""
                ws[f'B{row}'] = ""
                ws[f'C{row}'] = ""
                ws[f'D{row}'] = ""
                ws[f'E{row}'] = ""
                ws[f'F{row}'] = ""
                ws[f'G{row}'] = ""
                ws[f'H{row}'] = ""
                ws[f'I{row}'] = ""
                ws[f'J{row}'] = ""
                ws[f'K{row}'] = ""
                ws[f'L{row}'] = ""
                ws[f'M{row}'] = ""
                ws[f'N{row}'] = ""
                ws[f'O{row}'] = ""
                ws[f'P{row}'] = ""
                ws[f'Q{row}'] = ""
                ws[f'R{row}'] = ""
                ws[f'S{row}'] = ""
                ws[f'T{row}'] = ""
                ws[f'U{row}'] = ""
                ws[f'V{row}'] = ""
                ws[f'W{row}'] = ""
                ws[f'X{row}'] = ""
                row += 1
                continue
            
            # Extract fields from response
            lines = response_text.split('\n')
            appt_date = ""
            patient_name = ""
            patient_birthdate = ""
            facility_info = ""
            insurance_company = ""
            mem_id = ""
            
            for line in lines:
                if "Appt Date:" in line:
                    appt_date = line.split("Appt Date:")[-1].strip().replace("*", "")
                elif "Name of Patient:" in line:
                    patient_name = line.split("Name of Patient:")[-1].strip().replace("*", "")
                elif "Patient Birthdate:" in line:
                    patient_birthdate = line.split("Patient Birthdate:")[-1].strip().replace("*", "")
                elif "Facility Information:" in line:
                    facility_info = line.split("Facility Information:")[-1].strip().replace("*", "")
                elif "Primary:" in line:
                    extracted_primary = line.split("Primary:")[-1].strip().replace("*", "")
                    # Only set insurance_company if Primary has actual content (not empty)
                    if extracted_primary and len(extracted_primary.strip()) > 0:
                        insurance_company = extracted_primary
                elif "Sub/Member No.:" in line:
                    extracted_member = line.split("Sub/Member No.:")[-1].strip().replace("*", "")
                    # Only set mem_id if Sub/Member No. has actual content (not empty)
                    if extracted_member and len(extracted_member.strip()) > 0:
                        mem_id = extracted_member
            
            # Add all rows regardless of whether primary data is found
            # Only skip if completely empty
            has_any_data = (patient_name and len(patient_name.strip()) > 0) or \
                          (facility_info and len(facility_info.strip()) > 0) or \
                          (appt_date and len(appt_date.strip()) > 0)
            
            if has_any_data:
                facility_final = match_facility(facility_info)
                
                # Format dates to dd/mm/yyyy
                formatted_appt_date = format_date_to_ddmmyyyy(appt_date)
                formatted_birthdate = format_date_to_ddmmyyyy(patient_birthdate)
                
                ws[f'A{row}'] = ""  # Name of the Phleb
                # Write Date only if it's different from the previous row
                ws[f'B{row}'] = formatted_appt_date if formatted_appt_date != prev_date else ""  # Date
                ws[f'C{row}'] = ""  # No of Patient
                ws[f'D{row}'] = ""  # Patient ID
                ws[f'E{row}'] = ""  # patient bod
                ws[f'F{row}'] = patient_name  # Name of Patient
                ws[f'G{row}'] = formatted_birthdate  # Patient Birthdate
                # Write Facility only if it's different from the previous row
                ws[f'H{row}'] = facility_final if facility_final != prev_facility else ""  # Facility Information
                ws[f'I{row}'] = ""  # Patients ICD Code
                ws[f'J{row}'] = ""  # From
                ws[f'K{row}'] = ""  # To
                ws[f'L{row}'] = ""  # Miles
                ws[f'M{row}'] = ""  # To
                ws[f'N{row}'] = ""  # Miles
                ws[f'O{row}'] = ""  # To
                ws[f'P{row}'] = ""  # Miles
                ws[f'Q{row}'] = ""  # To
                ws[f'R{row}'] = ""  # Miles
                ws[f'S{row}'] = ""  # To
                ws[f'T{row}'] = ""  # Miles
                ws[f'U{row}'] = ""  # Total Miles
                ws[f'V{row}'] = insurance_company  # Insurance Company
                ws[f'W{row}'] = mem_id  # Mem ID
                ws[f'X{row}'] = ""  # Group Mem ID
                
                # Update tracking variables
                if formatted_appt_date:
                    prev_date = formatted_appt_date
                if facility_final:
                    prev_facility = facility_final
                
                # Apply red highlighting for missing fields
                red_fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
                
                # Check conditions for highlighting (only if NOT a DROP SHEET)
                is_drop_sheet = "DROP SHEET" in response_text.upper()
                
                if not is_drop_sheet:
                    insurance_present = bool(insurance_company and insurance_company.strip())
                    mem_id_present = bool(mem_id and mem_id.strip())
                    
                    # Condition 1: Both Insurance Company AND Mem ID are blank
                    # → Clear all three fields (Insurance Company, Mem ID, Group Mem ID)
                    if not insurance_present and not mem_id_present:
                        ws[f'V{row}'] = ""  # Insurance Company
                        ws[f'W{row}'] = ""  # Mem ID
                        ws[f'X{row}'] = ""  # Group Mem ID
                    
                    # Condition 2: Insurance Company is available BUT Mem ID is blank
                    # → Highlight Mem ID and Group Mem ID in red
                    elif insurance_present and not mem_id_present:
                        ws[f'W{row}'].fill = red_fill  # Mem ID
                        ws[f'X{row}'].fill = red_fill  # Group Mem ID
                
                row += 1
        
        # Save Excel file
        pdf_filename = os.path.splitext(pdf_file.name)[0]
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_file = f"{pdf_filename}_{timestamp}.xlsx"
        wb.save(output_file)
        
        return output_file
    
    finally:
        # Close PDF document properly
        if doc is not None:
            doc.close()
        
        # Clean up temp file with retry logic for Windows file locking
        try:
            os.unlink(pdf_path)
        except PermissionError:
            import time
            time.sleep(0.5)  # Wait a bit for file to be released
            try:
                os.unlink(pdf_path)
            except Exception as e:
                # If still can't delete, just log it (temp files will be cleaned up by OS)
                pass

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
