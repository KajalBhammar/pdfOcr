import streamlit as st
import base64
import fitz
from openpyxl import Workbook
from openpyxl.styles import PatternFill
import os
from datetime import datetime
from mistralai import Mistral
import tempfile
from datetime import datetime as dt

try:
    api_key = st.secrets["mistral_api_key"]
except KeyError:
    api_key = os.environ.get("MISTRAL_API_KEY")
    if not api_key:
        st.error("❌ API Key not found. Please set 'mistral_api_key' in Streamlit secrets or MISTRAL_API_KEY environment variable.")
        st.stop()

client = Mistral(api_key=api_key)

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
    if not date_str or date_str.lower() == "not found":
        return ""
    
    date_str = date_str.strip()
    
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
            return parsed_date.strftime("%m/%d/%Y")
        except ValueError:
            continue
    
    return ""


def extract_from_pdf(pdf_file, progress_bar, status_text):
    with tempfile.NamedTemporaryFile(delete=False, suffix='.pdf') as tmp:
        tmp.write(pdf_file.read())
        pdf_path = tmp.name
    
    doc = None
    try:
        doc = fitz.open(pdf_path)
        total_pages = len(doc)
        images = [doc[page_num].get_pixmap(matrix=fitz.Matrix(2, 2)) for page_num in range(total_pages)]
        
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
        
        for page_num, pixmap in enumerate(images, 1):
            progress = page_num / total_pages
            progress_bar.progress(progress)
            status_text.text(f"Processing page {page_num} of {total_pages}...")
            
            png_bytes = pixmap.tobytes("png")
            image_base64 = base64.b64encode(png_bytes).decode("utf-8")
            
            prompt_text = """You are a medical document extraction engine.

You will receive a single PDF page image.

Your job is to decide whether this page is a valid PATIENT REPORT with a BARCODE.
Only pages with a visible barcode are allowed to be processed.

----------------------------------------------------
STEP 1 — BARCODE CHECK

If NO barcode or QR code is visible:
Return ONLY:
Barcode Image Present: NO
and STOP.

If a barcode is visible:
Continue.

----------------------------------------------------
STEP 2 — DROP SHEET RULE

If the page contains the words "DROP SHEET" and also contains a barcode → STOP and return:
Barcode Image Present: NO

Drop sheets must NEVER be extracted.

----------------------------------------------------
STEP 3 — TEXT RULES

• Extract ONLY printed or typed text  
• Ignore ALL handwritten text  
• Never guess, infer, or copy values between fields  
• Never invent missing data  

----------------------------------------------------
STEP 4 — REQUIRED FIELDS (ALWAYS PRESENT ON BARCODE REPORTS)

From PATIENT INFORMATION section extract:

Patient Name  
Look for:
Name:
Patient:
Patient Name:
Name of Patient:

DOB  
Look for:
DOB:
Date of Birth:
Birthdate:
Birth Date:
D.O.B.:

These two MUST be found or the page is invalid.

----------------------------------------------------
STEP 5 — FACILITY INFORMATION

From FACILITY INFORMATION section extract:
• Facility Name
• Facility Address

Return them combined as:
Facility Name, Address

This value must later be matched against this allowed list:

Alliance Health at Marina Bay, 2 Seaport, Quincy, MA  
Alliance Health at West Acres, 804 Pleasant St, Brockton, MA  
Sherrill House, 135 S Huntington Ave, Jamaica Plain, MA  
Alliance Health at Maples, 90 Taunton St, Wrentham, MA 02097  
Oak Knoll, 9 Ambetter Dr, Framingham, MA  
Sippican Rehab & Healthcare, 15 Mill St, Marion, MA  
Alliance Health at Braintree, 175 Grove St, Braintree, MA  
Harrington House Rehab & Healthcare, 160 Main St, Walpole, MA  
Bethany Healthcare Rest Home, 97 Bethany Rd, Framingham, MA  
Alliance Health at Marie Esther, 720 Boston, Marlborough, MA  
Shrewsbury Nursing & Rehab Center, 40 Julio Dr, Shrewsbury, MA 01545  
Alliance Health at Doolittle Unit 1, 16 Bird St, Foxboro, MA  
CareOne at Concord, 57 Old Rd to 9 Acre Corner, Concord, MA 01742  
The Commons at Lincoln, 3 Harvest Cir, Lincoln, MA  
Rivercrest Nursing and Wellness, 100 Newbury Ct, Concord, MA  
Brookhaven at Lexington Independent Living, 1010 Waltham St, Lexington, MA  
Woburn Rehabilitation & Nursing Center, 18 Frances St, Woburn, MA 01801  
CareOne at Wilmington, 750 Woburn St, Wilmington, MA 01887  
CareOne at Lexington, 178 Lowell St, Lexington, MA 02420  
Winchester Rehabilitation and Nursing Center, 223 Swanton St, Winchester, MA 01890  
Aberjona Rehabilitation & Nursing Center, 184 Swanton St, Winchester, MA 01890  
The Commons in Lincoln, 1 Harvest Cir, Lincoln, MA 01773  
CareOne at Essex Park, 265 Essex St, Beverly, MA 01915  
CareOne at Peabody, 199 Andover St, Peabody, MA 01960  

----------------------------------------------------
STEP 6 — BILLING INFORMATION

From Billing Information extract:

Primary Insurance  
Look for:
Primary:
Insurance:
Primary Insurance:

Sub/Member No.  
Only if one of these labels exists:
Sub/Member No.
Subscriber ID
Subscriber No.
Member ID
Member No.
Mem #
ID Number (insurance section only)

Group Number  
Only if one of these labels exists:
Group Number
Group #
Grp #
Group No.
Group ID
Group Code

If a label does not exist → leave the field blank  
Never copy values between these fields

----------------------------------------------------
STEP 7 — DIAGNOSIS CODES

From the Symptoms or Diagnosis section:

Extract ONLY ICD-10 codes (examples: E11.9, I50.9, J44.9)  
Ignore all text descriptions  
Remove duplicates  
Return comma-separated  

----------------------------------------------------
FINAL OUTPUT FORMAT (EXACT — NO EXTRA TEXT)

Barcode Image Present: YES
Patient Name:
DOB:
Facility:
Primary:
Sub/Member No.:
Group Number:
Diagnosis Codes:"""

            inputs = [
                {
                    "role": "user",
                    "content": [
                        {
                            "type": "text", 
                            "text": prompt_text
                        },
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
            
            if hasattr(response.outputs[0], 'content'):
                response_text = response.outputs[0].content
            elif hasattr(response.outputs[0], 'text'):
                response_text = response.outputs[0].text
            elif hasattr(response.outputs[0], 'message') and hasattr(response.outputs[0].message, 'content'):
                response_text = response.outputs[0].message.content
            else:
                response_text = str(response.outputs[0])
            
            # Check if page has barcode image
            has_barcode = "Barcode Image Present: YES" in response_text or "barcode image present: yes" in response_text.lower()
            
            if not has_barcode:
                # Add a blank row for skipped pages
                row += 1
                continue
            
            lines = response_text.split('\n')
            patient_name = ""
            patient_dob = ""
            facility_info = ""
            insurance_company = ""
            mem_id = ""
            group_number = ""
            diagnosis_codes = ""
            
            for line in lines:
                if "Patient Name:" in line:
                    patient_name = line.split("Patient Name:")[-1].strip().replace("*", "")
                elif "DOB:" in line:
                    patient_dob = line.split("DOB:")[-1].strip().replace("*", "")
                elif "Facility:" in line:
                    facility_info = line.split("Facility:")[-1].strip().replace("*", "")
                elif "Primary:" in line:
                    extracted_primary = line.split("Primary:")[-1].strip().replace("*", "")
                    if extracted_primary and len(extracted_primary.strip()) > 0:
                        insurance_company = extracted_primary
                elif "Sub/Member No.:" in line:
                    extracted_member = line.split("Sub/Member No.:")[-1].strip().replace("*", "")
                    if extracted_member and len(extracted_member.strip()) > 0:
                        mem_id = extracted_member
                elif "Group Number:" in line:
                    extracted_group = line.split("Group Number:")[-1].strip().replace("*", "")
                    if extracted_group and len(extracted_group.strip()) > 0:
                        group_number = extracted_group
                elif "Diagnosis Codes:" in line:
                    extracted_codes = line.split("Diagnosis Codes:")[-1].strip().replace("*", "")
                    if extracted_codes and len(extracted_codes.strip()) > 0:
                        diagnosis_codes = extracted_codes
            
            has_any_data = (patient_name and len(patient_name.strip()) > 0) or \
                          (patient_dob and len(patient_dob.strip()) > 0) or \
                          (facility_info and len(facility_info.strip()) > 0) or \
                          (insurance_company and len(insurance_company.strip()) > 0) or \
                          (mem_id and len(mem_id.strip()) > 0) or \
                          (diagnosis_codes and len(diagnosis_codes.strip()) > 0)
            
            if has_any_data:
                facility_final = match_facility(facility_info)
                
                formatted_birthdate = format_date_to_ddmmyyyy(patient_dob)
                
                ws[f'A{row}'] = ""
                ws[f'B{row}'] = ""
                ws[f'C{row}'] = ""
                ws[f'D{row}'] = ""
                ws[f'E{row}'] = ""
                ws[f'F{row}'] = patient_name
                ws[f'G{row}'] = formatted_birthdate
                ws[f'H{row}'] = facility_final if facility_final != prev_facility else ""
                ws[f'I{row}'] = diagnosis_codes
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
                ws[f'V{row}'] = insurance_company
                ws[f'W{row}'] = mem_id
                ws[f'X{row}'] = group_number
                
                if facility_final:
                    prev_facility = facility_final
                
                red_fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
                
                is_drop_sheet = "DROP SHEET" in response_text.upper()
                
                if not is_drop_sheet:
                    insurance_present = bool(insurance_company and insurance_company.strip())
                    mem_id_present = bool(mem_id and mem_id.strip())
                    group_number_present = bool(group_number and group_number.strip())
                    
                    if not insurance_present:
                        ws[f'V{row}'].fill = red_fill
                    if not mem_id_present:
                        ws[f'W{row}'].fill = red_fill
                    if not group_number_present:
                        ws[f'X{row}'].fill = red_fill
                
                row += 1
        
        pdf_filename = os.path.splitext(pdf_file.name)[0]
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_file = f"{pdf_filename}_{timestamp}.xlsx"
        wb.save(output_file)
        
        return output_file
    
    finally:
        if doc is not None:
            doc.close()
        
        try:
            os.unlink(pdf_path)
        except PermissionError:
            import time
            time.sleep(0.5)
            try:
                os.unlink(pdf_path)
            except Exception as e:
                pass


st.set_page_config(page_title="PDF Data Extractor", layout="centered")
st.title("📄 PDF Data Extractor")

uploaded_file = st.file_uploader("Upload PDF", type="pdf")

if uploaded_file is not None:
    st.success(f"✅ File uploaded: {uploaded_file.name}")
    
    if st.button("🔄 Process PDF", use_container_width=True):
        progress_bar = st.progress(0)
        status_text = st.empty()
        
        try:
            output_file = extract_from_pdf(uploaded_file, progress_bar, status_text)
            
            progress_bar.progress(1.0)
            status_text.text("✅ Processing complete!")
            
            with open(output_file, "rb") as file:
                st.download_button(
                    label="📥 Download Excel File",
                    data=file,
                    file_name=os.path.basename(output_file),
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
            
            os.unlink(output_file)
        
        except Exception as e:
            st.error(f"❌ Error: {str(e)}")
