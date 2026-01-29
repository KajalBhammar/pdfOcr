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
            
            inputs = [
                {
                    "role": "user",
                    "content": [
                        {
                            "type": "text", 
                            "text": "IMPORTANT: First check if this page contains a visible barcode/QR code image.\n\nBarcode Image Present: [YES or NO]\n\nIf NO barcode image is present on this page, respond ONLY with 'Barcode Image Present: NO' and do NOT extract any data.\n\nIf YES barcode image is present, continue with extraction.\n\n---\n\nAlso check if this page contains 'DROP SHEET'. If it does:\nDROP SHEET\n\nIf this page is a DROP SHEET without a barcode image, do NOT extract any data.\nIf this page is a DROP SHEET with a barcode image, continue to extract data.\n\n---\n\nIf this is NOT a DROP SHEET page and HAS a barcode image, extract ONLY PRINTED/TYPED text from this document. Skip all handwritten text completely.\n\n🚨 ABSOLUTE CRITICAL RULES - VIOLATING THESE IS AN ERROR 🚨\n\n⚠️ MANDATORY FIELDS ON BARCODE PAGES:\n- EVERY page with a barcode WILL have \"Name of Patient\" and \"Patient Birthdate\" fields\n- These fields are ALWAYS present on barcode pages\n- If you cannot find these fields, you are not looking carefully enough - LOOK AGAIN\n- Look for these fields in the \"Patient Information\" section of the document\n\n1. ❌ NEVER COPY VALUES BETWEEN FIELDS\n   - If you copy ANY value from one field to another field, you have made a critical error\n   - Example of WRONG behavior: If Sub/Member No. is \"101338884000\" and you don't see \"Group Number:\" label, DO NOT write \"101338884000\" for Group Number\n   - CORRECT behavior: If you don't see the label, leave that field BLANK\n\n2. ❌ NEVER USE PLACEHOLDER TEXT\n   - If a field is not found, leave it COMPLETELY BLANK\n   - DO NOT write: '[NOT FOUND]', 'N/A', '(not printed)', '[NOT PROVIDED]', 'None', 'blank', or any other text\n   - Just leave it empty with nothing after the colon\n\n3. ✅ ONLY EXTRACT WHAT YOU SEE WITH ITS LABEL\n   - Only extract a field if you can see its specific label in the document\n   - If the label doesn't exist, the field stays blank\n\n---\n\n*** CRITICAL: SUB/MEMBER NO. vs GROUP NUMBER ***\n\n🛑 STOP AND READ THIS CAREFULLY:\n\nThese are TWO COMPLETELY SEPARATE AND INDEPENDENT fields. They are NEVER related to each other.\n\n📋 Sub/Member No. EXTRACTION PROCESS:\nStep 1: Search the document for these EXACT label words:\n   - \"Sub/Member No.\"\n   - \"Subscriber ID\" or \"Subscriber No.\" or \"Subscriber #\"\n   - \"Member ID\" or \"Member No.\" or \"Member #\" or \"Mem #\"\n   - \"ID Number\" (only in insurance billing section)\n\nStep 2: If you find ONE of these labels, extract the value next to it\nStep 3: If you DO NOT find any of these labels, leave Sub/Member No. BLANK\n\n🏢 Group Number EXTRACTION PROCESS:\nStep 1: Search the document for these EXACT label words (DIFFERENT from above):\n   - \"Group Number\"\n   - \"Group #\" or \"Grp #\" or \"Group No.\"\n   - \"Group ID\" or \"Group Code\"\n\nStep 2: If you find ONE of these labels, extract the value next to it\nStep 3: If you DO NOT find any of these labels, leave Group Number BLANK\n\n⚠️ CRITICAL VALIDATION BEFORE RESPONDING:\n\n❌ WRONG EXAMPLES (These are ERRORS):\n   Sub/Member No.: 101338884000\n   Group Number: 101338884000\n   ↑ ERROR: Same value for both = YOU COPIED THE FIELD\n\n   Sub/Member No.: 5M62HF1DR98\n   Group Number: 5M62HF1DR98\n   ↑ ERROR: Same value for both = YOU COPIED THE FIELD\n\n✅ CORRECT EXAMPLES:\n   Sub/Member No.: 8F23MW6MX09\n   Group Number: MCRMA\n   ↑ CORRECT: Different values extracted from their own labels\n\n   Sub/Member No.: 101338884000\n   Group Number:\n   ↑ CORRECT: Only Sub/Member No. label was found, Group Number label was not found so it's blank\n\n   Sub/Member No.: 5KK5N02QN87\n   Group Number:\n   ↑ CORRECT: Only found Member ID label, no Group Number label found so it's blank\n\n🔴 IF SUB/MEMBER NO. AND GROUP NUMBER HAVE THE SAME VALUE, YOU MADE A MISTAKE!\n🔴 GO BACK AND CHECK THE DOCUMENT - DON'T JUST COPY VALUES!\n\n---\n\nExtract the following information:\n1. Appointment Date (Appt Date) - Find the appointment date\n2. Name of Patient - Copy EXACTLY as printed, no rearranging\n3. Patient Birthdate (DOB) - Find the date of birth\n4. Facility Information - Facility name AND full address combined\n5. Primary Insurance Company - Insurance company name\n6. Sub/Member No. - Find ONLY if \"Member ID\" or \"Subscriber ID\" label exists (otherwise BLANK)\n7. Group Number - Find ONLY if \"Group Number\" or \"Group ID\" label exists (otherwise BLANK)\n8. Diagnosis Codes - All ICD codes found\n\nIMPORTANT FOR PATIENT NAME:\n- Copy EXACTLY as it appears in the PDF\n- Do NOT rearrange, reformat, or change the order\n- If it says 'DAVID FREITAS W', write 'DAVID FREITAS W'\n- Do NOT change it to 'FREITAS, DAVID'\n\nIMPORTANT FOR DIAGNOSIS CODES:\n- Extract all diagnosis codes (E11.9, I50.9, I10, J44.9, etc.)\n- If the same code appears multiple times, include it only once\n- Output format: code1, code2, code3 (comma-separated, no descriptions)\n\nReturn the data in this EXACT format:\nBarcode Image Present: [YES or NO]\nAppt Date: [date or blank]\nName of Patient: [name or blank]\nPatient Birthdate: [DOB or blank]\nFacility Information: [facility name and address or blank]\nPrimary: [Insurance Company name or blank]\nSub/Member No.: [value only if you found its label, otherwise blank]\nGroup Number: [value only if you found its label, otherwise blank]\nDiagnosis Codes: [code1, code2, code3 or blank]\n\n🔍 MANDATORY FINAL VERIFICATION - CHECK EACH ITEM:\n✓ Are Sub/Member No. and Group Number DIFFERENT values? (If same = ERROR)\n✓ Did I extract Sub/Member No. only from a Member/Subscriber ID label? (Not guessed)\n✓ Did I extract Group Number only from a Group Number/ID label? (Not copied from Member ID)\n✓ Did I leave fields blank if their labels were not found? (No placeholder text)\n✓ Did I copy patient name exactly as shown? (No reformatting)"
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
            appt_date = ""
            patient_name = ""
            patient_birthdate = ""
            facility_info = ""
            insurance_company = ""
            mem_id = ""
            group_mem_id = ""
            diagnosis_codes = ""
            
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
                    if extracted_primary and len(extracted_primary.strip()) > 0:
                        insurance_company = extracted_primary
                elif "Sub/Member No.:" in line:
                    extracted_member = line.split("Sub/Member No.:")[-1].strip().replace("*", "")
                    if extracted_member and len(extracted_member.strip()) > 0:
                        mem_id = extracted_member
                elif "Group Number:" in line or "Group ID:" in line or "Grp #:" in line:
                    if "Group Number:" in line:
                        extracted_group = line.split("Group Number:")[-1].strip().replace("*", "")
                    elif "Group ID:" in line:
                        extracted_group = line.split("Group ID:")[-1].strip().replace("*", "")
                    else:
                        extracted_group = line.split("Grp #:")[-1].strip().replace("*", "")
                    
                    # Make sure it's not capturing address, Mem ID, Member ID, or other unwanted data
                    exclude_words = ["address", "mem", "member", "sub", "insurance", "company", ","]
                    should_exclude = any(word in extracted_group.lower() or extracted_group.find(word) != -1 for word in exclude_words)
                    
                    if extracted_group and len(extracted_group.strip()) > 0 and not should_exclude:
                        group_mem_id = extracted_group
                elif "Diagnosis Codes:" in line:
                    extracted_codes = line.split("Diagnosis Codes:")[-1].strip().replace("*", "")
                    if extracted_codes and len(extracted_codes.strip()) > 0:
                        diagnosis_codes = extracted_codes
            
            has_any_data = (patient_name and len(patient_name.strip()) > 0) or \
                          (patient_birthdate and len(patient_birthdate.strip()) > 0) or \
                          (facility_info and len(facility_info.strip()) > 0) or \
                          (appt_date and len(appt_date.strip()) > 0) or \
                          (insurance_company and len(insurance_company.strip()) > 0) or \
                          (mem_id and len(mem_id.strip()) > 0) or \
                          (diagnosis_codes and len(diagnosis_codes.strip()) > 0)
            
            if has_any_data:
                facility_final = match_facility(facility_info)
                
                formatted_appt_date = format_date_to_ddmmyyyy(appt_date)
                formatted_birthdate = format_date_to_ddmmyyyy(patient_birthdate)
                
                ws[f'A{row}'] = ""
                ws[f'B{row}'] = formatted_appt_date if formatted_appt_date != prev_date else ""
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
                ws[f'X{row}'] = group_mem_id
                
                if formatted_appt_date:
                    prev_date = formatted_appt_date
                if facility_final:
                    prev_facility = facility_final
                
                red_fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")
                
                is_drop_sheet = "DROP SHEET" in response_text.upper()
                
                if not is_drop_sheet:
                    insurance_present = bool(insurance_company and insurance_company.strip())
                    mem_id_present = bool(mem_id and mem_id.strip())
                    group_mem_id_present = bool(group_mem_id and group_mem_id.strip())
                    
                    if not insurance_present:
                        ws[f'V{row}'].fill = red_fill
                    if not mem_id_present:
                        ws[f'W{row}'].fill = red_fill
                    if not group_mem_id_present:
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
def extract_unique_diagnosis_codes(text):
    import re
    
    pattern = r'\b([A-Z0-9]+(?:\.[0-9]+)?)\b'
    
    matches = re.findall(pattern, text)
    
    codes = []
    for match in matches:
        if re.match(r'^[A-Z]\d+(?:\.\d+)?$', match) or re.match(r'^\d+(?:\.\d+)?$', match):
            codes.append(match)
    
    seen = set()
    unique_codes = []
    for code in codes:
        code_upper = code.upper()
        if code_upper not in seen:
            seen.add(code_upper)
            unique_codes.append(code)
    
    return unique_codes


st.set_page_config(page_title="PDF Data Extractor & Diagnosis Code Extractor", layout="centered")
st.title("📄 PDF Data Extractor & 🏥 Diagnosis Code Extractor")

tab1, tab2 = st.tabs(["PDF Data Extractor", "Diagnosis Code Extractor"])

with tab1:
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

with tab2:
    st.header("Extract Unique Diagnosis Codes")
    st.write("Paste unstructured medical text containing diagnosis codes. The tool will extract and deduplicate them.")
    
    medical_text = st.text_area(
        "Paste medical text here:",
        height=200,
        placeholder="Example: Patient has E11.9 diabetes, 150.9 hypertension, and E11.9 type 2 diabetes..."
    )
    
    if st.button("🔍 Extract Diagnosis Codes", use_container_width=True):
        if medical_text.strip():
            unique_codes = extract_unique_diagnosis_codes(medical_text)
            
            if unique_codes:
                codes_output = ", ".join(unique_codes)
                st.success("✅ Unique Diagnosis Codes:")
                st.code(codes_output, language="text")
                
                st.write("**Result:**")
                col1, col2 = st.columns([3, 1])
                with col1:
                    st.text_input("Codes:", value=codes_output, disabled=True)
                with col2:
                    st.write("")
                    st.write("")
                    if st.button("📋 Copy", use_container_width=True):
                        st.write("Copied to clipboard!")
            else:
                st.warning("⚠️ No diagnosis codes found in the text.")
        else:
            st.warning("⚠️ Please enter some medical text.")
