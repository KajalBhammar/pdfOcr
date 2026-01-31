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
                            "text": "IMPORTANT: First check if this page contains a visible barcode/QR code image.\n\nBarcode Image Present: [YES or NO]\n\nIf NO barcode image is present, respond ONLY with 'Barcode Image Present: NO' and stop.\n\nIf YES barcode image is present, continue with extraction.\n\n---\n\n🚨 100% ACCURACY EXTRACTION - 2 TIMES SEARCH, 2 TIMES FILTER, 2 TIMES CHECK 🚨\n\nFOR EVERY FIELD: Search 2 times, Filter 2 times, Check 2 times before deciding BLANK\nThis ensures 100% accurate results - NO SKIPPED FIELDS\n\nFIELDS TO EXTRACT:\n1. Appt Date (if present)\n2. Name of Patient (NOT always present - but extract if found)\n3. Patient Birthdate/DOB (MANDATORY - always present)\n4. Facility Information (Name AND Address together)\n5. Primary Insurance\n6. Sub/Member No. (Search 2 times - some reports have this)\n7. Group Number (Search 2 times - some reports have this)\n8. Diagnosis Codes (ONLY medical codes)\n\n---\n\n🔍 2-PASS + 2-FILTER + 2-CHECK VERIFICATION METHOD:\n\nFor EACH field, follow this process to avoid skipping:\n\n**SEARCH 1 - First Look:**\n- Scan entire document top to bottom\n- Look in ALL sections (Patient, Billing, Insurance, Symptoms, etc.)\n- Check different parts of page\n\n**SEARCH 2 - Second Look:**\n- Search again - different areas\n- Some fields are hidden or in unexpected places\n- Check headers, footers, middle sections\n\n**FILTER 1 - Initial Filter:**\n- From both searches, identify exact label and value\n- Write down what you found\n\n**FILTER 2 - Re-Filter:**\n- Go back and verify the value is complete\n- Not cut off, not partial, not truncated\n\n**CHECK 1 - Initial Verification:**\n- Verify this looks like a real value\n- Not an error or placeholder\n\n**CHECK 2 - Final Check:**\n- Search one more time to be absolutely sure\n- Only then decide if blank\n\n---\n\n📋 DETAILED EXTRACTION INSTRUCTIONS:\n\n1️⃣ Appt Date:\n   - Search 1: Look for \"Appt Date\", \"Appointment Date\", \"Collection Date\", \"Order Date\"\n   - Search 2: Check billing, header, footer sections\n   - Extract: DATE VALUE only\n   - If not found after 2 searches: leave BLANK\n\n2️⃣ Name of Patient:\n   - Search 1: Look for \"Name:\", \"Patient:\", \"Patient Name:\", \"Name of Patient:\"\n   - Search 2: Check multiple sections - name might appear multiple times\n   - Filter: Copy EXACTLY as printed - NO REARRANGING\n   - Check: Verify it's a real name, not data error\n   - ⚠️ NOT all reports have name - but DO search 2 times before deciding blank\n   - Example: \"LINDA BONARRIGO\" stays \"LINDA BONARRIGO\" (NOT \"BONARRIGO, LINDA\")\n\n3️⃣ Patient Birthdate/DOB (MANDATORY - ALWAYS PRESENT):\n   - Search 1: Look for \"DOB:\", \"Date of Birth:\", \"Birthdate:\", \"Birth Date:\", \"D.O.B.:\"\n   - Search 2: Check all sections - DOB appears on every barcode\n   - Filter: Get complete DATE\n   - Check: Verify date is valid (not cut off)\n   - CRITICAL: This is mandatory - search 2 times minimum\n\n4️⃣ Facility Information:\n   - Search 1: Look for \"Facility:\", \"Facility Name:\", \"Location:\", \"Provider:\", \"Nursing Home:\"\n   - Search 2: Check header, billing section, middle of document\n   - Filter: Get BOTH name AND address together\n   - Combine if split across lines\n   - Example: \"Alliance Health at Marina Bay, 2 Seaport, Quincy, MA\"\n\n5️⃣ Primary (Insurance Company):\n   - Search 1: Look for \"Primary:\", \"Insurance:\", \"Primary Insurance:\", \"Carrier:\"\n   - Search 2: Check billing section multiple times\n   - Filter: Extract company name ONLY\n   - If not found after 2 searches: leave BLANK\n\n6️⃣ Sub/Member No. (Search 2 times - DO NOT SKIP):\n   - Search 1: Look for \"Sub/Member No.:\", \"Member ID:\", \"Member No.:\", \"Mem #:\", \"Subscriber ID:\"\n   - Search 2: Search ENTIRE document again - might be in multiple places\n   - Filter 1: Extract exact number next to label\n   - Filter 2: Verify it's complete (might be multi-part like 8F23MW6MX09)\n   - Check 1: Verify it's a real ID number\n   - Check 2: Search one more time before deciding blank\n   - CRITICAL: Search 2 FULL TIMES - some reports have this, some don't\n   - If label NOT found after 2 complete searches, leave BLANK\n\n7️⃣ Group Number (Search 2 times - DO NOT SKIP):\n   - Search 1: Look for \"Group Number:\", \"Group #:\", \"Grp #:\", \"Group ID:\", \"Group Code:\", \"Group Name:\"\n   - Search 2: Search ENTIRE document again - check all sections\n   - Filter 1: Extract exact value next to label\n   - Filter 2: Verify complete (could be short like \"MCRMA\" or numeric)\n   - Check 1: Verify it's a real group ID\n   - Check 2: Search one more time before deciding blank\n   - CRITICAL: Some reports HAVE group numbers - search 2 FULL TIMES\n   - If label NOT found after 2 complete searches, leave BLANK\n\n⚠️ CRITICAL: Sub/Member No. vs Group Number:\n   - They must be DIFFERENT (or one blank)\n   - ❌ ERROR: Same value = YOU DIDN'T SEARCH 2 TIMES PROPERLY\n   - ✅ CORRECT: Different values OR one is blank\n\n8️⃣ Diagnosis Codes:\n   - Search 1: Look for \"Diagnosis:\", \"Diagnoses:\", \"Codes:\", \"ICD:\", \"Symptoms:\", \"Conditions:\"\n   - Search 2: Search entire document - codes often in multiple places\n   - Filter: Extract ONLY medical codes (E11.9, I50.9, I10, etc.)\n   - Format: code1, code2, code3 (separated by comma and space)\n   - Check: List each code ONCE only (no duplicates)\n   - ONLY codes - NO descriptions, NO text\n\n---\n\n🚨 ZERO TOLERANCE FOR PLACEHOLDER TEXT 🚨\n\nSTRICTLY FORBIDDEN - LLM CANNOT WRITE:\n❌ \"not found\", \"[blank]\", \"blank\", \"n/a\", \"N/A\"\n❌ \"missing\", \"not available\", \"(blank)\", \"[not available]\"\n❌ \"none\", \"no data\", \"not provided\", \"(missing)\"\n❌ Empty brackets [], parentheses (), dashes or any symbols\n❌ Descriptive text like \"no data\", \"blank field\", \"not applicable\"\n\nLLM MUST ONLY:\n✅ Write ACTUAL extracted values from the document\n✅ OR leave the field COMPLETELY EMPTY if not found\n✅ Copy text EXACTLY as printed (preserve case, spacing, formatting)\n✅ Do NOT invent or assume ANY text\n\nIF Insurance Company, Mem ID, or Group Mem ID fields:\n🔴 Have placeholder text → REJECT IT (treat as empty)\n🔴 Are completely empty → Leave BLANK (will be highlighted RED in Excel)\n\n---\n\n✅ DO THIS:\n- Search 2 complete times for each field\n- Filter results to get exact values\n- Check 2 times before deciding blank\n- Never skip a field - search thoroughly\n- Copy names EXACTLY as printed\n- Get complete multi-part values (like 8F23MW6MX09)\n- Use only medical codes\n\n❌ DO NOT DO THIS:\n- Copy values between different fields\n- Use placeholder text\n- Skip fields without searching 2 times\n- Assume a field doesn't exist\n- Modify patient names\n\n---\n\nReturn in this EXACT format (FIXED FIELD NAMES):\nBarcode Image Present: [YES or NO]\nAppt Date: [value or blank]\nName of Patient: [value or blank]\nPatient Birthdate: [value]\nFacility Information: [value or blank]\nPrimary: [value or blank]\nSub/Member No.: [value or blank]\nGroup Number: [value or blank]\nDiagnosis Codes: [value or blank]\n\n🔍 FINAL 100% ACCURACY CHECKLIST:\n✓ Patient Birthdate found? (MANDATORY)\n✓ Searched 2 times for Name of Patient? (Optional but search 2 times)\n✓ Searched 2 times for Sub/Member No.? (Some have it)\n✓ Searched 2 times for Group Number? (Some have it)\n✓ Sub/Member No. and Group Number different or one blank?\n✓ All field names exact?\n✓ No placeholder text?\n✓ Did NOT skip any field - searched 2 times minimum?"
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
            
            # Define placeholder rejection list - ZERO TOLERANCE FOR PLACEHOLDERS
            placeholder_keywords = ["not found", "[blank]", "blank", "n/a", "missing", "not available", "(blank)", "[not available]", "none", "no data", "not provided"]
            
            def is_valid_value(text):
                """Check if text is a valid value (not a placeholder) - STRICT VALIDATION"""
                if not text or len(text.strip()) == 0:
                    return False
                text_lower = text.lower().strip()
                return not any(keyword in text_lower for keyword in placeholder_keywords)
            
            for line in lines:
                if "Appt Date:" in line:
                    appt_date = line.split("Appt Date:")[-1].strip().replace("*", "")
                    if not is_valid_value(appt_date):
                        appt_date = ""
                elif "Name of Patient:" in line:
                    patient_name = line.split("Name of Patient:")[-1].strip().replace("*", "")
                    if not is_valid_value(patient_name):
                        patient_name = ""
                elif "Patient Birthdate:" in line:
                    patient_birthdate = line.split("Patient Birthdate:")[-1].strip().replace("*", "")
                    if not is_valid_value(patient_birthdate):
                        patient_birthdate = ""
                elif "Facility Information:" in line:
                    facility_info = line.split("Facility Information:")[-1].strip().replace("*", "")
                    if not is_valid_value(facility_info):
                        facility_info = ""
                elif "Primary:" in line:
                    extracted_primary = line.split("Primary:")[-1].strip().replace("*", "")
                    if is_valid_value(extracted_primary):
                        insurance_company = extracted_primary
                elif "Sub/Member No.:" in line:
                    extracted_member = line.split("Sub/Member No.:")[-1].strip().replace("*", "")
                    if is_valid_value(extracted_member):
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
                    
                    if is_valid_value(extracted_group) and not should_exclude:
                        group_mem_id = extracted_group
                elif "Diagnosis Codes:" in line:
                    extracted_codes = line.split("Diagnosis Codes:")[-1].strip().replace("*", "")
                    if is_valid_value(extracted_codes):
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
