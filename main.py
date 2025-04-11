import streamlit as st
import pandas as pd
import PyPDF2
import io
import os
import tempfile
import base64
import google.generativeai as genai
from PIL import Image
import fitz
import time
from dotenv import load_dotenv
import json
import re
import hashlib


load_dotenv()

GOOGLE_API_KEY = os.getenv('GEMINI_API_KEY')
genai.configure(api_key=GOOGLE_API_KEY)

st.title('PDF Data Extraction')
st.write('Upload a PDF file to extract structured data for Excel')

def convert_pdf_page_to_image(pdf_bytes, page_num):
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    page = doc.load_page(page_num)
    pix = page.get_pixmap(matrix=fitz.Matrix(2, 2))
    img = Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
    return img

def process_image_with_gemini(image):
    model = genai.GenerativeModel("gemini-1.5-flash")
    
    with io.BytesIO() as output:
        image.save(output, format="PNG")
        image_bytes = output.getvalue()
    
    image_parts = [
        {
            "mime_type": "image/png",
            "data": base64.b64encode(image_bytes).decode('utf-8')
        }
    ]
    
    prompt = """
    Extract the text content from this image of a TATA Capital Housing Finance document.
    Focus on capturing all table data, especially the sections on:
    - Tranche disbursement details
    - Cumulative Disbursement amounts
    - Construction percentages
    - Collection/Promoters' contributions
    - Pre-Disbursement conditions
    - Covenants and their timelines
    
    Return all the text content from the image, preserving the structure and relationships.
    """
    
    try:
        response = model.generate_content(
            [prompt, image_parts[0]],
            generation_config={"temperature": 0.1}
        )
        return response.text
    except Exception as e:
        st.error(f"Error processing image with Gemini API: {e}")
        return ""
    
def extract_structured_data(full_text):
    model = genai.GenerativeModel("gemini-1.5-flash")

    
    prompt = f"""
    From the following extracted text from TATA Capital Housing Finance documents:
    
    {full_text}
    
    Extract and organize the data into two parts:
    
    PART 1: Extract this table data with these columns aligned by row:
    - Sr. No.
    - Tranche Amount (Rs Cr)
    - Cumulative Disbursement (Rs Cr)
    - Construction % (Europa, Mynsa & Capella)
    - Incremental Collection/Promoters' Contribution (Rs Cr)
    
    PART 2: Extract these as separate bullet point lists that apply to all rows:
    - Conditions Precedent: These are the "Pre-Disbursement" conditions for all loans
    - Conditions Subsequent with Frequency: These are the "Covenants" with both the Covenant description and Timeline
    
    Return as valid JSON in this exact format:
    {{
      "table_data": [
        {{
          "Sr. No.": 1,
          "Tranche Amount (Rs Cr)": 12.00,
          "Cumulative Disbursement (Rs Cr)": 12.00,
          "Construction % (Europa, Mynsa & Capella)": "",
          "Incremental Collection/Promoters' Contribution (Rs Cr)": ""
        }},
        {{
          "Sr. No.": 2,
          "Tranche Amount (Rs Cr)": 5.00,
          "Cumulative Disbursement (Rs Cr)": 17.00,
          "Construction % (Europa, Mynsa & Capella)": "10.00%",
          "Incremental Collection/Promoters' Contribution (Rs Cr)": 5.00
        }},
        // more rows...
      ],
      "conditions_precedent": [
        "Condition 1",
        "Condition 2",
        // more conditions...
      ],
      "conditions_subsequent": [
        "Covenant 1 - Timeline: Within X days...",
        "Covenant 2 - Timeline: Quarterly...",
        // more covenants...
      ]
    }}
    
    No explanations, no markdown formatting, just the JSON object.
    """
    
    try:
        response = model.generate_content(
            prompt,
            generation_config={"temperature": 0.1}
        )
        return response.text
    except Exception as e:
        st.error(f"Error extracting structured data: {e}")
        return ""

def create_excel(data):
    try:
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            
            table_data = pd.DataFrame(data.get("table_data", []))
            
            conditions_precedent = data.get("conditions_precedent", [])
            conditions_subsequent = data.get("conditions_subsequent", [])
            
            conditions_precedent_text = "\n".join([f"{i+1}. {item}" for i, item in enumerate(conditions_precedent)])
            conditions_subsequent_text = "\n".join([f"{i+1}. {item}" for i, item in enumerate(conditions_subsequent)])
            
            if not table_data.empty:
                table_data["Conditions Precedent"] = conditions_precedent_text
                table_data["Conditions Subsequent with Frequency"] = conditions_subsequent_text
            
            table_data.to_excel(writer, sheet_name="Extracted Data", index=False)
            
            conditions_df = pd.DataFrame({
                "Conditions Precedent": pd.Series(conditions_precedent),
                "Conditions Subsequent with Frequency": pd.Series(conditions_subsequent)
            })
            conditions_df.to_excel(writer, sheet_name="Conditions Detail", index=False)
            
            workbook = writer.book
            worksheet = writer.sheets["Extracted Data"]
            
            wrap_format = workbook.add_format({'text_wrap': True, 'valign': 'top'})
            worksheet.set_column('F:G', 50, wrap_format)
        
        return output.getvalue()
    except Exception as e:
        st.error(f"Error creating Excel file: {e}")
        return None
    
def get_file_hash(file_bytes):
    return hashlib.md5(file_bytes).hexdigest()
    
uploaded_file = st.file_uploader("Choose a PDF file", type="pdf")

if uploaded_file is not None:
    pdf_bytes = uploaded_file.read()
    file_hash = get_file_hash(pdf_bytes)

    if st.session_state.get("file_hash") != file_hash:
        with st.spinner('Processing PDF...'):
            reader = PyPDF2.PdfReader(io.BytesIO(pdf_bytes))
            num_pages = len(reader.pages)
            full_text = ""

            progress_bar = st.progress(0)
            for i in range(num_pages):
                progress_bar.progress((i + 1) / num_pages)
                image = convert_pdf_page_to_image(pdf_bytes, i)
                page_text = process_image_with_gemini(image)
                full_text += f"\n\n--- PAGE {i+1} ---\n\n{page_text}"
                time.sleep(1)

            json_data = extract_structured_data(full_text)

            try:
                try:
                    data = json.loads(json_data)
                except:
                    json_match = re.search(r'(\{.*\})', json_data, re.DOTALL)
                    if json_match:
                        clean_json = json_match.group(1)
                        data = json.loads(clean_json)
                    else:
                        st.error("Could not parse JSON data from response")
                        st.text(json_data)
                        st.download_button(
                            label="Download raw extracted text",
                            data=full_text,
                            file_name="raw_extracted_text.txt",
                            mime="text/plain"
                        )
                        st.stop()

                excel_data = create_excel(data)
                st.session_state["file_hash"] = file_hash
                st.session_state["full_text"] = full_text
                st.session_state["json_data"] = json_data
                st.session_state["parsed_data"] = data
                st.session_state["excel_data"] = excel_data

            except Exception as e:
                st.error(f"Error processing data: {str(e)}")
                st.text(json_data)

# After processing
if "excel_data" in st.session_state:
    st.download_button(
        label="Download Excel file",
        data=st.session_state["excel_data"],
        file_name="tata_finance_data.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )

    st.subheader("Preview of extracted data:")
    data = st.session_state["parsed_data"]

    if "table_data" in data:
        st.write("Table Data:")
        st.dataframe(pd.DataFrame(data["table_data"]))

    if "conditions_precedent" in data:
        st.write("Conditions Precedent:")
        for i, item in enumerate(data["conditions_precedent"]):
            st.write(f"{i+1}. {item}")

    if "conditions_subsequent" in data:
        st.write("Conditions Subsequent with Frequency:")
        for i, item in enumerate(data["conditions_subsequent"]):
            st.write(f"{i+1}. {item}")