import json
import streamlit as st
import os
from logic import (
    create_new_document,
    save_uploaded_file,
    extract_text_from_pdf,
    extract_text_from_image,
    extract_details_with_gpt,
    extract_plan_details_with_gpt,
    generate_pension_review_section,
    generate_safe_withdrawal_rate_section,
    extract_fund_performance_with_gpt,
    extract_dark_star_performance_with_gpt,
    extract_sap_comparison_with_gpt,
    extract_annuity_quotes_with_gpt,
    extract_fund_comparison_with_gpt,
    extract_iht_details_with_gpt
)
import openai



# Define folders for uploaded and generated documents
UPLOAD_FOLDER = "uploaded_docs"
OUTPUT_FOLDER = "generated_docs"

# Streamlit Page Configuration
st.set_page_config(page_title="Zomi AI Persona", page_icon="💼", layout="wide")

# Inject CSS for Dark Mode Styling
st.markdown(
    """
    <style>
    /* [Your existing CSS styles here] */
    /* ... */
    </style>
    """,
    unsafe_allow_html=True,
)

# Title Section
st.markdown('<div class="title">Zomi Wealth AI</div>', unsafe_allow_html=True)
st.markdown('<div class="subtitle">Generate personalized financial reports with ease.</div>', unsafe_allow_html=True)

# File Upload Section
st.markdown('<div class="upload-section">', unsafe_allow_html=True)

uploaded_template = st.file_uploader("📄 Upload Report Template (.docx)", type="docx")
uploaded_factfind = st.file_uploader("📄 Upload FactFind Document (.pdf)", type="pdf")
uploaded_risk_profile = st.file_uploader("Upload Risk Profile Document/Image", type=["png", "jpg", "jpeg", "pdf"])

# New: Upload Plan Report Files
uploaded_files = st.file_uploader(
    "📤 Upload the Plan Report Files", type=["docx", "pdf", "png", "jpg", "jpeg"], accept_multiple_files=True
)

# New: Upload Client Fund Fact Sheet
uploaded_fund_fact_sheet = st.file_uploader("📄 Upload Client Fund Fact Sheet (.pdf)", type="pdf")
uploaded_dark_star_fact_sheet = st.file_uploader("📄 Upload Dark Star Fact Sheet (.pdf)", type="pdf")
# Add upload button for SAP report
uploaded_sap_report = st.file_uploader("📄 Upload SAP Report File (.pdf)", type="pdf")
# Add upload button for Annuity Quotes image
uploaded_annuity_image = st.file_uploader("📤 Upload Annuity Quotes Image", type=["png", "jpg", "jpeg"])

# Add upload button for multiple files required for the Fund Comparison table
st.markdown("### Fund Comparison Files")
st.markdown("Upload the necessary files for the fund comparison table.")

# Define the required files for the comparison table
REQUIRED_FILES = ["Client Fund Details", "P1 Fund Details", "Additional Comparison Data"]

# Track uploaded files for the comparison table
if "uploaded_comparison_files" not in st.session_state:
    st.session_state["uploaded_comparison_files"] = {}

# Multiple file uploader for fund comparison
uploaded_comparison_files = st.file_uploader(
    "📤 Upload files for Fund Comparison (PDF or Excel)", 
    type=["pdf", "xlsx"], 
    accept_multiple_files=True, 
    key="comparison_file_uploader"
)


st.markdown('</div>', unsafe_allow_html=True)

# Logic to handle uploaded files
if uploaded_template and uploaded_factfind and uploaded_risk_profile:
    template_path = save_uploaded_file(uploaded_template, UPLOAD_FOLDER)
    factfind_path = save_uploaded_file(uploaded_factfind, UPLOAD_FOLDER)
    risk_profile_path = save_uploaded_file(uploaded_risk_profile, UPLOAD_FOLDER)

    plan_report_data = []
    plan_report_text = ""
    fund_performance_data = []
    dark_star_performance_data = [] 

    
     # New list for Dark Star performance
    factfinding_text = extract_text_from_pdf(factfind_path)
    risk_text = extract_text_from_pdf(risk_profile_path) if uploaded_risk_profile.name.endswith("pdf") else extract_text_from_image(risk_profile_path)
    risk_details = extract_details_with_gpt(risk_text)
        
    
    if uploaded_files:
        for file in uploaded_files:
            file_path = save_uploaded_file(file, UPLOAD_FOLDER)
            extracted_text = extract_text_from_pdf(file_path) if file.name.endswith("pdf") else extract_text_from_image(file_path)
            if extracted_text:
                plan_details = extract_plan_details_with_gpt(extracted_text)
                plan_report_data.extend(plan_details)
                plan_report_text += extracted_text + "\n"
    
    
    product_report_text = plan_report_text  # Modify this line based on actual data sources
            
    # Handle Client Fund Fact Sheet
    if uploaded_fund_fact_sheet:
        fund_fact_sheet_path = save_uploaded_file(uploaded_fund_fact_sheet, UPLOAD_FOLDER)
        extracted_fund_text = extract_text_from_pdf(fund_fact_sheet_path)
        if extracted_fund_text:
            fund_performance_data = extract_fund_performance_with_gpt(extracted_fund_text)

    # Handle Dark Star Fact Sheet
    if uploaded_dark_star_fact_sheet:
        dark_star_fact_sheet_path = save_uploaded_file(uploaded_dark_star_fact_sheet, UPLOAD_FOLDER)
        extracted_dark_star_text = extract_text_from_pdf(dark_star_fact_sheet_path)
        if extracted_dark_star_text:
            dark_star_performance_data = extract_dark_star_performance_with_gpt(extracted_dark_star_text)
    # Logic to handle the SAP report file
    sap_comparison_table = None
    if uploaded_sap_report:
        sap_report_path = save_uploaded_file(uploaded_sap_report, UPLOAD_FOLDER)
        extracted_sap_text = extract_text_from_pdf(sap_report_path)
        if extracted_sap_text:
            try:
                sap_comparison_table = extract_sap_comparison_with_gpt(extracted_sap_text)
            except Exception as e:
                st.error(f"❌ An error occurred while processing the SAP report: {e}") 

    annuity_quotes_text = None
    if uploaded_annuity_image:
        annuity_image_path = save_uploaded_file(uploaded_annuity_image, UPLOAD_FOLDER)
        extracted_annuity_text = extract_text_from_image(annuity_image_path)

        # Use GPT to structure the extracted data
        try:
            annuity_quotes_text = extract_annuity_quotes_with_gpt(extracted_annuity_text)
        except Exception as e:
            st.error(f"❌ An error occurred while processing the annuity image: {e}")
            
    fund_comparison_text = None
    if uploaded_comparison_files:
        # Ensure at least three files are uploaded for comparison
        if len(uploaded_comparison_files) >= 3:
            fund1_file = uploaded_comparison_files[0]
            fund2_file1 = uploaded_comparison_files[1]
            fund2_file2 = uploaded_comparison_files[2]
            fund1_path = save_uploaded_file(fund1_file, UPLOAD_FOLDER)
            fund2_file1_path = save_uploaded_file(fund2_file1, UPLOAD_FOLDER)
            fund2_file2_path = save_uploaded_file(fund2_file2, UPLOAD_FOLDER)
            fund1_extracted_text = extract_text_from_pdf(fund1_path)
            fund2_file1_extracted_text = extract_text_from_pdf(fund2_file1_path)
            fund2_file2_extracted_text = extract_text_from_pdf(fund2_file2_path)
            try:
                fund_comparison_text = extract_fund_comparison_with_gpt(
                    fund1_text=fund1_extracted_text,
                    fund2_file1_text=fund2_file1_extracted_text,
                    fund2_file2_text=fund2_file2_extracted_text
                )
            except Exception as e:
                st.error(f"❌ Error generating fund comparison: {e}")
        else:
            st.error("❌ Please upload at least three files for fund comparison.")
    
    if factfinding_text and plan_report_text:
        try:
            # Use the CSV version of the function:
            iht_bullet_points = extract_iht_details_with_gpt(factfinding_text, product_report_text)

        except Exception as e:
            st.error("Error extracting IHT details: " + repr(e))
    else:
        st.warning("Please upload the FactFind and Plan Report files to extract IHT details.")
         

    # Show a warning or success message based on upload completion
    if all(file in st.session_state["uploaded_comparison_files"] for file in REQUIRED_FILES):
        st.success("All required files for the fund comparison table are uploaded! You can proceed.")
    else:
        st.warning("Please upload all required files for the fund comparison table before proceeding.")         

    os.makedirs(OUTPUT_FOLDER, exist_ok=True)
    output_path = os.path.join(OUTPUT_FOLDER, "Generated_Report.docx")

    if st.button("Generate Report", key="generate_button"):
        try:
            st.markdown('<div style="text-align:center;">🛠️ Generating your personalized report...</div>', unsafe_allow_html=True)
            create_new_document(
                template_path=template_path,
                factfinding_text=factfinding_text,
                risk_details=risk_details,
                table_data=plan_report_data,
                product_report_text=product_report_text,
                plan_report_text=plan_report_text,
                fund_performance_data=fund_performance_data,
                dark_star_performance_data=dark_star_performance_data, 
                sap_comparison_table=sap_comparison_table,
                annuity_quotes_text=annuity_quotes_text, 
                fund_comparison_text=fund_comparison_text,
                iht_bullet_points=iht_bullet_points,  # Pass the bullet points instead of CSV
                output_path=output_path
            )
            with open(output_path, "rb") as f:
                st.download_button(
                    label="📥 Download Generated Report",
                    data=f,
                    file_name="Generated_Report.docx",
                    mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
                )
        except Exception as e:
            st.error(f"❌ An error occurred: {e}")
# Footer Section
st.markdown('<div class="footer">Working Hours: Monday to Friday, 9:00 AM – 5:30 PM</div>', unsafe_allow_html=True)
st.markdown(
    """
    <div class="disclaimer">
    Zomi Wealth is a trading name of Holistic Wealth Management Limited, authorized and regulated by the FCA.<br>
    Guidance provided is subject to the UK regulatory regime and is targeted at UK consumers.<br>
    Investments can go down as well as up. Past performance is not indicative of future results.
    </div>
    """,
    unsafe_allow_html=True,
)