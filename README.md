# DEXA-Scan-Data

import streamlit as st
import pandas as pd
import pdfplumber
import re
import io
import xlsxwriter

# Function to clean and split the region data
def clean_region_text(region_lines):
    for line in region_lines:
        # Clean the line by removing any unwanted text and splitting by spaces
        cleaned_line = re.sub(r'[^\d.,\- ]+', '', line).strip()
        parts = re.split(r'\s+', cleaned_line)
        if len(parts) >= 6:  # Ensure there are at least 6 columns of data
            return parts
    return []

# Function to extract data for each region
def extract_region_data(region, region_lines):
    clean_lines = clean_region_text(region_lines)
    if clean_lines:
        return {
            f"{region} (%Fat)": clean_lines[0],
            f"{region} Centile": clean_lines[1],
            f"{region} Total Mass (kg)": clean_lines[2],
            f"{region} Fat (g)": clean_lines[3],
            f"{region} Lean (g)": clean_lines[4],
            f"{region} BMC (g)": clean_lines[5]
        }
    return {}

# Main function to extract DEXA data
def extract_dexa_data(pdf_file):
    data_composition = []
    patient_info = {}
    data_trends = []

    with pdfplumber.open(pdf_file) as pdf:
        for page in pdf.pages:
            text = page.extract_text()

            # Extract patient info from the first page
            if not patient_info:
                patient_info = extract_patient_info(text)

            # Extract Total Body Composition
            if "Total Body Tissue Quantitation Composition (Enhanced Analysis)" in text:
                regions = ['Arms', 'Arm Right', 'Arm Left', 'Legs', 'Leg Right', 'Leg Left', 'Trunk', 'Android', 'Gynoid', 'Total']
                for region in regions:
                    # Extract lines matching the region
                    region_lines = re.findall(f"{region}.*", text)
                    if region_lines:
                        region_data = extract_region_data(region, region_lines)
                        data_composition.append(region_data)

    return patient_info, data_composition

# Helper function to extract values using regular expressions
def extract_value_from_text(text, pattern):
    match = re.search(pattern, text)
    if match:
        return match.group(1).replace(",", "")  # Clean up commas in numbers
    return None

# Function to extract patient information
def extract_patient_info(text):
    return {
        "Patient ID": extract_value_from_text(text, r"Patient:\s*(.+?)\s"),
        "Birth Date": extract_value_from_text(text, r"Birth Date:\s*(.+?)\s"),
        "Age": extract_value_from_text(text, r"Age:\s*(.+?)\s"),
        "Height (in)": extract_value_from_text(text, r"Height:\s*(.+?)\sin"),
        "Weight (lbs)": extract_value_from_text(text, r"Weight:\s*(.+?)\slbs"),
        "Sex": extract_value_from_text(text, r"Sex:\s*(.+?)\s"),
        "Ethnicity": extract_value_from_text(text, r"Ethnicity:\s*(.+?)\s"),
        "Measured": extract_value_from_text(text, r"Measured:\s*(.+?)\s"),
        "Analyzed": extract_value_from_text(text, r"Analyzed:\s*(.+?)\s")
    }

# Streamlit UI
st.title('DEXA Scan Data Extraction to Excel')

# File upload option (Allow multiple files)
uploaded_files = st.file_uploader("Upload DEXA Scan PDFs", type="pdf", accept_multiple_files=True)

# Initialize output for storing data from all files
all_data = []

if uploaded_files:
    for uploaded_file in uploaded_files:
        # Extract data for each uploaded PDF
        patient_info, data_composition = extract_dexa_data(uploaded_file)

        # Combine patient info and body composition data into one row
        patient_row = {**patient_info}  # Start with patient info
        for data in data_composition:
            patient_row.update(data)  # Add each region's data to the row

        # Convert the row into a DataFrame for each patient
        df_patient = pd.DataFrame([patient_row])
        all_data.append(df_patient)

    # Concatenate all patient data into one DataFrame
    if all_data:
        df_all_patients = pd.concat(all_data, ignore_index=True)

        # Save to Excel in memory with custom formatting
        output = io.BytesIO()
        with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
            workbook = writer.book
            
            # Formatting options
            bold_format = workbook.add_format({'bold': True, 'align': 'center'})
            header_format = workbook.add_format({'bold': True, 'bg_color': '#F9F9F9', 'align': 'center', 'border': 1})
            centered_format = workbook.add_format({'align': 'center'})
            
            # Write patient info section
            df_patient_info = df_all_patients[['Patient ID', 'Birth Date', 'Age', 'Height (in)', 'Weight (lbs)', 'Sex', 'Ethnicity', 'Measured', 'Analyzed']]
            df_patient_info.to_excel(writer, sheet_name="DEXA Data", startrow=0, index=False, header=True)
            
            worksheet = writer.sheets['DEXA Data']
            
            # Apply formatting to patient info section
            for col_num, value in enumerate(df_patient_info.columns.values):
                worksheet.write(0, col_num, value, header_format)
            
            # Leave a blank row before the next section
            start_row = len(df_patient_info) + 2

            # Define the regions to write and their column groupings
            regions = {
                "ARMS": ['Arms (%Fat)', 'Arms Centile', 'Arms Total Mass (kg)', 'Arms Fat (g)', 'Arms Lean (g)', 'Arms BMC (g)'],
                "RIGHT ARM": ['Arm Right (%Fat)', 'Arm Right Centile', 'Arm Right Total Mass (kg)', 'Arm Right Fat (g)', 'Arm Right Lean (g)', 'Arm Right BMC (g)'],
                "LEFT ARM": ['Arm Left (%Fat)', 'Arm Left Centile', 'Arm Left Total Mass (kg)', 'Arm Left Fat (g)', 'Arm Left Lean (g)', 'Arm Left BMC (g)'],
                "LEGS": ['Legs (%Fat)', 'Legs Centile', 'Legs Total Mass (kg)', 'Legs Fat (g)', 'Legs Lean (g)', 'Legs BMC (g)'],
                "RIGHT LEG": ['Leg Right (%Fat)', 'Leg Right Centile', 'Leg Right Total Mass (kg)', 'Leg Right Fat (g)', 'Leg Right Lean (g)', 'Leg Right BMC (g)'],
                "LEFT LEG": ['Leg Left (%Fat)', 'Leg Left Centile', 'Leg Left Total Mass (kg)', 'Leg Left Fat (g)', 'Leg Left Lean (g)', 'Leg Left BMC (g)'],
                "TRUNK": ['Trunk (%Fat)', 'Trunk Centile', 'Trunk Total Mass (kg)', 'Trunk Fat (g)', 'Trunk Lean (g)', 'Trunk BMC (g)'],
                "ANDROID": ['Android (%Fat)', 'Android Centile', 'Android Total Mass (kg)', 'Android Fat (g)', 'Android Lean (g)', 'Android BMC (g)'],
                "GYNOID": ['Gynoid (%Fat)', 'Gynoid Centile', 'Gynoid Total Mass (kg)', 'Gynoid Fat (g)', 'Gynoid Lean (g)', 'Gynoid BMC (g)'],
                "TOTAL": ['Total (%Fat)', 'Total Centile', 'Total Total Mass (kg)', 'Total Fat (g)', 'Total Lean (g)', 'Total BMC (g)']
            }

            # Write each region's data, adding spacing between them
            for region_name, columns in regions.items():
                worksheet.write(start_row, 0, region_name, bold_format)
                df_region = df_all_patients[['Patient ID'] + columns]
                df_region.to_excel(writer, sheet_name="DEXA Data", startrow=start_row + 1, index=False, header=True)

                # Apply formatting to each region's section
                for col_num, value in enumerate(df_region.columns.values):
                    worksheet.write(start_row + 1, col_num, value, header_format)

                # Add space between sections
                start_row += len(df_region) + 2

        # Retrieve the processed data
        processed_data = output.getvalue()

        # Provide download option for the Excel file
        st.download_button(
            label="Download All Data as Excel",
            data=processed_data,
            file_name="dexa_scan_data_formatted.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        # Display data preview
        st.write("Extracted Data Preview")
       
