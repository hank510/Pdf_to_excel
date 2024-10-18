import pandas as pd
import PyPDF2
import re

def extract_table_data_from_pdf(pdf_path, form_number):
    data = []
    with open(pdf_path, 'rb') as file:
        reader = PyPDF2.PdfReader(file)
        for page in reader.pages:
            text = page.extract_text()
            # Extract the data for the specified form number
            if form_number in text:
                # Assuming the table starts after the form number and includes multiple lines
                # Split text into lines
                lines = text.split('\n')
                start_extraction = False
                for line in lines:
                    # Start extracting once the form number is found
                    if form_number in line:
                        start_extraction = True
                    if start_extraction:
                        # Break if we reach the next form or an empty line
                        if re.match(r'^[\d\s]*$', line) or 'FORM' in line:
                            break
                        # Process and store the line data
                        processed_line = [item.strip() for item in line.split() if item.strip()]
                        if processed_line:
                            data.append(processed_line)
    return data

def create_excel_from_data(form_data, output_file):
    # Create a DataFrame
    df = pd.DataFrame(form_data)
    # Save the DataFrame to Excel
    df.to_excel(output_file, index=False, header=False)

# Paths
pdf_path = 'HDFC Q1 S 2023-2024.pdf'
output_file = 'HDFC_Report.xlsx'

# Extract data from specific forms
l1_data = extract_table_data_from_pdf(pdf_path, 'L-1-A-RA')
l2_data = extract_table_data_from_pdf(pdf_path, 'L-2-A-PL')

# Combine data from both forms
combined_data = l1_data + l2_data

# Create an Excel file with the combined data
create_excel_from_data(combined_data, output_file)

print(f'Data extracted and saved to {output_file}')
