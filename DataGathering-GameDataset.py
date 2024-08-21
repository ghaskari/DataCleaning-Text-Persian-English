import os
import pandas as pd


# Function to convert each sheet in an Excel file and append to a CSV
def convert_excel_to_csv(file_path, output_csv):
    # Load the Excel file
    xls = pd.ExcelFile(file_path)

    # Extract the base name of the file (without directory and extension)
    base_name = os.path.splitext(os.path.basename(file_path))[0]

    # Iterate through each sheet in the Excel file
    for sheet_name in xls.sheet_names:
        # Load the sheet into a DataFrame
        df = pd.read_excel(xls, sheet_name=sheet_name)

        # Ensure the DataFrame contains the required columns
        if 'انگلیسی' in df.columns and 'فارسی' in df.columns:
            # Select only the "انگلیسی" (English) and "فارسی" (Farsi) columns
            df = df[['انگلیسی', 'فارسی']]

            # Rename columns to English for the output CSV
            df.columns = ['English', 'Farsi']
            print(df.head())

            # If the output CSV file does not exist, create it with the header
            if not os.path.isfile(output_csv):
                df.to_csv(output_csv, index=False, encoding='utf-8')
            else:  # Otherwise, append to the CSV file without the header
                df.to_csv(output_csv, index=False, encoding='utf-8')


# Function to process all Excel files in a directory and combine into one CSV
def process_directory(input_dir, output_csv):
    # Iterate through each file in the input directory
    for file_name in os.listdir(input_dir):
        # Check if the file is an Excel file
        if file_name.endswith('.xlsx') or file_name.endswith('.xls'):
            # Full path to the input file
            file_path = os.path.join(input_dir, file_name)
            # Process the file and append to the CSV
            convert_excel_to_csv(file_path, output_csv)


# Example usage
input_directory = os.path.abspath('Dataset_Game')  # Replace with your input directory path
output_csv = os.path.abspath('DataSource/Combined_Game_Dataset.csv')  # Output CSV file

process_directory(input_directory, output_csv)
