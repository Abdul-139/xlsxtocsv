# file: xlsx_to_csv_converter.py

import pandas as pd
import os

def convert_xlsx_to_csv(xlsx_file_path, csv_directory_path):
    """
    Converts all XLSX files in a specified directory to CSV files.
    
    Parameters:
    - xlsx_file_path: str, path to the input XLSX directory.
    - csv_directory_path: str, path to the output CSV directory.
    """
    try:
        # Ensure the CSV directory exists
        if not os.path.exists(csv_directory_path):
            os.makedirs(csv_directory_path)

        # List all XLSX files in the specified directory
        xlsx_files = [f for f in os.listdir(xlsx_file_path) if f.endswith('.xlsx')]
        
        for xlsx_file in xlsx_files:
            xlsx_full_path = os.path.join(xlsx_file_path, xlsx_file)
            csv_file_name = xlsx_file.replace('.xlsx', '.csv')
            csv_full_path = os.path.join(csv_directory_path, csv_file_name)
            
            # Load the XLSX file
            df = pd.read_excel(xlsx_full_path)
            
            # Save the data to a CSV file
            df.to_csv(csv_full_path, index=False)
            
            print(f"File {xlsx_file} converted successfully and saved to {csv_full_path}")
            
    except Exception as e:
        print(f"An error occurred: {e}")

# Example usage:
xlsx_file_path = r'C:\Users\Abdul\OneDrive - Chulalongkorn University\fadhli nitip\xlsx'
csv_directory_path = r'C:\Users\Abdul\OneDrive - Chulalongkorn University\fadhli nitip\csv'

convert_xlsx_to_csv(xlsx_file_path, csv_directory_path)
