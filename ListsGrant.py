import pandas as pd

def process_excel_file(input_file, output_file):
    # Read the original Excel file
    df = pd.read_excel(input_file)

    # Filter out rows where Grantor is 'UNITED STATES OF AMERICA' or 'DEPARTMENT OF THE TREASURY INTERNAL REVENUE SERVICE'
    df = df[(df['Grantor'] != 'UNITED STATES OF AMERICA') & (df['Grantor'] != 'DEPARTMENT OF THE TREASURY INTERNAL REVENUE SERVICE')]

    # Filter out rows containing 'UTILITIES' in Grantor or Grantee
    df = df[~df['Grantor'].str.contains('UTILITIES', case=False)]
    df = df[~df['Grantee'].str.contains('UTILITIES', case=False)]

    # Filter out rows containing 'CITY OF' in Grantor or Grantee
    df = df[~df['Grantor'].str.contains('CITY OF', case=False)]
    df = df[~df['Grantee'].str.contains('CITY OF', case=False)]

    # More filter
    df = df[~df['Grantor'].str.contains('INTERNAL REVENUE SERVICE', case=False)]
    df = df[~df['Grantee'].str.contains('INTERNAL REVENUE SERVICE', case=False)]

    df = df[~df['Grantor'].str.contains('DEPARTMENT OF REVENUE', case=False)]
    df = df[~df['Grantee'].str.contains('DEPARTMENT OF REVENUE', case=False)]

    df = df[~df['Grantor'].str.contains('UTILITY', case=False)]
    df = df[~df['Grantee'].str.contains('UTILITY', case=False)]

    df = df[~df['Grantor'].str.contains('STATE OF', case=False)]
    df = df[~df['Grantee'].str.contains('STATE OF', case=False)]

    # Create a new DataFrame with the Company name and Owner's Name columns
    new_df = pd.DataFrame()
    new_df['Company name'] = df['Grantor']

    # Create the other columns with empty values
    new_df['Mailing Address'] = ''
    new_df['Unit'] = ''
    new_df['City'] = ''
    new_df['State'] = ''
    new_df['Zip'] = ''
    new_df['Owner\'s Name'] = df['Grantee']
    new_df['Owner\'s Mailing Address'] = ''
    new_df['Owner\'s City'] = ''
    new_df['Owner\'s State'] = ''
    new_df['Owner\'s Zip'] = ''

    # Write the new DataFrame to a new Excel file
    new_df.to_excel(output_file, index=False)

# Specify the paths for the input and output Excel files
input_file_path = '/home/xian/Documents/Lien Work/May-2023/Okaloosa/_ExportResults_2023_05_25 14_17_47.xlsx'
output_file_path = '/home/xian/Documents/Lien Work/May-2023/Okaloosa/5-1_5-25/5-1_5-25.xlsx'

# Call the function to process the Excel file
process_excel_file(input_file_path, output_file_path)
