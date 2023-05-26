import pandas as pd
from tkinter import Tk, Label, Button, filedialog

def process_excel_file(input_file, output_file):
    # Read the original Excel file
    df = pd.read_excel(input_file)

    # Filter out rows where Grantor or Direct Name is 'UNITED STATES OF AMERICA' or 'DEPARTMENT OF THE TREASURY INTERNAL REVENUE SERVICE'
    df = df[(df['Grantor'] != 'UNITED STATES OF AMERICA') & (df['Grantor'] != 'DEPARTMENT OF THE TREASURY INTERNAL REVENUE SERVICE') |
            (df['Direct Name'] != 'UNITED STATES OF AMERICA') & (df['Direct Name'] != 'DEPARTMENT OF THE TREASURY INTERNAL REVENUE SERVICE')]

    # Filter out rows containing 'UTILITIES' in Grantor or Grantee, or Direct Name or Reverse Name
    df = df[~df['Grantor'].str.contains('UTILITIES', case=False) |
            ~df['Grantee'].str.contains('UTILITIES', case=False) |
            ~df['Direct Name'].str.contains('UTILITIES', case=False) |
            ~df['Reverse Name'].str.contains('UTILITIES', case=False)]

    # Filter out rows containing 'CITY OF' in Grantor or Grantee, or Direct Name or Reverse Name
    df = df[~df['Grantor'].str.contains('CITY OF', case=False) |
            ~df['Grantee'].str.contains('CITY OF', case=False) |
            ~df['Direct Name'].str.contains('CITY OF', case=False) |
            ~df['Reverse Name'].str.contains('CITY OF', case=False)]

    # More filters
    df = df[~df['Grantor'].str.contains('INTERNAL REVENUE SERVICE', case=False) |
            ~df['Grantee'].str.contains('INTERNAL REVENUE SERVICE', case=False) |
            ~df['Direct Name'].str.contains('INTERNAL REVENUE SERVICE', case=False) |
            ~df['Reverse Name'].str.contains('INTERNAL REVENUE SERVICE', case=False)]

    df = df[~df['Grantor'].str.contains('DEPARTMENT OF REVENUE', case=False) |
            ~df['Grantee'].str.contains('DEPARTMENT OF REVENUE', case=False) |
            ~df['Direct Name'].str.contains('DEPARTMENT OF REVENUE', case=False) |
            ~df['Reverse Name'].str.contains('DEPARTMENT OF REVENUE', case=False)]

    df = df[~df['Grantor'].str.contains('UTILITY', case=False) |
            ~df['Grantee'].str.contains('UTILITY', case=False) |
            ~df['Direct Name'].str.contains('UTILITY', case=False) |
            ~df['Reverse Name'].str.contains('UTILITY', case=False)]

    # Create a new DataFrame with the Company name and Owner's Name columns
    new_df = pd.DataFrame()
    new_df['Company name'] = df['Grantor'].fillna(df['Direct Name'])

    # Create a new DataFrame with the Company name and Owner's Name columns
    new_df = pd.DataFrame()
    new_df['Company name'] = df['Grantor'].fillna(df['Direct Name'])
    new_df['Mailing Address'] = ''
    new_df['Unit'] = ''
    new_df['City'] = ''
    new_df['State'] = ''
    new_df['Zip'] = ''
    new_df['Owner\'s Name'] = df['Grantee'].fillna(df['Reverse Name'])
    new_df['Owner\'s Mailing Address'] = ''
    new_df['Owner\'s City'] = ''
    new_df['Owner\'s State'] = ''
    new_df['Owner\'s Zip'] = ''

    # Write the new DataFrame to a new Excel file
    new_df.to_excel(output_file, index=False)

def select_input_file():
    root = Tk()
    root.withdraw()
    input_file = filedialog.askopenfilename(filetypes=[('Excel Files', '*.xlsx')])
    return input_file

def select_output_directory():
    root = Tk()
    root.withdraw()
    output_dir = filedialog.askdirectory()
    return output_dir

def process_files():
    input_file = select_input_file()
    output_dir = select_output_directory()
    output_file = f"{output_dir}/new_file.xlsx"
    process_excel_file(input_file, output_file)
    print("Processing completed!")

# Run the file processing when the button is clicked
process_button = Button(text="Process Files", command=process_files)
process_button.pack()

# Start the GUI event loop
Tk().mainloop()

