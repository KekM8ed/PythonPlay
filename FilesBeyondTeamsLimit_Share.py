import os
import pandas as pd
from openpyxl import load_workbook

# Function to list files and calculate adjusted lengths and space counts
def list_files_with_flags(directory_path):
    files_list = []
    # Walk through the directory tree
    for root, dirs, files in os.walk(directory_path):
        for file in files:
            full_path = os.path.join(root, file)  # Get the full file path
            path_length = len(full_path)  # Calculate the length of the full file path
            space_count = full_path.count(' ')  # Count the number of spaces in the file path
            adjusted_length = path_length + (space_count * 2)  # Calculate the adjusted length
            is_exceeding = adjusted_length > 220  # Check if the adjusted length exceeds 220
            # Add the details to the list
            files_list.append({
                'full_path': full_path,
                'adjusted_length': adjusted_length,
                'spaces': space_count,
                'exceeds_220': is_exceeding
            })
    return files_list

# Function to save the file details to an Excel file
def save_to_excel(files_with_details, directory_path):
    # Define the path to the desktop and the Excel file
    desktop_path = os.path.join(os.path.expanduser("~"), "Desktop")
    excel_file_path = os.path.join(desktop_path, "file_paths.xlsx")
    
    # Create a DataFrame from the list of file details
    df = pd.DataFrame(files_with_details)
    df['exceeds_220'] = df['exceeds_220'].apply(lambda x: "Yes" if x else "No")  # Convert the exceeds_220 flag to "Yes" or "No"
    df = df[['full_path', 'adjusted_length', 'exceeds_220']]  # Select the columns to save
    
    # Extract the last part of the directory path to use as the sheet name
    sheet_name = os.path.basename(os.path.normpath(directory_path))
    
    # Check if the Excel file already exists
    if os.path.exists(excel_file_path):
        book = load_workbook(excel_file_path)  # Load the existing Excel file
        with pd.ExcelWriter(excel_file_path, engine='openpyxl') as writer:
            writer.book = book
            df.to_excel(writer, sheet_name=sheet_name, index=False)  # Add a new sheet with the file details
    else:
        with pd.ExcelWriter(excel_file_path, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name=sheet_name, index=False)  # Create a new Excel file with the file details

if __name__ == "__main__":
    # Prompt the user to enter the directory path
    directory_path = input("Enter the directory path: ")
    
    # Get the list of files with details
    files_with_details = list_files_with_flags(directory_path)
    # Sort the list by adjusted length in descending order
    files_with_details.sort(key=lambda x: x['adjusted_length'], reverse=True)
    
    # Print the details of each file
    for details in files_with_details:
        flag = " (exceeds 220 characters)" if details['exceeds_220'] else ""
        print(f"{details['full_path']}: {details['adjusted_length']} adjusted length, {details['spaces']} spaces{flag}")
    
    # Save the file details to an Excel file on the desktop
    save_to_excel(files_with_details, directory_path)
    print("File paths saved to Excel on your desktop.")
