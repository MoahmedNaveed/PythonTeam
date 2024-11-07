import os
import pandas as pd
import logging

logging.basicConfig(level=logging.INFO)

header_variations_filename = input('Enter the file name of the CP_Template with Correct Headers (Without extension): ')+'.xlsx'
folder_to_check = input('Enter the folder name that contains all Excel files: ')

current_dir = os.getcwd()
header_variations_file = os.path.join(current_dir, header_variations_filename)
folder_to_check = os.path.join(current_dir, folder_to_check)

header_variations_df = pd.read_excel(header_variations_file)
header_variations = {column: header_variations_df[column].iloc[1:].dropna().tolist() for column in header_variations_df.columns}

all_dataframes = []

def process_excel_file(file_path):
    if file_path == header_variations_file:
        return

    try:
        df = pd.read_excel(file_path, header=0)
    except Exception as e:
        logging.error(f"Failed to read {file_path}: {e}")
        return

    matched_columns = [col for col in df.columns if any(col in variations for variations in header_variations.values())]
    if not matched_columns:
        logging.warning(f"No matching headers found in {file_path}. Skipping file.")
        return

    for standard_header, variations in header_variations.items():
        for col in df.columns:
            if col in variations:
                df.rename(columns={col: standard_header}, inplace=True)

    df = df[[col for col in df.columns if col in header_variations]]

    if not df.empty:
        all_dataframes.append(df)
        logging.info(f"Processed file: {file_path} successfully.")
    else:
        logging.warning(f"The DataFrame from {file_path} is empty or did not match required headers. No data to append.")

def process_excel_files_in_directory(directory):
    for root, _, files in os.walk(directory):
        for file in files:
            if file.endswith(('.xlsx', '.xls')) and not file.startswith('~$'):
                file_path = os.path.join(root, file)
                logging.info(f"Processing file: {file_path}")
                process_excel_file(file_path)

if __name__ == "__main__":
    process_excel_files_in_directory(folder_to_check)

    if all_dataframes:
        combined_df = pd.concat(all_dataframes, ignore_index=True).drop_duplicates()
        combined_df = combined_df.reindex(columns=header_variations.keys())

        output_file = 'combined_output_data.xlsx'
        try:
            combined_df.to_excel(output_file, index=False)
            print(f"All data written successfully to {output_file}.")
        except PermissionError:
            print(f"Permission denied: The file '{output_file}' might be open. Please close it and press Enter to try again.")
            input("Press Enter once the file is closed...")
            try:
                combined_df.to_excel(output_file, index=False)
                print(f"All data written successfully to {output_file} after retry.")
            except PermissionError:
                print(f"Failed again: Please make sure '{output_file}' is closed and try rerunning the program.")
            except Exception as e:
                print(f"An unexpected error occurred while writing to Excel after retry: {e}")
        except Exception as e:
            print(f"An unexpected error occurred while writing to Excel: {e}")
    else:
        logging.warning("No data to write. No valid Excel files were processed.")
####################################################################################################################









# import os
# import pandas as pd
# import logging

# logging.basicConfig(level=logging.INFO)

# header_variations_file = input('File path of the CP_Template with Correct Headers at the end with file name: ')

# # Step 1: Load the header variations file, use the header row and ignore rows below it
# header_variations_df = pd.read_excel(header_variations_file)
# header_variations = {column: header_variations_df[column].iloc[1:].dropna().tolist() for column in header_variations_df.columns}

# all_dataframes = []

# def process_excel_file(file_path):
#     # Skip the header file itself
#     if file_path == header_variations_file:
#         return

#     try:
#         # Read each data file and apply the header mapping
#         df = pd.read_excel(file_path, header=0)
#     except Exception as e:
#         logging.error(f"Failed to read {file_path}: {e}")
#         return

#     # Check for matching columns in other input files
#     matched_columns = [col for col in df.columns if any(col in variations for variations in header_variations.values())]
#     if not matched_columns:
#         logging.warning(f"No matching headers found in {file_path}. Skipping file.")
#         return

#     # Rename columns based on the header mapping
#     for standard_header, variations in header_variations.items():
#         for col in df.columns:
#             if col in variations:
#                 df.rename(columns={col: standard_header}, inplace=True)

#     # Retain only columns that match the standard headers in the final output
#     df = df[[col for col in df.columns if col in header_variations]]

#     if not df.empty:
#         all_dataframes.append(df)
#         logging.info(f"Processed file: {file_path} successfully.")
#     else:
#         logging.warning(f"The DataFrame from {file_path} is empty or did not match required headers. No data to append.")

# def process_excel_files_in_directory(directory):
#     for root, _, files in os.walk(directory):
#         for file in files:
#             if file.endswith(('.xlsx', '.xls')) and not file.startswith('~$'):
#                 file_path = os.path.join(root, file)
#                 logging.info(f"Processing file: {file_path}")
#                 process_excel_file(file_path)

# if __name__ == "__main__":
#     folder_to_check = input('Excel file folder path which contains all files: ')
#     process_excel_files_in_directory(folder_to_check)

#     if all_dataframes:
#         # Concatenate all data and align with headers
#         combined_df = pd.concat(all_dataframes, ignore_index=True).drop_duplicates()
#         combined_df = combined_df.reindex(columns=header_variations.keys())

#         output_file = 'combined_output_data.xlsx'
#         try:
#             combined_df.to_excel(output_file, index=False)
#             print(f"All data written successfully to {output_file}.")
#         except PermissionError:
#             print(f"Permission denied: The file '{output_file}' might be open. Please close it and press Enter to try again.")
#             input("Press Enter once the file is closed...")
#             try:
#                 combined_df.to_excel(output_file, index=False)
#                 print(f"All data written successfully to {output_file} after retry.")
#             except PermissionError:
#                 print(f"Failed again: Please make sure '{output_file}' is closed and try rerunning the program.")
#             except Exception as e:
#                 print(f"An unexpected error occurred while writing to Excel after retry: {e}")
#         except Exception as e:
#             print(f"An unexpected error occurred while writing to Excel: {e}")
#     else:
#         logging.warning("No data to write. No valid Excel files were processed.")


