import os
import pandas as pd
import logging
import psycopg2
from psycopg2 import sql
from psycopg2.extras import execute_values

logging.basicConfig(level=logging.INFO, format='%(message)s')

header_variations_filename = input('Enter the file name of the CP_Template with Correct Headers (Without extension): ') + '.xlsx'
folder_to_check = input('Enter the folder name that contains all Excel files: ')

current_dir = os.getcwd()
header_variations_file = os.path.join(current_dir, header_variations_filename)
folder_to_check = os.path.join(current_dir, folder_to_check)

header_variations_df = pd.read_excel(header_variations_file)
header_variations = {
    column: header_variations_df[column].iloc[1:].dropna().tolist()
    for column in header_variations_df.columns
}

reverse_lookup = {
    variation: standard_header
    for standard_header, variations in header_variations.items()
    for variation in variations
}

def process_excel_file(file_path):
    if file_path == header_variations_file:
        return

    try:
        df = pd.read_excel(file_path, header=0)
    except Exception as e:
        logging.error(f"Failed to read {file_path}: {e}")
        return

    df.rename(columns=reverse_lookup, inplace=True)
    df = df[[col for col in df.columns if col in header_variations]] 
    
    if not df.empty:
        # logging.info(f"Processed file completed successfully.")
        push_to_database(df)  
    else:
        logging.warning(f"The DataFrame from {file_path} is empty or did not match required headers. No data to append.")

def process_excel_files_in_directory(directory):
    file_count = 0  

    for root, _, files in os.walk(directory):
        for file in files:
            if file.endswith(('.xlsx', '.xls')) and not file.startswith('~$'):
                file_path = os.path.join(root, file)
                logging.info(f"Processing file: {file_path}-----{file_count+1}")
                process_excel_file(file_path)
                file_count += 1

def push_to_database(df):
    connection = None
    try:
        connection = psycopg2.connect(
            host="ec2-34-193-90-221.compute-1.amazonaws.com",
            port="5432",
            user="postgres",
            password="contentx123",
            database="contentx_data"
        )
        cursor = connection.cursor()

        columns = list(df.columns)
        insert_query = sql.SQL("INSERT INTO product_data ({}) VALUES %s").format(
            sql.SQL(', ').join(map(sql.Identifier, columns))
        )

        execute_values(cursor, insert_query, df.itertuples(index=False, name=None))

        connection.commit()
        logging.info("---------Data successfully pushed to the database.----------")

    except Exception as e:
        logging.error(f"Failed to push data to the database: {e}")
    finally:
        if connection:
            cursor.close()
            connection.close()

if __name__ == "__main__":
    process_excel_files_in_directory(folder_to_check)