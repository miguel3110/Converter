import os
import re
import subprocess
import pandas as pd


def convert_to_excel(dat_file_path):
    try:
        if os.path.exists(dat_file_path) and dat_file_path.lower().endswith(".dat"):
            with open(dat_file_path, 'r', encoding='utf-8', errors='ignore') as file:
                content = file.read()
            clean_content = content
            # Remove the weird characters
            if (clean_content[0] == '�'):
                clean_content = re.sub(r'�', '', clean_content)
            elif (clean_content[0] == 'þ'):
                clean_content = re.sub(r'þ', '', clean_content)
            clean_content = re.sub(r'[^\x0A\x20-\x7E]', ',', clean_content)
            # Split the cleaned content into lines based on newline character (x0A)
            lines = clean_content.split('\n')

            # Extract the header row (first line)
            header = lines[0]

            # Remove the header row from the list of lines
            lines = lines[1:]

            # Split each remaining line into columns based on the comma separator
            data = [line.split(',') for line in lines]

            # Create a DataFrame from the data with the extracted header as the column names
            df = pd.DataFrame(data, columns=header.split(','))

            # Drop completely empty columns
            df = df.dropna(axis=1, how='all')

            # so here we are creating the variable that will tell tell the df
            # the name o of the output file and the directory where it will be saved
            # plottwist name and directory are the same as the input file
            output_excel_file_path = os.path.join(
                os.path.dirname(dat_file_path), f"{os.path.splitext(os.path.basename(dat_file_path))[0]}.xlsx")

            # Write the DataFrame to the Excel file
            df.to_excel(output_excel_file_path, index=False)

            return f"{os.path.dirname(dat_file_path)}"+f"\{os.path.splitext(os.path.basename(dat_file_path))[0]}.xlsx"
        else:
            return "Invalid file path or file extension. Please select a valid .dat file."
    except Exception as e:
        return f"An error ocurred: {e}"


def open_excel_file(excel_file_path):
    if os.path.exists(excel_file_path):
        try:
            command = ['start', excel_file_path]
            subprocess.Popen(command, shell=True)
        except Exception as e:
            return f"An error occurred while opening the file: {e}"
    else:
        return "Invalid file path or file extension. Please select a valid .dat file."
