import os
import platform
import re
import subprocess
import pandas as pd


def convert_to_excel(dat_file_path):
    try:
        if os.path.exists(dat_file_path) and dat_file_path.lower().endswith(".dat"):
            with open(dat_file_path, 'r', encoding='utf-8') as file:
                content = file.read()

            # Remove the weird characters
            clean_content = re.sub(r'þ', '', content)
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


'''
if __name__ == "__main__":
    while True:
        # Prompt the user to enter the path of the .dat file
        dat_file_path = input("Enter the path of the .dat file: ")

        # Validate if the file exists and if it has a .dat extension
        if os.path.exists(dat_file_path) and dat_file_path.lower().endswith(".dat"):
            convert_to_excel(dat_file_path)
        else:
            print(
                "Invalid file path or file extension. Please provide a valid .dat file.")
            continue

        # Ask if the user wants to convert another file
        answer = input(
            "Do you want to convert another file? (Enter 'yes' to continue or anything else to end): ").lower()
        if answer != 'yes':
            break

    print("Thank you for using the converter. Goodbye!")
'''
