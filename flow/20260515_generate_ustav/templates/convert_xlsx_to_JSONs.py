import os
import pandas as pd


def excel_sheets_to_json(excel_file_path):
    """Reads an Excel file and converts each sheet into an individual JSON file

    saved in the same directory.
    """
    # 1. Validate that the file exists
    if not os.path.exists(excel_file_path):
        print(f"Error: The file '{excel_file_path}' does not exist.")
        return

    # Get the directory where the Excel file is located
    output_dir = os.path.dirname(os.path.abspath(excel_file_path))

    try:
        # 2. Load the Excel file (None loads all sheets into a dictionary)
        print(f"Reading {excel_file_path}...")
        excel_data = pd.read_excel(excel_file_path, sheet_name=None)

        # 3. Iterate through each sheet and save as JSON
        for sheet_name, df in excel_data.items():
            # Clean the sheet name to make a safe filename
            safe_sheet_name = "".join(
                c for c in sheet_name if c.isalnum() or c in (" ", "_", "-")
            ).rstrip()
            json_filename = f"{safe_sheet_name}.json"
            json_file_path = os.path.join(output_dir, json_filename)

            # Convert dataframe to JSON
            # 'records' format creates an array of objects (one per row)
            # indent=4 makes the JSON file human-readable
            df.to_json(
                json_file_path, orient="records", indent=4, date_format="iso" , force_ascii=False
            )

            with open(json_file_path, "r", encoding="utf-8") as f:
                content = f.read()

            fixed_content = content.replace(r"\/", "/")

            with open(json_file_path, "w", encoding="utf-8") as f:
                f.write(fixed_content)


            print(f"-> Created: {json_filename}")

        print("\nSuccess! All sheets have been converted.")

    except Exception as e:
        print(f"An error occurred: {e}")


if __name__ == "__main__":

    filename = "everyday JSON.xlsx"


    # Remove surrounding quotes if the user dragged and dropped the file into the terminal
    filename = filename.strip("'\"")

    excel_sheets_to_json(filename)