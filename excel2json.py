from os import environ, path
import tkinter as tk
from tkinter import filedialog
import pandas as pd
import json


def main():
    # Get the file to read
    try:
        file_path = environ['FILE_PATH']
    except KeyError:
        # Ask to select a file to open
        root = tk.Tk()
        root.withdraw()
        file_path = filedialog.askopenfilename()
    # Return if file_path is empty
    if not file_path:
        return -1
    # Get file name and file directory from file_path
    file_name, extension = path.splitext(path.basename(file_path))
    file_directory = path.dirname(file_path)
    # Load the file
    file_handler = pd.ExcelFile(file_path)
    sheet_names = file_handler.sheet_names
    # Iterate over sheets and append the json data
    all_data = {}
    for sheet in sheet_names:
        sheet_data = file_handler.parse(sheet)
        # Convert sheet data to JSON
        sheet_json = sheet_data.to_json(date_format='iso', orient='records', indent=4)
        all_data[sheet] = json.loads(sheet_json)

    # Save the data
    out_file_name = path.join(file_directory, file_name + '.json')
    with open(out_file_name, 'w') as destination_handler:
        json_data = json.dumps(all_data)
        destination_handler.write(json_data)
        destination_handler.close()

    return 0


if __name__ == '__main__':
    main()

