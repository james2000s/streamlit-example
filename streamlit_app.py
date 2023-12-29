# Author:   Carlos Rodriguez
# Project:  FileTypeTransfer
# Date:     12/25/2023
# Purpose:  Convert between different file types

import tkinter as tk
from tkinter import filedialog
import subprocess
import re
import sys

# Forward declarations
file_path, file_dir = '', ''


def convert_to(folder, source, timeout=None):
    try:
        args = [libreoffice_exec(), '--headless', '--convert-to', 'pdf', '--outdir', folder, source]

        process = subprocess.run(args, stdout=subprocess.PIPE, stderr=subprocess.PIPE, timeout=timeout)
        filename = re.search('-> (.*?) using filter', process.stdout.decode())

        if filename is None:
            raise Exception
        else:
            print('else, ' + filename.group(1))
            return filename.group(1)
    except Exception as e:
        print(e)


def libreoffice_exec():
    # TODO: Provide support for more platforms
    if sys.platform == 'win32':
        return 'C:\\Program Files\\LibreOffice\\program\\swriter.exe'
    return 'libreoffice'


def open_file_dialog():
    global file_path, file_dir

    root = tk.Tk()
    root.withdraw()  # Hide the main window

    file_path = filedialog.askopenfilename()  # Show the file dialog and return the selected file's path
    file_dir = file_path[:file_path.rfind('/')+1]
    if file_path:
        print("Path of Directory: " + file_dir)
        print("Selected file: ", file_path)
        # Do something with the selected file path, such as further processing or opening the file


def main():
    global file_path, file_dir
    # Call the function to open the file dialog
    open_file_dialog()
    try:
        convert_to(file_dir, file_path)
    except Exception as e:
        print(e)


main()

