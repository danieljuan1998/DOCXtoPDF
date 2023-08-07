markdown

# Word to PDF Converter

This script is a simple utility for converting Word documents (.doc and .docx) to PDF format using the `win32com` library in Python.

## Prerequisites

- Python 3.x
- `win32com.client` library (usually bundled with the `pywin32` package)

## Installation

1. Clone or download this repository to your local machine.
2. Make sure you have Python 3.x installed. If not, you can download it from the official [Python website](https://www.python.org/downloads/).
3. Install the `pywin32` package by running the following command: pip install pywin32



## Usage

1. Open a terminal or command prompt.
2. Navigate to the directory where you cloned/downloaded the repository.
3. Modify the script's `os.chdir` line to specify the directory containing the Word files you want to convert:

```python
os.chdir(r'C:\Desktop')

Replace 'C:\Desktop' with the actual path to your target directory.

    Run the script:

    python word_to_pdf_converter.py

    The script will iterate through all the .doc and .docx files in the specified directory, convert them to PDF format, and save them in the same location.

Notes

    Make sure Microsoft Word is installed on your system.
    The converted PDF files will have the same name as the original Word files, with the .pdf extension.
    You can customize the script to handle different file paths or formats if needed.
