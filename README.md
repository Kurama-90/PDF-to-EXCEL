# PDF to Excel Converter with Table and Text Extraction

This Python script extracts both text and tables from a PDF file and exports the content to an Excel file. It uses the `PyMuPDF` (fitz) library to read and process the PDF, and the `pandas` library to create an Excel file.

## Features
- **Text Extraction**: Extracts raw text from PDF pages excluding the text within tables.
- **Table Extraction**: Detects and extracts tables along with their coordinates.
- **Sorting**: The extracted content is sorted in the order it appears on the page.
- **Excel Export**: Combines both text and table data into a single Excel sheet.

## Requirements
Before using the script, you need to have Python installed, along with the following libraries:

- `PyMuPDF` (for PDF processing)
- `pandas` (for creating Excel files)
- `xlsxwriter` (for writing Excel files)

To install the required libraries, run the following command:

```bash
pip install pymupdf pandas xlsxwriter
```

## How It Works

### Extract Text and Tables:
- The script processes each page of the PDF.
- Text blocks are extracted excluding any text inside tables.
- Tables are extracted using the `find_tables()` method, and their coordinates are also recorded.
- Both text and table data are stored and sorted based on the order they appear on the page.

### Export to Excel:
- The extracted data (both text and tables) is written into an Excel file.
- Tables are added as rows, and the text is added in the order it appears.

## Usage

1. Clone the repository to your local machine:

    ```bash
    git clone https://github.com/Kurama-90/PDF-to-EXCEL.git
    cd PDF-to-EXCEL
    ```

2. Ensure your PDFs are placed in the `uploads` folder.

3. Run the script:

    ```bash
    python main.py
    ```

4. The script will ask you to enter the PDF filename you want to process (ensure it has the `.pdf` extension).

5. After the PDF is processed, the output Excel file will be saved in the `outputs` folder.

## File Structure
```
📂 project-directory
 ├── 📂 uploads  # Place scanned PDFs here
 ├── 📂 outputs  # Processed files will be saved here
 ├── main.py  # Main script for processing PDFs
 ├── README.md  # Project documentation
 ├── LICENSE  # MIT License file
```

## Contributing
Feel free to contribute by submitting issues or pull requests.

## License
This project is licensed under the MIT License.

## Author
[Kurama-90](https://github.com/Kurama-90)

