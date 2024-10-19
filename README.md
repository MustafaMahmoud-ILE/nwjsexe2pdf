# Executable to PDF Converter

## Overview
This Python application converts executable e-book files into PDF documents. Designed to address the issue of universities switching course materials from PDFs to EXE files, this tool helps users, such as students, convert their course e-books to a printable format. The app parses the content within the EXE files, extracts XHTML content, processes it, and generates PDFs.

## Features
- **Convert EXE to PDF**: Select one or multiple EXE files and convert them into PDFs.
- **HTML/CSS/JavaScript Support**: Handles external CSS and JavaScript files referenced within the XHTML.
- **Progress Bar**: Displays conversion progress.
- **Total Pages and Pricing**: Estimates total printing price based on the number of pages processed.
- **Dark Theme Interface**: Cool, user-friendly, dark-themed GUI.

## Installation
1. Make sure you have the following prerequisites:
   - Python 3.x
   - PyQt5
   - pdfkit
   - wkhtmltopdf
   - BeautifulSoup4
   - PyPDF2

2. Clone this repository:
   ```bash
   git clone https://github.com/MustafaMahmoud-ILE/nwjsexe2pdf.git
   cd nwjsexe2pdf
   ```

3. Install the required Python packages:
   ```bash
   pip install -r requirements.txt
   ```

4. Make sure `wkhtmltopdf` is installed. You can download it from [here](https://wkhtmltopdf.org/downloads.html).

## Usage
1. Launch the application:
   ```bash
   python nwjsexe2pdf.py
   ```
2. Click the "Browse" button to select EXE files.
3. Click "Convert" to start the process.
4. The output PDFs will be saved in the same directory as the selected EXE files.

## Pricing Calculation
The app estimates the printing cost based on the number of pages. A base price is added for each page, and additional costs are scaled down with each subsequent page. The final price is rounded to the nearest 50 EGP.

## Screenshots
_Add screenshots of the app here._

## Contributing
Feel free to fork this repository, submit issues, or make pull requests. Contributions are welcome!

## License
This project is licensed under the MIT License - see the LICENSE file for details.
