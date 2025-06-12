Word to PDF Converter
A simple Python script that converts Word documents (.docx) to PDF files using ReportLab.

Features
Interactive command-line interface
Converts .docx files to PDF format
Customizable output location
Automatic font handling
Multi-page support with proper pagination
Prerequisites
Before running this script, make sure you have Python installed on your system. This script requires Python 3.6 or higher.

Installation
Clone this repository:
bash
git clone github repo link
cd word-to-pdf-converter
Install the required dependencies:
bash
pip install -r requirements.txt
Usage
Run the script from the command line:

bash
python cnvrtr.py
The script will guide you through the conversion process:

Input File: Enter the full path to your Word document (.docx file)
Output Location: Choose to save in the default location (same directory as input) or specify a custom path
Conversion: The script will convert your document and save it as a PDF
Example
Word Document to PDF Converter
-----------------------------------
Enter the full path to the Word document: /path/to/your/document.docx

Default PDF save location: /path/to/your/document.pdf
Do you want to save in this location? (yes/no): yes

PDF converted successfully: /path/to/your/document.pdf
Dependencies
python-docx: For reading Word documents
reportlab: For PDF generation
os: Built-in Python module for file operations
File Structure
word-to-pdf-converter/
│
├── cnvrtr.py          # Main conversion script
├── requirements.txt   # Python dependencies
├── README.md         # Project documentation
└── LICENSE           # License file
How It Works
Input Validation: The script validates that the input file exists and is a .docx file
Document Reading: Uses python-docx to extract text content from the Word document
PDF Generation: Creates a PDF using ReportLab with proper formatting and pagination
Font Handling: Automatically registers Arial font for consistent text rendering
Limitations
Currently supports basic text conversion (no images, tables, or complex formatting)
Requires Arial font to be installed on the system
Works with .docx files only (not .doc)
Contributing
Fork the repository
Create a feature branch (git checkout -b feature/new-feature)
Commit your changes (git commit -am 'Add new feature')
Push to the branch (git push origin feature/new-feature)
Create a Pull Request
Future Enhancements
 Support for images and tables
 Preserve original document formatting
 Batch conversion support
 GUI interface
 Support for .doc files
 Custom font selection


 
With Love Xyden

Acknowledgments
ReportLab for PDF generation capabilities
python-docx for Word document processing
