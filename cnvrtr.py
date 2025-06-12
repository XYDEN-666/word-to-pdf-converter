import os
import docx
from reportlab.pdfgen import canvas
from reportlab.lib.pagesizes import letter
from reportlab.lib.units import inch
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

def get_input_file():
    """
    Prompt user to enter the input Word document path.
    
    :return: Path to the input Word document
    """
    while True:
        input_path = input("Enter the full path to the Word document: ").strip()
        
        # Expand any user path shortcuts and normalize the path
        input_path = os.path.expanduser(input_path)
        
        # Check if file exists and is a .docx
        if os.path.exists(input_path) and input_path.lower().endswith('.docx'):
            return input_path
        else:
            print("Error: File does not exist or is not a Word document (.docx). Please try again.")

def get_output_path(input_path):
    """
    Ask user if they want to use the default output path or choose a custom one.
    
    :param input_path: Path of the input Word document
    :return: Path for the output PDF
    """
    # Get the directory of the input file
    default_dir = os.path.dirname(input_path)
    
    # Generate default output filename
    default_filename = os.path.splitext(os.path.basename(input_path))[0] + '.pdf'
    default_output_path = os.path.join(default_dir, default_filename)
    
    print(f"\nDefault PDF save location: {default_output_path}")
    
    while True:
        choice = input("Do you want to save in this location? (yes/no): ").lower().strip()
        
        if choice in ['yes', 'y']:
            return default_output_path
        elif choice in ['no', 'n']:
            custom_path = input("Enter the full path where you want to save the PDF: ").strip()
            
            # Expand any user path shortcuts and normalize the path
            custom_path = os.path.expanduser(custom_path)
            
            # Ensure the directory exists
            os.makedirs(custom_path, exist_ok=True)
            
            # Generate filename in the custom directory
            custom_output_path = os.path.join(custom_path, default_filename)
            
            return custom_output_path
        else:
            print("Invalid input. Please enter 'yes' or 'no'.")

def convert_docx_to_pdf(input_docx, output_pdf):
    """
    Convert a Word document to a PDF file.
    
    :param input_docx: Path to the input Word document
    :param output_pdf: Path to save the output PDF file
    """
    try:
        # Register a standard font
        pdfmetrics.registerFont(TTFont('Arial', 'Arial.ttf'))
        
        # Read the Word document
        doc = docx.Document(input_docx)
        
        # Create a new PDF with Reportlab
        c = canvas.Canvas(output_pdf, pagesize=letter)
        width, height = letter
        
        # Set font and initial position
        c.setFont('Arial', 11)
        text_object = c.beginText(inch, height - inch)
        
        # Iterate through paragraphs in the document
        for para in doc.paragraphs:
            # Check if we need a new page
            if text_object.getY() < inch:
                c.drawText(text_object)
                c.showPage()
                text_object = c.beginText(inch, height - inch)
                c.setFont('Arial', 11)
            
            # Add paragraph text
            text_object.textLine(para.text)
        
        # Draw any remaining text
        if text_object.getY() < height:
            c.drawText(text_object)
        
        # Save the PDF
        c.save()
        print(f"PDF converted successfully: {output_pdf}")
    
    except Exception as e:
        print(f"An error occurred during conversion: {e}")

def main():
    """
    Main function to orchestrate the Word to PDF conversion process.
    """
    print("Word Document to PDF Converter")
    print("-" * 35)
    
    # Get input file path
    input_file = get_input_file()
    
    # Get output file path
    output_file = get_output_path(input_file)
    
    # Convert the document
    convert_docx_to_pdf(input_file, output_file)

if __name__ == "__main__":
    main()