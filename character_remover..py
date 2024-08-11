import re
from docx import Document

def clean_text(text):
    # Define the pattern for special characters to be removed
    pattern = r'[\*\#\@\!\$\%\^\&\(\)\[\]\{\}\<\>\|\?\/\\]'
    
    # Remove the special characters using regex
    cleaned_text = re.sub(pattern, '', text)
    
    return cleaned_text

def process_word_document(input_path, output_path):
    # Load the Word document
    doc = Document(input_path)
    
    # Iterate through each paragraph in the document
    for para in doc.paragraphs:
        # Clean the text in the paragraph
        para.text = clean_text(para.text)
    
    # Save the cleaned document
    doc.save(output_path)
    print(f"Processed document saved as {output_path}")

if __name__ == "__main__":
    input_file = "srsreport.docx"  # Specify the input file name
    output_file = "srsreportcleaned.docx"  # Specify the output file name
    
    process_word_document(input_file, output_file)
