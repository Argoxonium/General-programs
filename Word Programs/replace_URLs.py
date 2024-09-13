import re
import os
from docx import Document
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

def main() -> None:
    #Get the path and open the word document
    path = input('What is the path location for the word document')
    doc = open_word_document(path)
    hyperlinks = extract_hyperlinks(doc)
    print(hyperlinks)

def open_word_document(file_path: str) -> Document:
    """
    Opens a Word document (.docx) from the specified file path and returns a Document object.

    :param file_path: The path to the Word document file to be opened.
    :type file_path: str
    :return: A Document object representing the Word document.
    :rtype: docx.Document

    :raises FileNotFoundError: If the file at the specified path does not exist.
    :raises ValueError: If the file path does not point to a .docx file.
    """
    # Check if the file path is valid and points to a .docx file
    if not file_path.endswith('.docx'):
        raise ValueError("The file path must point to a .docx file.")

    # Attempt to open the document
    try:
        doc = Document(file_path)
        print(f"Successfully opened the document: {file_path}")
        return doc
    except FileNotFoundError:
        print(f"Error: The file at {file_path} was not found.")
        raise
    except Exception as e:
        print(f"An error occurred while opening the document: {e}")
        raise

def extract_hyperlinks(doc: Document) -> dict:
    """
    Extracts all hyperlinks from a Word document and returns them in a dictionary with 
    the hyperlink text as keys and the URLs as values.

    :param file_path: The document that was inputted in from the user.
    :type file_path: Document
    :return: A dictionary containing hyperlink texts and their corresponding URLs.
    :rtype: dict
    """

    hyperlinks = {}  # Dictionary to store hyperlinks and their texts

    # Loop through paragraphs and runs to find hyperlinks
    for paragraph in doc.paragraphs:
        for run in paragraph.runs:
            # Check if the run contains a hyperlink
            r_element = run._r  # Access the run element
            if r_element is not None:
                # Look for the hyperlink relationship (r:hyperlink)
                hyperlink_tag = r_element.find(qn('w:hyperlink'))
                if hyperlink_tag is not None:
                    # Extract the hyperlink ID from the attribute
                    hyperlink_id = hyperlink_tag.get('{http://schemas.openxmlformats.org/officeDocument/2006/relationships}id')
                    
                    if hyperlink_id is not None:
                        # Get the actual URL from the document's relationships
                        url = doc.part.rels[hyperlink_id].target
                        # Add the hyperlink text and URL to the dictionary
                        hyperlinks[run.text] = url
    
    print("Hyperlinks extracted successfully.")
    return hyperlinks

def change_url(urls: list) -> list: ...

def update_urls_in_doc(doc, modified_urls: list) -> None: ...

def save_doc(doc, path: str) -> None: ...

if __name__ == '__main__':
    main()
