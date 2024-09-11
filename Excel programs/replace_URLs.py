import re
import os
import logging
import win32com.client as win32  # Import the win32 client for Word automation

def main() -> None:
    # Configure logging
    logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
    
    # Define the path to the Word document
    doc_path = 'your_document_path.docx'  # Replace with your actual path
    
    # Load the Word document
    doc = get_doc(doc_path)
    
    # Get the URLs from the document
    urls = get_url(doc)
    
    # Change the URLs as per your requirement
    modified_urls = change_url(urls)
    
    # Update the document with the modified URLs
    update_urls_in_doc(doc, modified_urls)
    
    # Save the modified document
    save_doc(doc, doc_path)
    
    logging.info("Processing complete. Document saved with updated URLs.")

def get_doc(path: str):
    """
    Opens the Word document using win32com.client.
    """
    # Initialize Word application
    word = win32.Dispatch('Word.Application')
    word.Visible = False  # Keep Word invisible to the user

    # Open the document
    try:
        doc = word.Documents.Open(path)
        logging.info(f'Document {path} loaded successfully.')
        return doc
    except Exception as e:
        logging.error(f'Error loading document: {e}')
        raise

def get_url(doc) -> list:
    """
    Extracts all hyperlinks from the Word document.
    """
    urls = []
    for hyperlink in doc.Hyperlinks:
        urls.append(hyperlink.Address)  # Extract the hyperlink URL
    logging.info(f'Found {len(urls)} hyperlinks in the document.')
    return urls

def change_url(urls: list) -> list:
    """
    Modify the URLs using regex or other logic to cut off and replace parts of the URLs.
    """
    modified_urls = []
    for url in urls:
        # Example: Replace part of the URL (customize this with your own regex)
        modified_url = re.sub(r'old-part', 'new-part', url)
        modified_urls.append(modified_url)
        logging.info(f'Changed URL from {url} to {modified_url}')
    return modified_urls

def update_urls_in_doc(doc, modified_urls: list) -> None:
    """
    Update the URLs in the Word document with modified URLs.
    """
    for i, hyperlink in enumerate(doc.Hyperlinks):
        hyperlink.Address = modified_urls[i]
    logging.info('All URLs have been updated in the document.')

def save_doc(doc, path: str) -> None:
    """
    Save the modified Word document.
    """
    try:
        doc.SaveAs(path)
        doc.Close()
        logging.info(f'Document saved successfully at {path}.')
    except Exception as e:
        logging.error(f'Error saving document: {e}')
        raise

if __name__ == '__main__':
    main()
