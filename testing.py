from docx import Document
from docx.opc.constants import RELATIONSHIP_TYPE
from docx.oxml.ns import qn

def extract_hyperlinks(doc_path):
    # Load the document
    doc = Document(doc_path)
    
    # Dictionary to hold relationship IDs and hyperlinks
    rels = {}
    
    # Retrieve all the relationship IDs and URLs (external hyperlinks)
    for rel in doc.part.rels.values():
        if rel.reltype == RELATIONSHIP_TYPE.HYPERLINK:
            rels[rel.rId] = rel._target
    
    # Initialize a list to store extracted text and hyperlinks
    extracted_data = []
    
    # Iterate over paragraphs in the document
    for para in doc.paragraphs:
        
        for run in para.runs:
            # Check if the run is part of a hyperlink
            hyperlink = run._element.getparent().find(qn('w:hyperlink'))
            if hyperlink is not None:
                rId = hyperlink.get(qn('r:id'))
                if rId in rels:
                    # Extract hyperlink URL
                    full_hyperlink = rels[rId]
                    # Append the hyperlink text and URL
                    extracted_data.append((run.text, full_hyperlink))
            else:
                # Add regular text without a hyperlink
                extracted_data.append((run.text, None))
    
    return extracted_data

# Example usage
doc_path = 'D:\SMIT\SEMESTER 7\DELL intern\iOSDocumentForIntuneRelease.docx'
hyperlink_data = extract_hyperlinks(doc_path)

# Print extracted text and hyperlinks
for text, hyperlink in hyperlink_data:
    if hyperlink:
        print(f"Text: {text} -> Hyperlink: {hyperlink}")
    else:
        print(f"Text: {text}")
