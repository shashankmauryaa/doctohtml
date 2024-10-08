import webbrowser
from docx import Document
import os
import re

def get_font_style(run):
    font_style = ''
    if run.bold:
        font_style += 'font-weight: bold; '
    if run.italic:
        font_style += 'font-style: italic; '
    if run.underline:
        font_style += 'text-decoration: underline; '
    if run.font.name:
        font_style += f'font-family: {run.font.name}; '
    if run.font.size:
        font_style += f'font-size: {run.font.size.pt}pt; '
    return font_style.strip()

def convert_docx_to_html(docx_file, html_file, output_folder):
    doc = Document(docx_file)
    image_filenames = []
    used_image_filenames = set()  # Keep track of used image filenames

    with open(html_file, 'w', encoding='utf-8') as f:
        f.write('<html>\n<head>\n<title>Converted Document</title>\n')
        f.write('<style>\n.content {\n padding-left: 250px; \n padding-right: 250px;\n}\n</style>\n')
        f.write('</head>\n<body>\n')
        f.write('<div class="content">\n')
        
        bullet_points = []

        for paragraph in doc.paragraphs:
            if paragraph.style.name.startswith('List') and paragraph.text.strip():
                bullet_points.append(paragraph.text.strip())
            else:
                pattern = r'http[s]?://\S+'  # Regular expression pattern to match hyperlinks
                after_split = re.split(f'({pattern})', paragraph.text)  # Split the text at each hyperlink
                f.write('<p>')
                for item in after_split:
                    if item.strip():  # Skip empty strings
                        if re.match(pattern, item):  # Check if the item matches the hyperlink pattern
                            f.write(f'<a href="{item.strip()}">{item.strip()}</a>')  # Write as a hyperlink
                        else:
                            # Write the text with its font style
                            for run in paragraph.runs:
                                if run.text in item:  # Check if the run's text is part of the item
                                    font_style = get_font_style(run)
                                    if font_style:
                                        f.write(f'<span style="{font_style}">{run.text}</span>')
                                    else:
                                        f.write(run.text)
                f.write('</p>')

            # Handle images in runs
            for run in paragraph.runs:
                if run.text.strip() == "":
                    # Check for drawing elements for images
                    for drawing in run._element.findall('.//{http://schemas.openxmlformats.org/drawingml/2006/main}blip'):
                        rId = drawing.get("{http://schemas.openxmlformats.org/officeDocument/2006/relationships}embed")
                        if rId is not None:
                            image_filename = f"image{rId}.png"
                            image_path = os.path.join(output_folder, image_filename)
                            if image_filename not in used_image_filenames:
                                image_data = doc.part.related_parts[rId].blob
                                with open(image_path, "wb") as img_file:
                                    img_file.write(image_data)
                                image_filenames.append(image_filename)
                                used_image_filenames.add(image_filename)
                                f.write(f'<img src="{image_filename}" alt="Image" style="width: 100%; height: auto;" /><br />\n')

        if bullet_points:
            f.write('<ul>\n')
            for bullet_point in bullet_points:
                f.write(f'<li>{bullet_point}</li>\n')
            f.write('</ul>\n')

        f.write('</div>\n')
        f.write('</body>\n</html>')

docx_file = r"C:\Users\Shashank_Maurya\OneDrive - Dell Technologies\Documents\doc-to-html\iOSDocumentForIntuneRelease.docx"
html_file = 'output.html'
output_folder = r"C:\Users\Shashank_Maurya\OneDrive - Dell Technologies\Documents\doc-to-html"

convert_docx_to_html(docx_file, html_file, output_folder)

webbrowser.open(html_file)
