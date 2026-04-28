import docx
from docx.table import Table
from docx.text.paragraph import Paragraph

def inspect_docx(path):
    doc = docx.Document(path)
    body = doc.element.body
    
    elements = []
    for child in body.xpath('./w:p | ./w:tbl'):
        if child.tag.endswith('p'):
            elements.append(Paragraph(child, doc))
        elif child.tag.endswith('tbl'):
            elements.append(Table(child, doc))
            
    found_start = False
    
    for item in elements:
        text = ''
        if isinstance(item, Paragraph):
            text = item.text
        elif isinstance(item, Table):
            text = "".join(cell.text for row in item.rows for cell in row.cells)
            
        if '2.3.S.4.1' in text and not found_start:
            found_start = True
            
        if found_start:
            if '2.3.S.6' in text and not ('2.3.S.4.1' in text):
                 break
            
            if isinstance(item, Paragraph):
                print(f"P: {item.text[:100]}")
            elif isinstance(item, Table):
                rows = len(item.rows)
                cols = len(item.columns)
                first_row = [cell.text.strip() for cell in item.rows[0].cells if cell.text.strip()]
                print(f"T: {rows}x{cols}, First row: {first_row}")

inspect_docx(r'D:\JB Pharma internal\JB Pharma\Cardiolek\Output\Cardiolek_QIS.docx')
