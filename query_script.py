import fitz
import docx
import os

pdf_path = r'D:\JB Pharma internal\JB Pharma\Cardiolek\Module 3\32-Body-Data\32S-Drug Substance\3.2.S. Cardiolek-320 Injection\3.2.S.2.3.pdf'
docx_path = r'D:\JB Pharma internal\JB Pharma\Cardiolek\Output\Cardiolek_QIS.docx'

print("--- PDF Extraction ---")
if os.path.exists(pdf_path):
    doc = fitz.open(pdf_path)
    page = doc[0]
    text = page.get_text()
    print(text[:1200])
else:
    print(f"PDF not found: {pdf_path}")

print("\n--- DOCX Content ---")
if os.path.exists(docx_path):
    doc = docx.Document(docx_path)
    found_s23 = False
    count_s23 = 0
    
    for i, para in enumerate(doc.paragraphs):
        if '3.2.S.2.3' in para.text:
            found_s23 = True
            print(f"Found heading: {para.text}")
            # Get next 12 non-empty paragraphs
            collected = 0
            for j in range(i + 1, len(doc.paragraphs)):
                if collected >= 12: break
                p_text = doc.paragraphs[j].text.strip()
                if p_text:
                    print(f"P{collected+1}: {p_text}")
                    collected += 1
            break
            
    print("\n--- DOCX Tables Search ---")
    # Finding tables between 2.3.S.4.1 and next section heading
    # In python-docx, tables and paragraphs are separate but we can iterate through the document body elements
    from docx.oxml.table import CT_Tbl
    from docx.oxml.text.paragraph import CT_P
    from docx.table import Table
    from docx.text.paragraph import Paragraph

    def iter_block_items(parent):
        from docx.document import Document
        if isinstance(parent, Document):
            parent_elm = parent.element.body
        else:
            parent_elm = parent._element
        for child in parent_elm.iterchildren():
            if isinstance(child, CT_P):
                yield Paragraph(child, parent)
            elif isinstance(child, CT_Tbl):
                yield Table(child, parent)

    found_s41 = False
    tables_found = []
    for item in iter_block_items(doc):
        if isinstance(item, Paragraph):
            if '2.3.S.4.1' in item.text:
                found_s41 = True
                continue
            # Logic for next section heading: assume next section starts with X.X.S or digit
            if found_s41 and item.style.name.startswith('Heading') and item.text.strip():
                 # Filter out if it's the heading we just found
                 if '2.3.S.4.1' not in item.text:
                    break
        elif isinstance(item, Table) and found_s41:
            tables_found.append(item)
            
    print(f"Tables found between 2.3.S.4.1 and next section: {len(tables_found)}")
    for idx, tbl in enumerate(tables_found):
        rows = len(tbl.rows)
        cols = len(tbl.columns)
        first_row = [cell.text.strip() for cell in tbl.rows[0].cells]
        print(f"Table {idx+1}: {rows}x{cols}, Header: {first_row}")
else:
    print(f"DOCX not found: {docx_path}")

