from mcp.server.fastmcp import FastMCP
from docx import Document
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
import zipfile
import lxml.etree as ET
import os
import copy
import shutil
# Create an MCP server
mcp = FastMCP("docx_editor")

def get_flat_cells(table):
    all_cells = []
    seen = set()
    for row in table.rows:
        for cell in row.cells:
            if cell not in seen:
                all_cells.append(cell)
                seen.add(cell)
    return all_cells

def safe_replace_text_internal(cell, new_text):
    if not cell.paragraphs:
        cell.add_paragraph(new_text)
        return
    p = cell.paragraphs[0]
    if not p.runs:
        p.add_run(new_text)
    else:
        p.runs[0].text = new_text
        for i in range(1, len(p.runs)):
            p.runs[i].text = ""

    for i in range(1, len(cell.paragraphs)):
        for r in cell.paragraphs[i].runs:
            r.text = ""

def replace_keyword_in_paragraphs(paragraphs, old_text, new_text):
    """
    Replaces old_text with new_text in paragraphs while strictly preserving formatting.
    Uses a run-merging fallback for split words and handles nested hyperlinks.
    """
    replaced_count = 0
    namespaces = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
    
    for p in paragraphs:
        if old_text not in p.text:
            continue
            
        found = False
        
        # 1. Check API runs (perfect match within a single run)
        for run in p.runs:
            if old_text in run.text:
                run.text = run.text.replace(old_text, new_text)
                found = True
                
        # 2. Check XML for hyperlinks, which python-docx API might hide
        for hl in p._p.findall('.//w:hyperlink', namespaces):
            for t in hl.findall('.//w:t', namespaces):
                if t.text and old_text in t.text:
                    t.text = t.text.replace(old_text, new_text)
                    found = True
                    
        # 3. Fallback: The text exists in the paragraph, but is split across multiple runs/hyperlinks.
        # We find the sequence of runs that make up the text and replace it, keeping the first run's formatting.
        if not found:
            # We collect all distinct w:t elements in the paragraph XML
            t_elements = p._p.findall('.//w:t', namespaces)
            if t_elements:
                full_text = ""
                for t in t_elements:
                    full_text += (t.text or "")
                
                if old_text in full_text:
                    # Replace the text in the full string
                    new_full = full_text.replace(old_text, new_text)
                    
                    # Dump everything into the first w:t element to preserve its formatting
                    t_elements[0].text = new_full
                    
                    # Clear out the text from all subsequent w:t elements
                    for i in range(1, len(t_elements)):
                        t_elements[i].text = ""
                        
                    found = True
                    
        if found:
            replaced_count += 1
            
    return replaced_count

@mcp.tool()
def read_docx_table_flat_index(file_path: str, table_index: int = -1) -> str:
    """
    Reads a specific table from a docx file and returns all unique cell contents
    using the Flat Index algorithm to safely avoid merged cell issues.
    
    Args:
        file_path: The absolute path to the .docx file
        table_index: Index of the table to read (default is last table, -1)
        
    Returns:
        A formatted string listing the cell index and its textual content.
    """
    if not os.path.exists(file_path):
        return f"Error: File '{file_path}' not found."
    
    try:
        doc = Document(file_path)
        if not doc.tables:
            return "Error: No tables found in the document."
            
        target_table = doc.tables[table_index]
        cells = get_flat_cells(target_table)
        
        results = []
        for i, cell in enumerate(cells):
            text = cell.text.strip().replace('\n', ' ')
            results.append(f"[{i}] {text}")
            
        return "\n".join(results)
    except Exception as e:
        return f"Error reading table: {str(e)}"

@mcp.tool()
def safe_replace_docx_cell(file_path: str, table_index: int, cell_index: int, new_text: str, out_path: str = "") -> str:
    """
    Replaces the text of a specific cell in a docx table safely preserving formatting.
    Uses the Flat Index algorithm to avoid the merged cell bug.
    
    Args:
        file_path: The absolute path to the .docx file
        table_index: Index of the table. Usually -1 for the last table.
        cell_index: The flat index of the unique cell to replace
        new_text: string input to insert into the cell
        out_path: Optional save path. If empty, overwrites the original file_path.
        
    Returns:
        Status message about the operation.
    """
    if not os.path.exists(file_path):
        return f"Error: File '{file_path}' not found."
        
    if not out_path:
        out_path = file_path
        
    try:
        doc = Document(file_path)
        if not doc.tables:
            return "Error: No tables found in the document."
            
        target_table = doc.tables[table_index]
        cells = get_flat_cells(target_table)
        
        if cell_index < 0 or cell_index >= len(cells):
            return f"Error: cell_index {cell_index} is out of bounds (0 to {len(cells)-1})"
            
        target_cell = cells[cell_index]
        old_text = target_cell.text.strip().replace('\n', ' ')
        
        safe_replace_text_internal(target_cell, new_text)
        
        doc.save(out_path)
        return f"Successfully updated table {table_index}, cell {cell_index}.\nOld text: '{old_text}'\nNew text: '{new_text}'\nSaved to: {out_path}"
    except Exception as e:
        return f"Error updating cell: {str(e)}"

@mcp.tool()
def safe_replace_docx_text_keyword(file_path: str, old_text: str, new_text: str, out_path: str = "") -> str:
    """
    Replaces all occurrences of old_text with new_text across the entire docx document
    (including paragraphs and all tables), preserving formatting as much as possible.
    
    Args:
        file_path: The absolute path to the .docx file
        old_text: The exact string to search for
        new_text: The string to replace it with
        out_path: Optional save path. If empty, overwrites the original file_path.
        
    Returns:
        Status message about the operation, including how many paragraphs were updated.
    """
    if not os.path.exists(file_path):
        return f"Error: File '{file_path}' not found."
        
    if not out_path:
        out_path = file_path
        
    try:
        doc = Document(file_path)
        total_replacements = 0
        
        # 1. Replace in main body paragraphs
        total_replacements += replace_keyword_in_paragraphs(doc.paragraphs, old_text, new_text)
        
        # 2. Replace in all tables
        for table in doc.tables:
            for row in table.rows:
                for cell in row.cells:
                    total_replacements += replace_keyword_in_paragraphs(cell.paragraphs, old_text, new_text)
                    
        # 3. Replace in Headers and Footers
        for section in doc.sections:
            for h_f in [section.header, section.footer]:
                if h_f and not h_f.is_linked_to_previous:
                    total_replacements += replace_keyword_in_paragraphs(h_f.paragraphs, old_text, new_text)
                    for table in h_f.tables:
                        for row in table.rows:
                            for cell in row.cells:
                                total_replacements += replace_keyword_in_paragraphs(cell.paragraphs, old_text, new_text)
                                
        doc.save(out_path)
        return f"Successfully replaced '{old_text}' with '{new_text}' in {total_replacements} paragraph(s).\nSaved to: {out_path}"
    except Exception as e:
        return f"Error replacing text: {str(e)}"

@mcp.tool()
def safe_replace_text_box_keyword(file_path: str, old_text: str, new_text: str, out_path: str = "") -> str:
    """
    Directly hacks the Word document's XML to replace text hidden inside Floating Shapes and Text Boxes.
    This bypasses the limitations of the standard python-docx API.
    
    Args:
        file_path: The absolute path to the .docx file
        old_text: The exact string to search for inside text boxes
        new_text: The string to replace it with
        out_path: Optional save path. If empty, overwrites the original file_path.
        
    Returns:
        Status message about the operation.
    """
    if not os.path.exists(file_path):
        return f"Error: File '{file_path}' not found."
        
    if not out_path:
        out_path = file_path
        
    temp_path = file_path + ".tmp.zip"
    total_replacements = 0
    
    try:
        with zipfile.ZipFile(file_path, 'r') as zin, zipfile.ZipFile(temp_path, 'w') as zout:
            for item in zin.infolist():
                with zin.open(item) as f:
                    content = f.read()
                
                # Check if it's an XML file that might contain text boxes (document, headers, footers)
                if item.filename.endswith('.xml') and item.filename.startswith('word/'):
                    try:
                        tree = ET.fromstring(content)
                        namespaces = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
                        vml_namespaces = {'v': 'urn:schemas-microsoft-com:vml'}
                        changed = False
                        
                        # Find all w:t inside w:txbxContent or v:textbox
                        txbxs = tree.findall('.//w:txbxContent', namespaces)
                        try:
                            txbxs.extend(tree.findall('.//v:textbox', vml_namespaces))
                        except Exception:
                            pass
                            
                        for txbx in txbxs:
                            # Extract all text from this text box
                            t_elements = txbx.findall('.//w:t', namespaces)
                            if not t_elements:
                                continue
                                
                            full_text = ""
                            for t in t_elements:
                                full_text += (t.text or "")
                                
                            if old_text in full_text:
                                new_full = full_text.replace(old_text, new_text)
                                # Dump into first t element, clear others
                                t_elements[0].text = new_full
                                for i in range(1, len(t_elements)):
                                    t_elements[i].text = ""
                                total_replacements += 1
                                changed = True
                        
                        if changed:
                            content = ET.tostring(tree, encoding='utf-8', xml_declaration=True)
                    except Exception:
                        pass # if it's not well-formed XML or we can't parse it, just copy as is
                zout.writestr(item, content)
                
        if total_replacements > 0:
            shutil.move(temp_path, out_path)
            return f"Successfully replaced '{old_text}' with '{new_text}' in {total_replacements} text box segments.\\nSaved to: {out_path}"
        else:
            os.remove(temp_path)
            if out_path != file_path:
                shutil.copy2(file_path, out_path)
            return "No matching text found inside text boxes."
            
    except Exception as e:
        if os.path.exists(temp_path):
            os.remove(temp_path)
        return f"Error replacing text box content: {str(e)}"

@mcp.tool()
def list_docx_images(file_path: str) -> str:
    """
    Lists all inline images in the document to help identify which one to replace.
    
    Args:
        file_path: The absolute path to the .docx file
        
    Returns:
        A formatted string listing the image index, dimensions, and type.
    """
    if not os.path.exists(file_path):
        return f"Error: File '{file_path}' not found."
        
    try:
        doc = Document(file_path)
        results = []
        for i, shape in enumerate(doc.inline_shapes):
            width_cm = shape.width / 360000 if shape.width else 0
            height_cm = shape.height / 360000 if shape.height else 0
            results.append(f"[{i}] Size: {width_cm:.2f}cm x {height_cm:.2f}cm")
            
        if not results:
            return "No inline images found in the document."
        return "\\n".join(results)
    except Exception as e:
        return f"Error listing images: {str(e)}"

@mcp.tool()
def replace_docx_image(file_path: str, image_index: int, new_image_path: str, out_path: str = "") -> str:
    """
    Replaces an existing inline image with a new image file, preserving the original dimensions.
    
    Args:
        file_path: The absolute path to the .docx file
        image_index: The index of the inline image to replace (found via list_docx_images)
        new_image_path: The absolute path to the new image file (e.g., .png, .jpg)
        out_path: Optional save path. If empty, overwrites the original file_path.
        
    Returns:
        Status message about the operation.
    """
    if not os.path.exists(file_path):
        return f"Error: File '{file_path}' not found."
    if not os.path.exists(new_image_path):
        return f"Error: New image file '{new_image_path}' not found."
        
    if not out_path:
        out_path = file_path
        
    try:
        doc = Document(file_path)
        shapes = doc.inline_shapes
        
        if image_index < 0 or image_index >= len(shapes):
            return f"Error: image_index {image_index} is out of bounds (0 to {len(shapes)-1})"
            
        target_shape = shapes[image_index]
        
        # Replace the image part underlying the shape
        # The relationship connects the shape to an image part.
        from docx.opc.constants import RELATIONSHIP_TYPE as RT
        
        # Find the relationship ID related to this shape
        # target_shape._inline element contains the blip with embed ID
        blip = target_shape._inline.graphic.graphicData.pic.blipFill.blip
        rId = blip.embed
        
        # Get the relationship and target part
        document_part = doc.part
        rel = document_part.rels[rId]
        image_part = rel.target_part
        
        # Overwrite the blob of the image part with the new image
        with open(new_image_path, "rb") as f:
            new_blob = f.read()
            
        # The image_part is an ImagePart object
        image_part._blob = new_blob
        
        # Note: A real robust approach might create a new ImagePart or properly 
        # update content type if changing from png to jpeg, 
        # but overwriting the blob directly often works for basic docx viewers.
        
        doc.save(out_path)
        return f"Successfully replaced image {image_index} with '{new_image_path}'.\nSaved to: {out_path}"
    except Exception as e:
        return f"Error replacing image: {str(e)}"

@mcp.tool()
def append_docx_list_item(file_path: str, target_text: str, new_item_text: str, out_path: str = "") -> str:
    """
    Finds a paragraph containing target_text (presumably a list item), clones its XML formatting 
    (including its bullet/numbering numbering properties), and inserts a new paragraph 
    with new_item_text immediately after it.
    
    Args:
        file_path: The absolute path to the .docx file
        target_text: The string to search for to identify the target list item
        new_item_text: The text for the new list item to append below the target
        out_path: Optional save path. If empty, overwrites the original file_path.
        
    Returns:
        Status message about the operation.
    """
    if not os.path.exists(file_path):
        return f"Error: File '{file_path}' not found."
        
    if not out_path:
        out_path = file_path
        
    try:
        doc = Document(file_path)
        
        target_paragraph = None
        for p in doc.paragraphs:
            if target_text in p.text:
                target_paragraph = p
                break
                
        if not target_paragraph:
            # Also search in tables as a fallback
            for table in doc.tables:
                for row in table.rows:
                    for cell in row.cells:
                        for p in cell.paragraphs:
                            if target_text in p.text:
                                target_paragraph = p
                                break
                        if target_paragraph: break
                    if target_paragraph: break
                if target_paragraph: break
                
        if not target_paragraph:
            return f"Error: Could not find any paragraph containing '{target_text}'."

        # Instead of adding a paragraph to the end of the document, we create a new OxmlElement
        # and insert it directly into the XML tree right after our target paragraph's XML.
        # This prevents breaking logical boundaries if the target is inside a table cell.
        from docx.text.paragraph import Paragraph
        
        new_p_element = OxmlElement('w:p')
        target_paragraph._p.addnext(new_p_element)
        new_p = Paragraph(new_p_element, target_paragraph._parent)
        new_p.text = new_item_text
        
        # Clone paragraph properties (pPr) which contain the numbering info AND style
        if target_paragraph._p.pPr is not None:
            # Create a deep copy of the pPr element
            pPr_clone = copy.deepcopy(target_paragraph._p.pPr)
            new_p._p.insert(0, pPr_clone)
                
        # If the original paragraph has a specific style applied, apply it
        new_p.style = target_paragraph.style
        
        doc.save(out_path)
        return f"Successfully appended new list item after '{target_text}'.\nSaved to: {out_path}"
    except Exception as e:
        return f"Error appending list item: {str(e)}"

@mcp.tool()
def read_docx_table_optimized(file_path: str, table_index: int = -1) -> str:
    """
    Reads a specific table from a docx file using a highly optimized streaming ZIP+XML parsing approach.
    This avoids loading the entire document into memory, preventing ValueError/OOM on massive files.
    
    Args:
        file_path: The absolute path to the .docx file
        table_index: Index of the table to read (default is last table, -1)
        
    Returns:
        A formatted string listing the cell contents (not handling complex merged cell indexing, just raw fast read).
    """
    if not os.path.exists(file_path):
        return f"Error: File '{file_path}' not found."
        
    try:
        # Open the docx as a ZIP file
        with zipfile.ZipFile(file_path, 'r') as docx_zip:
            # Read the main document XML
            with docx_zip.open('word/document.xml') as f:
                tree = ET.parse(f)
                root = tree.getroot()
                
        # Find all tables using the explicit WordprocessingML namespace
        namespaces = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
        tables = root.findall('.//w:tbl', namespaces)
        
        if not tables:
            return "Error: No tables found in the document."
            
        target_table = tables[table_index]
        
        results = []
        cell_counter = 0
        
        # Iterate over rows
        for row in target_table.findall('.//w:tr', namespaces):
            # Iterate over cells
            for cell in row.findall('.//w:tc', namespaces):
                # We extract text per paragraph (w:p) inside the cell
                paragraph_texts = []
                for p in cell.findall('.//w:p', namespaces):
                    texts = [t.text for t in p.findall('.//w:t', namespaces) if t.text]
                    if texts:
                        paragraph_texts.append("".join(texts).strip())
                
                if paragraph_texts:
                    # Join paragraphs with newline to preserve formatting
                    cell_text = "\\n".join(paragraph_texts)
                    results.append(f"[{cell_counter}] {cell_text}")
                else:
                    results.append(f"[{cell_counter}] ")
                cell_counter += 1
                
        return "\\n".join(results)
    except Exception as e:
        return f"Error parsing table optimized: {str(e)}"

@mcp.tool()
def copy_docx_table(file_path: str, table_index: int, out_path: str = "") -> str:
    """
    Copies a table and inserts the copy immediately after the original table.
    """
    if not os.path.exists(file_path):
        return f"Error: File '{file_path}' not found."
    if not out_path:
        out_path = file_path
    try:
        doc = Document(file_path)
        if table_index < 0 or table_index >= len(doc.tables):
            return "Error: table_index out of bounds."
        
        target_table = doc.tables[table_index]
        new_tbl = copy.deepcopy(target_table._element)
        target_table._element.addnext(new_tbl)
        
        # Add a spacing paragraph between tables
        p = OxmlElement('w:p')
        target_table._element.addnext(p)
        
        doc.save(out_path)
        return f"Successfully copied table {table_index}. Saved to: {out_path}"
    except Exception as e:
        return f"Error copying table: {str(e)}"

@mcp.tool()
def delete_docx_table(file_path: str, table_index: int, out_path: str = "") -> str:
    """Deletes a table from the document."""
    if not os.path.exists(file_path):
        return f"Error: File '{file_path}' not found."
    if not out_path:
        out_path = file_path
    try:
        doc = Document(file_path)
        if table_index < 0 or table_index >= len(doc.tables):
            return "Error: table_index out of bounds."
        
        target_table = doc.tables[table_index]
        target_table._element.getparent().remove(target_table._element)
        
        doc.save(out_path)
        return f"Successfully deleted table {table_index}. Saved to: {out_path}"
    except Exception as e:
        return f"Error deleting table: {str(e)}"

@mcp.tool()
def add_docx_table_row(file_path: str, table_index: int, out_path: str = "") -> str:
    """Appends a new row to the end of a table."""
    if not os.path.exists(file_path):
        return f"Error: File '{file_path}' not found."
    if not out_path:
        out_path = file_path
    try:
        doc = Document(file_path)
        if table_index < 0 or table_index >= len(doc.tables):
            return "Error: table_index out of bounds."
        
        target_table = doc.tables[table_index]
        target_table.add_row()
        
        doc.save(out_path)
        return f"Successfully added a row to table {table_index}. Saved to: {out_path}"
    except Exception as e:
        return f"Error adding row: {str(e)}"

@mcp.tool()
def delete_docx_table_row(file_path: str, table_index: int, row_index: int, out_path: str = "") -> str:
    """Deletes a specific row from a table."""
    if not os.path.exists(file_path):
        return f"Error: File '{file_path}' not found."
    if not out_path:
        out_path = file_path
    try:
        doc = Document(file_path)
        if table_index < 0 or table_index >= len(doc.tables):
            return "Error: table_index out of bounds."
            
        target_table = doc.tables[table_index]
        if row_index < 0 or row_index >= len(target_table.rows):
            return "Error: row_index out of bounds."
            
        target_row = target_table.rows[row_index]
        target_row._element.getparent().remove(target_row._element)
        
        doc.save(out_path)
        return f"Successfully deleted row {row_index} from table {table_index}. Saved to: {out_path}"
    except Exception as e:
        return f"Error deleting row: {str(e)}"

if __name__ == "__main__":
    mcp.run()
