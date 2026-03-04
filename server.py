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

def get_flat_cells_with_position(table):
    """
    Returns unique cells with their (row_index, col_index) position.

    Returns:
        tuple: (all_cells, cell_to_position)
            - all_cells: List of unique cell objects
            - cell_to_position: Dict mapping cell_id to position info
    """
    all_cells = []
    seen = set()
    cell_to_position = {}

    for row_idx, row in enumerate(table.rows):
        for col_idx, cell in enumerate(row.cells):
            cell_id = id(cell)
            if cell_id not in seen:
                all_cells.append(cell)
                seen.add(cell_id)
                cell_to_position[cell_id] = {
                    "row_index": row_idx,
                    "col_index": col_idx
                }

    return all_cells, cell_to_position

def get_table_preview(table, rows=3, max_cell_len=50):
    """
    Generate a text preview of a table.

    Args:
        table: Table object
        rows: Number of rows to preview
        max_cell_len: Maximum characters per cell

    Returns:
        str: Formatted table preview
    """
    preview_lines = []
    row_count = min(rows, len(table.rows))

    for row_idx in range(row_count):
        row = table.rows[row_idx]
        cells_text = []
        for cell in row.cells:
            text = cell.text.strip().replace('\n', ' ')
            if len(text) > max_cell_len:
                text = text[:max_cell_len-3] + "..."
            cells_text.append(text)
        preview_lines.append(" | ".join(cells_text))

    if len(table.rows) > rows:
        preview_lines.append("...")

    return "\n".join(preview_lines)

def format_cell_text(text, max_len=50):
    """Format cell text for display."""
    text = text.strip().replace('\n', ' ')
    if len(text) > max_len:
        return text[:max_len-3] + "..."
    return text

def identify_table_location(doc, table_index):
    """
    Identify if a table is in header, footer, or body.

    Returns:
        str: "Header", "Footer", or "Body"
    """
    header_count = 0
    footer_count = 0

    for section in doc.sections:
        header_count += len(section.header.tables)
        footer_count += len(section.footer.tables)

    if table_index < header_count:
        return "Header"
    elif table_index >= len(doc.tables) - footer_count:
        return "Footer"
    else:
        return "Body"

def get_actual_table_index(doc, tbl_element):
    """
    Maps XML table element to doc.tables index.

    Args:
        doc: Document object
        tbl_element: CT_Tbl (XML table element)

    Returns:
        int: Index in doc.tables, or -1 if not found
    """
    for idx, table in enumerate(doc.tables):
        if table._element is tbl_element:
            return idx
    return -1

def validate_table_index(doc, table_index):
    """
    Validate table index and provide helpful error message.

    Raises:
        ValueError: If index is out of range
    """
    total_tables = len(doc.tables)

    # 음수 인덱스 처리
    if table_index < 0:
        actual_index = total_tables + table_index
        if actual_index < 0:
            raise ValueError(
                f"Table index {table_index} out of range. "
                f"Document has {total_tables} tables. "
                f"Use list_all_tables() to see available tables."
            )
        return actual_index

    # 양수 인덱스 처리
    if table_index >= total_tables:
        raise ValueError(
            f"Table index {table_index} out of range. "
            f"Document has {total_tables} tables (indices 0-{total_tables-1}). "
            f"Use list_all_tables() to see available tables."
        )

    return table_index

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

        # 테이블 인덱스 검증
        try:
            table_index = validate_table_index(doc, table_index)
        except ValueError as e:
            return f"Error: {str(e)}"

        target_table = doc.tables[table_index]
        # 개선된 함수 사용 - 위치 정보 포함
        cells, positions = get_flat_cells_with_position(target_table)

        results = []
        for i, cell in enumerate(cells):
            pos = positions[id(cell)]
            text = cell.text.strip().replace('\n', ' ')[:100]  # 최대 100자
            results.append(f"Cell {i} [Row {pos['row_index']}, Col {pos['col_index']}]: {text}")

        return "\n".join(results)
    except Exception as e:
        return f"Error reading table: {str(e)}"

@mcp.tool()
def safe_replace_docx_cell(file_path: str, table_index: int, cell_index: int, new_text: str, out_path: str = "", expected_old_text: str = "") -> str:
    """
    Replaces the text of a specific cell in a docx table safely preserving formatting.
    Uses the Flat Index algorithm to avoid the merged cell bug.
    
    Args:
        file_path: The absolute path to the .docx file
        table_index: Index of the table. Usually -1 for the last table.
        cell_index: The flat index of the unique cell to replace
        new_text: string input to insert into the cell
        out_path: Optional save path. If empty, overwrites the original file_path.
        expected_old_text: Optional. If provided, the replacement will ONLY proceed if this text is found within the target cell's original text. This acts as a safety guard to prevent overwriting the wrong cell.
        
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
        
        # Verify safety guard if provided
        if expected_old_text and expected_old_text not in target_cell.text:
            return f"Safety Guard Error: Expected text '{expected_old_text}' not found in target cell. Actual text was '{old_text}'. Aborting operation."
        
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
def safe_replace_docx_table_keyword(file_path: str, table_index: int, old_text: str, new_text: str, out_path: str = "") -> str:
    """
    Replaces all occurrences of old_text with new_text ONLY within a specific table,
    preserving formatting as much as possible.
    
    Args:
        file_path: The absolute path to the .docx file
        table_index: Index of the target table
        old_text: The exact string to search for
        new_text: The string to replace it with
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
        if table_index < 0 or table_index >= len(doc.tables):
            return f"Error: table_index {table_index} is out of bounds."
            
        target_table = doc.tables[table_index]
        total_replacements = 0
        
        for row in target_table.rows:
            for cell in row.cells:
                total_replacements += replace_keyword_in_paragraphs(cell.paragraphs, old_text, new_text)
                
        doc.save(out_path)
        return f"Successfully replaced '{old_text}' with '{new_text}' in {total_replacements} paragraph(s) of Table {table_index}.\nSaved to: {out_path}"
    except Exception as e:
        return f"Error replacing text in table: {str(e)}"

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
def insert_docx_table_row(file_path: str, table_index: int, row_index: int, row_data: list, out_path: str = "") -> str:
    """
    Inserts a new row with provided text data into a specific position of a table.
    Safely clones the XML of the row above or below to preserve complex formatting and avoid merged cells bugs.
    
    Args:
        file_path: The absolute path to the .docx file
        table_index: Index of the target table.
        row_index: The index where the new row should be inserted.
        row_data: A list of strings corresponding to each cell's text in the new row.
        out_path: Optional save path. If empty, overwrites the original file_path.
        
    Returns:
        Status message about the operation.
    """
    if not os.path.exists(file_path):
        return f"Error: File '{file_path}' not found."
    if not out_path:
        out_path = file_path

    temp_path = file_path + ".insert_row.tmp.zip"

    try:
        NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
        ns = {'w': NS}

        with zipfile.ZipFile(file_path, 'r') as zin:
            with zipfile.ZipFile(temp_path, 'w', zipfile.ZIP_DEFLATED) as zout:
                for item in zin.infolist():
                    data = zin.read(item.filename)

                    if item.filename == 'word/document.xml':
                        root = ET.fromstring(data)
                        all_tbl = root.findall('.//w:tbl', ns)
                        if table_index < 0 or table_index >= len(all_tbl):
                            if os.path.exists(temp_path):
                                os.remove(temp_path)
                            return f"Error: table_index {table_index} out of bounds (total {len(all_tbl)} tables)."

                        tbl = all_tbl[table_index]
                        tbl_tag = f'{{{NS}}}tr'
                        tr_list = [child for child in tbl if child.tag == tbl_tag]
                        num_rows = len(tr_list)

                        if row_index < 0 or row_index > num_rows:
                            if os.path.exists(temp_path):
                                os.remove(temp_path)
                            return f"Error: row_index {row_index} out of bounds (total {num_rows} rows)."

                        clone_idx = row_index - 1 if row_index > 0 else 0
                        ref_tr = tr_list[clone_idx]
                        new_tr = copy.deepcopy(ref_tr)

                        tc_tag = f'{{{NS}}}tc'
                        tc_list = [child for child in new_tr if child.tag == tc_tag]

                        for i, tc in enumerate(tc_list):
                            text_to_set = row_data[i] if i < len(row_data) else ""
                            all_t = tc.findall('.//w:t', ns)
                            injected = False
                            for t_elem in all_t:
                                if not injected:
                                    t_elem.text = text_to_set
                                    injected = True
                                else:
                                    t_elem.text = ""

                        tbl_children = list(tbl)
                        if row_index < num_rows:
                            anchor_tr = tr_list[row_index]
                            anchor_pos = tbl_children.index(anchor_tr)
                            tbl.insert(anchor_pos, new_tr)
                        else:
                            tbl.append(new_tr)

                        data = ET.tostring(root, encoding='UTF-8', xml_declaration=True)

                    zout.writestr(item, data)

        shutil.move(temp_path, out_path)
        return f"Successfully inserted a row into table {table_index} at index {row_index} with data: {row_data}. Saved to: {out_path}"

    except Exception as e:
        if os.path.exists(temp_path):
            os.remove(temp_path)
        return f"Error inserting row: {str(e)}"


@mcp.tool()
def delete_docx_table_row(file_path: str, table_index: int, row_index: int, out_path: str = "") -> str:
    """Deletes a specific row from a table."""
    if not os.path.exists(file_path):
        return f"Error: File '{file_path}' not found."
    if not out_path:
        out_path = file_path

    temp_path = file_path + ".delete_row.tmp.zip"

    try:
        NS = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
        ns = {'w': NS}

        with zipfile.ZipFile(file_path, 'r') as zin:
            with zipfile.ZipFile(temp_path, 'w', zipfile.ZIP_DEFLATED) as zout:
                for item in zin.infolist():
                    data = zin.read(item.filename)

                    if item.filename == 'word/document.xml':
                        root = ET.fromstring(data)
                        all_tbl = root.findall('.//w:tbl', ns)
                        if table_index < 0 or table_index >= len(all_tbl):
                            if os.path.exists(temp_path):
                                os.remove(temp_path)
                            return f"Error: table_index {table_index} out of bounds (total {len(all_tbl)} tables)."

                        tbl = all_tbl[table_index]
                        tbl_tag = f'{{{NS}}}tr'
                        tr_list = [child for child in tbl if child.tag == tbl_tag]
                        num_rows = len(tr_list)

                        if row_index < 0 or row_index >= num_rows:
                            if os.path.exists(temp_path):
                                os.remove(temp_path)
                            return f"Error: row_index {row_index} out of bounds (total {num_rows} rows)."

                        target_tr = tr_list[row_index]
                        tbl.remove(target_tr)

                        data = ET.tostring(root, encoding='UTF-8', xml_declaration=True)

                    zout.writestr(item, data)

        shutil.move(temp_path, out_path)
        return f"Successfully deleted row {row_index} from table {table_index}. Saved to: {out_path}"

    except Exception as e:
        if os.path.exists(temp_path):
            os.remove(temp_path)
        return f"Error deleting row: {str(e)}"
@mcp.tool()
def find_table_under_heading(file_path: str, target_heading: str) -> str:
    """
    Scans the document visually from top to bottom to find the given target_heading text,
    and returns the table_index of the VERY FIRST table that appears immediately after it.
    This avoids having to manually count tables.
    
    Args:
        file_path: The absolute path to the .docx file
        target_heading: The exact or partial text of the heading/paragraph to search for.
        
    Returns:
        A string containing the found table_index and a preview of the table's contents.
    """
    from docx.oxml.table import CT_Tbl
    from docx.oxml.text.paragraph import CT_P
    from docx.table import Table
    from docx.text.paragraph import Paragraph
    import os

    if not os.path.exists(file_path):
        return f"Error: File '{file_path}' not found."
        
    try:
        doc = Document(file_path)
        found_heading = False

        for child in doc.element.body.iterchildren():
            if isinstance(child, CT_P):
                p = Paragraph(child, doc)
                if target_heading in p.text:
                    found_heading = True
            elif isinstance(child, CT_Tbl):
                if found_heading:
                    # ✅ 핵심: XML 요소를 실제 doc.tables 인덱스로 변환
                    actual_index = get_actual_table_index(doc, child)

                    if actual_index == -1:
                        return f"Error: Could not map table to doc.tables index"

                    # 미리보기 제공
                    table = doc.tables[actual_index]
                    preview = get_table_preview(table, rows=3)

                    return f"🌟 Success! Found target heading '{target_heading}'.\nThe FIRST table after this heading is [Table Index: {actual_index}].\n\nPreview:\n{preview}"

        if not found_heading:
            return f"Error: Could not find any paragraph containing the text '{target_heading}'."
        else:
            return f"Error: Found the heading '{target_heading}', but there were no tables after it in the document."

    except Exception as e:
        return f"Error searching for table: {str(e)}"

@mcp.tool()
def list_all_tables(file_path: str) -> str:
    """
    Lists all tables in the document with their properties.

    Args:
        file_path: The absolute path to the .docx file

    Returns:
        A formatted string listing:
        - Table index (doc.tables 기준)
        - Location (본문/헤더/바닥글)
        - 행/열 수
        - 첫 3줄 미리보기
    """
    try:
        doc = Document(file_path)
        total_tables = len(doc.tables)

        if total_tables == 0:
            return "Document has no tables."

        # 헤더/바닥글 테이블 카운팅
        header_footer_count = 0
        for section in doc.sections:
            header_footer_count += len(section.header.tables)
            header_footer_count += len(section.footer.tables)

        results = []
        results.append(f"Document has {total_tables} tables:\n")

        for idx, table in enumerate(doc.tables):
            rows_count = len(table.rows)
            cols_count = len(table.columns) if table.columns else 0

            # 위치 판정
            location = identify_table_location(doc, idx)

            # 첫 3줄 미리보기
            preview = get_table_preview(table, rows=3)

            results.append(f"Table {idx} [{location}] ({rows_count} rows × {cols_count} cols)")
            if preview:
                # 미리보기를 들여쓰기
                preview_lines = preview.split('\n')
                for line in preview_lines:
                    results.append(f"  {line}")
            results.append("")  # 빈 줄 추가

        return "\n".join(results)

    except FileNotFoundError:
        return f"Error: File not found at {file_path}"
    except Exception as e:
        return f"Error listing tables: {str(e)}"

@mcp.tool()
def search_table_by_content(file_path: str, search_text: str, match_mode: str = "partial") -> str:
    """
    Searches for a table containing the specified text.

    Args:
        file_path: The absolute path to the .docx file
        search_text: Text to find
        match_mode: "partial" (포함) or "exact" (정확히)

    Returns:
        A formatted string with:
        - Table index
        - Cell location (row, col)
        - Cell content
        - Table preview
    """
    try:
        doc = Document(file_path)
        results = []
        found_count = 0

        for table_idx, table in enumerate(doc.tables):
            matches_in_table = []

            for row_idx, row in enumerate(table.rows):
                for col_idx, cell in enumerate(row.cells):
                    cell_text = cell.text.strip()

                    if match_mode == "partial":
                        if search_text in cell_text:
                            matches_in_table.append({
                                "row": row_idx,
                                "col": col_idx,
                                "content": format_cell_text(cell_text, 100)
                            })
                    elif match_mode == "exact":
                        if cell_text == search_text:
                            matches_in_table.append({
                                "row": row_idx,
                                "col": col_idx,
                                "content": format_cell_text(cell_text, 100)
                            })

            if matches_in_table:
                found_count += 1
                location = identify_table_location(doc, table_idx)
                results.append(f"\nTable {table_idx} [{location}]:")

                for match in matches_in_table:
                    results.append(f"  - Cell (Row {match['row']}, Col {match['col']}): \"{match['content']}\"")

                # 미리보기 추가
                preview = get_table_preview(table, rows=3)
                results.append("  Preview:")
                if preview:
                    preview_lines = preview.split('\n')
                    for line in preview_lines:
                        results.append(f"    {line}")

        if found_count == 0:
            return f"No tables found containing \"{search_text}\" (mode: {match_mode})"

        summary = f"Found {found_count} table(s) matching \"{search_text}\" (mode: {match_mode}):"
        return summary + "\n" + "\n".join(results)

    except FileNotFoundError:
        return f"Error: File not found at {file_path}"
    except Exception as e:
        return f"Error searching tables: {str(e)}"

if __name__ == "__main__":
    mcp.run()
