from mcp.server.fastmcp import FastMCP
from docx import Document
import os

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

if __name__ == "__main__":
    mcp.run()
