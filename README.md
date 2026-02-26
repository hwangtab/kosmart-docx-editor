# kosmart-docx-editor (MCP Server)

A custom Model Context Protocol (MCP) server designed to perfectly read and edit Microsoft Word `(.docx)` files, specifically engineered to bypass the infamous merged cells (infinite loop/index error) bug in `python-docx` using a Flat Index algorithm.

## Features
- **`read_docx_table_flat_index`**: Scans any complex `.docx` table mapping all unique cells into a flat 1D array, preventing duplicate grid representations.
- **`safe_replace_docx_cell`**: Replaces the text inside a targeted `.docx` cell without destroying the underlying XML paragraph formatting.

## Installation for AI Agents (Claude / Cursor)

### Prerequisites
Make sure you have Python 3.10+ installed on your system.

### Option 1: Claude Code (CLI)
You can directly add this MCP server to your Claude Code environment using the following command:

```bash
git clone https://github.com/hwangtab/kosmart-docx-editor.git
cd kosmart-docx-editor
python3 -m venv venv
source venv/bin/activate
pip install -r requirements.txt

# Add globally to Claude
claude mcp add -s global kosmart_docx_editor $(pwd)/venv/bin/python $(pwd)/server.py
```

### Option 2: Cursor / VSCode
1. Git clone this repository and install dependencies in a virtual environment as shown above.
2. In Cursor/VSCode MCP settings, add a new command tool:
   - **Name**: `kosmart_docx_editor`
   - **Command**: `/path/to/venv/bin/python /path/to/server.py`

## Why This Exists?
`python-docx` struggles with MS Word tables that contain merged cells, returning duplicate cell objects based on the layout grid rather than the unique cell structure. This often leads AI agents to infinite loops or `IndexError` when creating formatting replacement pipelines. This MCP server acts as an ultimate workaround.

## License
MIT License
