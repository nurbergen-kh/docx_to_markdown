import re

def format_table(rows):
    if not rows:
        return ""
    
    # Remove empty rows and clean whitespace
    rows = [[cell.strip() for cell in row] for row in rows if any(cell.strip() for cell in row)]
    if not rows:
        return ""

    # Get max width for each column
    col_widths = []
    for col in range(len(rows[0])):
        width = max(len(str(row[col])) for row in rows if col < len(row))
        col_widths.append(width)

    # Format the table
    table_lines = []
    for i, row in enumerate(rows):
        # Format cells with proper padding
        cells = []
        for j, cell in enumerate(row):
            if j < len(col_widths):
                cells.append(str(cell).ljust(col_widths[j]))
        
        # Add the row
        table_lines.append(f"| {' | '.join(cells)} |")
        
        # Add separator after header (first row)
        if i == 0:
            separators = ['-' * width for width in col_widths]
            table_lines.append(f"| {' | '.join(separators)} |")

    return '\n'.join(table_lines)

def clean_text(text):
    # Fix URLs by removing spaces
    text = re.sub(r'(https?:\s*/\s*/[^\s]+)\s*', lambda m: m.group(1).replace(' ', ''), text)
    
    # Italicize French text
    text = re.sub(r'["\'](appels à manifestation d\'intérêt)["\']', r'*\1*', text)
    
    # Remove tab characters
    text = text.replace('\t', '')
    
    # Remove multiple consecutive newlines
    text = re.sub(r'\n{3,}', '\n\n', text)
    
    return text.strip()

def process_table(table):
    rows = []
    for row in table.rows:
        cells = [clean_text(cell.text) for cell in row.cells if cell.text.strip()]
        if cells:  # Only add non-empty rows
            rows.append(cells)
    
    if not rows:
        return ""
    
    # Create table header
    md = []
    header = rows[0]
    md.append('| ' + ' | '.join(header) + ' |')
    md.append('|' + '|'.join(['---' for _ in header]) + '|')
    
    # Add remaining rows
    for row in rows[1:]:
        # Ensure row has same number of columns as header
        while len(row) < len(header):
            row.append('')
        md.append('| ' + ' | '.join(row) + ' |')
    
    return '\n'.join(md) + '\n\n'

def process_paragraph(paragraph):
    text = clean_text(paragraph.text)
    if not text:
        return ""
        
    # Check if it's a heading by style
    style = paragraph.style.name.lower()
    if 'heading' in style or 'title' in style:
        level = 1
        if any(str(i) in style for i in range(1, 7)):
            level = next(i for i in range(1, 7) if str(i) in style)
        return '#' * level + ' ' + text + '\n\n'
    
    return text + '\n\n'

def table_to_markdown(table):
    if not table.rows:
        return ''
    
    # For single-cell tables, convert to a header
    if len(table.rows) == 1 and len(table.rows[0]) == 1:
        return f"\n### {table.rows[0][0].strip()}\n"
    
    # Regular table processing
    markdown_rows = []
    for i, row in enumerate(table.rows):
        cells = [cell.strip() for cell in row]
        row_text = f"| {' | '.join(cells)} |"
        markdown_rows.append(row_text)
        
        # Add separator after header
        if i == 0:
            separator = '|' + '|'.join('-' * len(cell) for cell in cells) + '|'
            markdown_rows.append(separator)
    
    return '\n'.join(markdown_rows) + '\n'
