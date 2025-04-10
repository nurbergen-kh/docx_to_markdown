import sys
from docx import Document
from docx2python import docx2python
import re
import textwrap


def clean_text(text):
    """Clean and format text"""
    if not isinstance(text, str):
        return ""
    
    # Remove extra whitespace
    text = re.sub(r'\s+', ' ', text)
    text = text.strip()
    
    # Convert HTML links to markdown
    text = re.sub(r'<a href="([^"]+)">([^<]+)</a>', r'[\2](\1)', text)
    
    # Fix broken URLs
    text = re.sub(r'https:\s*//', 'https://', text)
    text = re.sub(r'www\.\s*', 'www.', text)
    
    # Fix email addresses
    text = re.sub(r'(\w+)\s*\.\s*(\w+)\s*@', r'\1.\2@', text)
    
    # Format URLs as markdown links (if not already formatted)
    text = re.sub(r'\[(https?://[^\]]+)\]\([^)]+\)[.)\s]*', r'[\1](\1)', text)
    text = re.sub(r'(?<!\[)(https?://[^\s]+)(?!\])', r'[\1](\1)', text)
    
    return text

def wrap_text(text, width=100):
    """Wrap text to specified width while preserving markdown formatting"""
    # Don't wrap URLs
    parts = []
    last_end = 0
    for match in re.finditer(r'\[([^\]]+)\]\([^)]+\)', text):
        # Add wrapped text before the URL
        if match.start() > last_end:
            parts.extend(textwrap.wrap(text[last_end:match.start()], width=width))
        # Add the URL unchanged
        parts.append(text[match.start():match.end()])
        last_end = match.end()
    
    # Add remaining text
    if last_end < len(text):
        parts.extend(textwrap.wrap(text[last_end:], width=width))
    
    return '\n'.join(parts)

def extract_tables(file_path):
    """Extract tables from DOCX file"""
    doc = Document(file_path)
    tables = []
    
    for table in doc.tables:
        rows = []
        for row in table.rows:
            columns = []
            text = set()
            for cell in row.cells:
                if cell.text.strip():
                    if cell.text not in text:
                        columns.append(cell.text)
                        text.add(cell.text)
            if columns:
                rows.append(columns)

        tables.append(rows)
    
    return tables

def format_table_as_markdown(table_data):
    """Format table data as markdown table"""
    if not table_data:
        return ""
    
    # Find the maximum number of columns
    max_cols = max(len(row) for row in table_data)
    
    # Format rows
    rows = []
    for i, row in enumerate(table_data):
        # Pad row to have the same number of columns
        padded_row = row + [""] * (max_cols - len(row))
        
        # Format row with separators
        formatted_row = "| " + " | ".join(padded_row) + " |"
        rows.append(formatted_row)
        
        # Add separator after header row
        if i == 0:
            separator = "|" + "|".join(["---" for _ in range(max_cols)]) + "|"
            rows.append(separator)
    
    return "\n---\n\n" + "\n".join(rows) + "\n\n"

def extract_text(file_path, output_path):
    """Extract text from DOCX file and save as markdown"""
    try:
        # Extract document content
        doc = docx2python(file_path)
        
        # Extract tables
        tables = extract_tables(file_path)
        
        # Create a set of all text from tables to skip them during processing
        table_texts = set()
        for table in tables:
            for row in table:
                for cell in row:
                    if cell.strip():
                        table_texts.add(cell.strip())
        
        # Process and format text
        markdown_lines = []
        seen_paragraphs = set()  # To avoid duplicates
        
        def process_block(block, level=0):
            if isinstance(block, (list, tuple)):
                # Process nested blocks
                for item in block:
                    process_block(item, level + 1)
            else:
                text = clean_text(str(block))
                if text and text not in seen_paragraphs and text not in table_texts:
                    # Wrap long paragraphs
                    wrapped_text = wrap_text(text)
                    markdown_lines.append(f"{wrapped_text}\n\n")
                    seen_paragraphs.add(text)
        
        # Process document content
        process_block(doc.body)
        
        # Add tables to the output after processing all text
        for table in tables:
            markdown_lines.append(format_table_as_markdown(table))
        
        # Clean up multiple newlines and spaces around URLs
        content = ''.join(markdown_lines)
        content = re.sub(r'\n{3,}', '\n\n', content)
        content = re.sub(r'\s+\[', ' [', content)
        content = re.sub(r'\]\s+', '] ', content)
        
        # Write to output file
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write(content)
            
        print(f"Text has been extracted and saved to {output_path}")
        
        # Print preview
        print("\nPreview of the first few lines:")
        print("-" * 80)
        with open(output_path, 'r', encoding='utf-8') as f:
            preview = ''.join(f.readlines()[:10])
            print(preview)
        print("-" * 80)
        
    except Exception as e:
        print(f"Error processing file: {str(e)}", file=sys.stderr)
        sys.exit(1)

if __name__ == '__main__':
    if len(sys.argv) != 2:
        print("Usage: python extract_text_docx2python.py <input_file>")
        sys.exit(1)
    
    input_file = sys.argv[1]
    output_file = 'output_docx2python.md'
    extract_text(input_file, output_file) 