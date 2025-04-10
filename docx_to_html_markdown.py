import sys
import mammoth
import html2text
from pathlib import Path
from tabulate import tabulate
from bs4 import BeautifulSoup
import re


def format_table_markdown(table_data):
    """Format table data as markdown using tabulate"""
    if not table_data:
        return ""
    
    # Clean and prepare table data
    cleaned_data = []
    for row in table_data:
        # Clean cells in the row but preserve all cells
        cleaned_row = []
        for cell in row:
            # Remove newlines and extra spaces
            cell = " ".join(cell.split())
            cleaned_row.append(cell)
        
        cleaned_data.append(cleaned_row)
    
    if not cleaned_data:
        return ""
    
    # Format table using tabulate
    table = tabulate(cleaned_data, headers="firstrow", tablefmt="pipe")
    
    # Add empty lines before and after table
    return f"\n{table}\n"


def extract_tables_from_html(html_content):
    """Extract tables from HTML content using BeautifulSoup"""
    soup = BeautifulSoup(html_content, 'html.parser')
    tables = []
    
    for table in soup.find_all('table'):
        table_data = []
        for row in table.find_all('tr'):
            row_data = []
            for cell in row.find_all(['td', 'th']):
                # Clean cell text but preserve empty cells
                text = cell.get_text(strip=True)
                row_data.append(text)
            
            # Include all rows, even if they appear empty
            table_data.append(row_data)
        
        if table_data:  # Only add non-empty tables
            tables.append(table_data)
    
    return tables


def convert_docx_to_html(docx_path):
    """Convert DOCX file to HTML string"""
    try:
        with open(docx_path, 'rb') as docx_file:
            result = mammoth.convert_to_html(docx_file)
            return result.value
    except Exception as e:
        print(f"Error converting DOCX to HTML: {str(e)}", file=sys.stderr)
        sys.exit(1)


def convert_html_to_markdown(html_content):
    """Convert HTML string to Markdown"""
    try:
        # Extract tables first
        tables = extract_tables_from_html(html_content)
        
        # Replace tables with placeholders
        soup = BeautifulSoup(html_content, 'html.parser')
        table_placeholders = []
        for table in soup.find_all('table'):
            placeholder = f'TABLE_PLACEHOLDER_{len(table_placeholders)}'
            table_placeholders.append(placeholder)
            table.replace_with(placeholder)
        
        # Configure html2text
        converter = html2text.HTML2Text()
        converter.body_width = 0  # Disable line wrapping
        converter.protect_links = True  # Don't convert links to references
        converter.unicode_snob = True  # Use Unicode characters
        converter.single_line_break = True  # Use single line breaks
        converter.tables = False  # Disable table formatting
        converter.images_to_alt = True  # Convert images to alt text
        
        # Convert to Markdown
        markdown = converter.handle(str(soup))
        
        # Replace placeholders with formatted tables
        for i, placeholder in enumerate(table_placeholders):
            if i < len(tables):
                markdown = markdown.replace(placeholder, format_table_markdown(tables[i]))
        
        # Add extra newlines after headers
        markdown = re.sub(r'(#+\s+.*)\n', r'\1\n\n', markdown)
        
        # Clean up multiple newlines
        markdown = re.sub(r'\n{3,}', '\n\n', markdown)
        
        return markdown
    except Exception as e:
        print(f"Error converting HTML to Markdown: {str(e)}", file=sys.stderr)
        sys.exit(1)


def process_file(input_path, output_path=None):
    """Process DOCX file and convert it to Markdown"""
    try:
        # Convert input path to Path object
        input_path = Path(input_path)
        
        # If output path is not specified, use input filename with .md extension
        if output_path is None:
            output_path = input_path.with_suffix('.md')
        
        # Convert DOCX to HTML
        print("Converting DOCX to HTML...")
        html_content = convert_docx_to_html(input_path)
        
        # Convert HTML to Markdown
        print("Converting HTML to Markdown...")
        markdown_content = convert_html_to_markdown(html_content)
        
        # Write output
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write(markdown_content)
            
        print(f"\nConversion completed successfully!")
        print(f"Output saved to: {output_path}")
        
        # Print preview
        print("\nPreview of the first few lines:")
        print("-" * 80)
        lines = markdown_content.split('\n')[:10]
        print('\n'.join(lines))
        print("-" * 80)
        
    except Exception as e:
        print(f"Error processing file: {str(e)}", file=sys.stderr)
        sys.exit(1)


if __name__ == '__main__':
    if len(sys.argv) < 2:
        print("Usage: python docx_to_html_markdown.py <input_file> [output_file]")
        sys.exit(1)
    
    input_file = sys.argv[1]
    output_file = sys.argv[2] if len(sys.argv) > 2 else None
    
    process_file(input_file, output_file) 