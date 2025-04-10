import mammoth
import re
import sys

def split_into_sections(text):
    """Split text into sections based on titles"""
    sections = []
    current_section = []
    lines = text.split('\n')
    
    for line in lines:
        line = line.strip()
        if not line:
            continue
            
        if line.isupper() and len(line.split()) > 1:
            if current_section:
                sections.append('\n'.join(current_section))
            current_section = [line]
        else:
            current_section.append(line)
    
    if current_section:
        sections.append('\n'.join(current_section))
    
    return sections

def clean_text(text):
    """Clean and format the extracted text"""
    # Fix broken URLs
    text = re.sub(r'https:\s*//', 'https://', text)
    text = re.sub(r'www\.\s*', 'www.', text)
    
    # Join broken lines (where a line ends with a word break)
    text = re.sub(r'(\w)\s*\n\s*(\w)', r'\1\2', text)
    
    # Remove multiple newlines
    text = re.sub(r'\n{3,}', '\n\n', text)
    
    # Remove spaces before punctuation
    text = re.sub(r'\s+([.,;:!?)])', r'\1', text)
    
    # Ensure proper spacing after punctuation
    text = re.sub(r'([.,;:!?])([^\s\n])', r'\1 \2', text)
    
    # Remove spaces at the start of lines
    text = re.sub(r'^\s+', '', text, flags=re.MULTILINE)
    
    # Fix list formatting
    text = re.sub(r'\n\s*-\s*', '\n- ', text)
    
    # Remove extra spaces
    text = re.sub(r' +', ' ', text)
    
    # Fix email addresses
    text = re.sub(r'(\w+)\s*\.\s*(\w+)\s*@', r'\1.\2@', text)
    
    # Fix broken lines that end with hyphen
    text = re.sub(r'(\w)-\s*\n\s*(\w)', r'\1\2', text)
    
    # Split into sections and process each section
    sections = split_into_sections(text)
    processed_sections = []
    
    for section in sections:
        # Split section into paragraphs
        paragraphs = re.split(r'\n\s*\n', section)
        processed_paragraphs = []
        
        for para in paragraphs:
            # Join lines within paragraph
            lines = para.split('\n')
            if any(line.isupper() and len(line.split()) > 1 for line in lines):
                processed_paragraphs.append('\n'.join(lines))
            else:
                processed_paragraphs.append(' '.join(line.strip() for line in lines))
        
        processed_sections.append('\n\n'.join(processed_paragraphs))
    
    return '\n\n'.join(processed_sections)

def format_table_data(data):
    """Format data as a markdown table"""
    if not data:
        return ""
        
    # Find the maximum width of each column
    col_widths = [max(len(str(row[i])) for row in data) for i in range(len(data[0]))]
    
    # Create the header row
    header = "| " + " | ".join(f"**{str(cell)}**".ljust(width) for cell, width in zip(data[0], col_widths)) + " |"
    
    # Create the separator row
    separator = "|" + "|".join("-" * (width + 2) for width in col_widths) + "|"
    
    # Create the data rows
    rows = []
    for row in data[1:]:
        rows.append("| " + " | ".join(str(cell).ljust(width) for cell, width in zip(row, col_widths)) + " |")
    
    return "\n".join([header, separator] + rows)

def convert_to_markdown(text):
    """Convert text to markdown format"""
    sections = text.split('\n\n')
    markdown_sections = []
    table_data = []
    in_table = False
    
    for i, section in enumerate(sections):
        section = section.strip()
        if not section:
            continue
            
        # Convert main title
        if i == 0 and section.isupper():
            markdown_sections.append(f"# {section}\n")
            continue
            
        # Convert section titles
        if section.isupper() and len(section.split()) > 1:
            if markdown_sections:
                markdown_sections.append("\n---\n")
            markdown_sections.append(f"\n## {section}")
            continue
            
        # Handle table-like data
        if ':' in section and len(section.split('\n')) == 1:
            title, value = section.split(':', 1)
            title = title.strip()
            value = value.strip()
            
            if title.istitle() and not title.isupper():
                if not in_table:
                    in_table = True
                    table_data = [["Field", "Value"]]
                table_data.append([title, value])
                continue
            
        # If we were in a table and now we're not, format the table
        if in_table and (not ':' in section or len(section.split('\n')) > 1):
            in_table = False
            if table_data:
                markdown_sections.append(format_table_data(table_data))
                table_data = []
        
        # Convert URLs to markdown links
        section = re.sub(r'(https?://[^\s)]+)', lambda m: f"[{m.group(1)}]({m.group(1)})", section)
        
        # Convert email addresses to markdown links
        section = re.sub(r'(\S+@\S+\.\S+)', lambda m: f"[{m.group(1)}](mailto:{m.group(1)})", section)
        
        # Add emphasis to French text
        section = re.sub(r'"([^"]*appels à manifestation d\'intérêt[^"]*)"', r'*"\1"*', section)
        
        # Format lists
        if section.startswith('- '):
            lines = section.split('\n')
            section = '\n'.join('- ' + line.strip().lstrip('- ') for line in lines if line.strip())
        
        markdown_sections.append(section)
    
    # Format any remaining table
    if in_table and table_data:
        markdown_sections.append(format_table_data(table_data))
    
    # Join sections with double newlines and ensure proper spacing
    text = '\n\n'.join(markdown_sections)
    
    # Fix spacing around horizontal rules
    text = re.sub(r'\n+---\n+', '\n\n---\n\n', text)
    
    return text

def extract_text_mammoth(file_path, output_path):
    """Extract all text from DOCX using mammoth with enhanced formatting"""
    try:
        with open(file_path, "rb") as docx_file:
            result = mammoth.extract_raw_text(docx_file)
            text = result.value
            
            # Clean and format the text
            text = clean_text(text)
            
            # Convert to markdown
            markdown_text = convert_to_markdown(text)
            
            # Save to output file
            with open(output_path, 'w', encoding='utf-8') as out_file:
                out_file.write(markdown_text)
            
            return markdown_text
    except Exception as e:
        print(f"Error processing file: {str(e)}", file=sys.stderr)
        sys.exit(1)

if __name__ == '__main__':
    if len(sys.argv) != 2:
        print("Usage: python extract_text_mammoth.py <input_file>")
        sys.exit(1)
    
    output_path = 'output.md'
    text = extract_text_mammoth(sys.argv[1], output_path)
    print(f"Markdown text has been saved to {output_path}")
    print("\nPreview of the first 500 characters:")
    print("-" * 80)
    print(text[:500])
    print("-" * 80)