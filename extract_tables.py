import sys
from docx2python import docx2python
import json
import re

def clean_cell(cell):
    """Clean cell content by removing extra brackets, quotes and whitespace"""
    if not isinstance(cell, str):
        return ""
    
    # Remove array markers and quotes
    text = re.sub(r'^\[\'|\'\]$', '', cell)
    text = re.sub(r'\',\s*\'', ' ', text)
    
    # Remove extra quotes and brackets
    text = text.strip('[]"\' ')
    
    # Clean whitespace
    text = re.sub(r'\s+', ' ', text)
    text = text.strip()
    
    return text

def print_block_structure(block, level=0):
    """Print the structure of a block for debugging"""
    indent = "  " * level
    if isinstance(block, (list, tuple)):
        print(f"{indent}List/Tuple of length {len(block)}:")
        for item in block[:2]:  # Print first 2 items only
            print_block_structure(item, level + 1)
        if len(block) > 2:
            print(f"{indent}  ...")
    else:
        print(f"{indent}Value ({type(block)}): {str(block)[:100]}")

def extract_tables(file_path, output_path):
    """Extract raw table data from DOCX file and save to JSON"""
    try:
        # Extract document content
        doc = docx2python(file_path)
        
        print("\nDocument structure:")
        print("-" * 80)
        print_block_structure(doc.body)
        print("-" * 80)
        
        # Store all tables
        tables_data = []
        table_count = 0
        
        # Process each block in the document
        def process_block(block):
            nonlocal table_count
            if isinstance(block, (list, tuple)):
                # Check if this block looks like a table
                if len(block) > 0 and all(isinstance(row, (list, tuple)) for row in block):
                    table_count += 1
                    print(f"\nPotential table #{table_count}:")
                    print(f"Number of rows: {len(block)}")
                    if len(block) > 0:
                        print(f"Sample row structure: {[type(cell) for cell in block[0]]}")
                    
                    # Convert all cells to strings and remove empty rows
                    table = []
                    seen_rows = set()  # For deduplication
                    
                    for row in block:
                        if isinstance(row, (list, tuple)):
                            # Clean each cell in the row
                            cleaned_row = [clean_cell(cell) for cell in row]
                            # Remove empty cells
                            cleaned_row = [cell for cell in cleaned_row if cell]
                            if cleaned_row:
                                # Convert to tuple for hashing
                                row_tuple = tuple(cleaned_row)
                                if row_tuple not in seen_rows:
                                    table.append(cleaned_row)
                                    seen_rows.add(row_tuple)
                    
                    if table:  # Only add non-empty tables
                        print(f"Added table with {len(table)} rows after cleaning")
                        tables_data.append(table)
                    else:
                        print("Table was empty after cleaning")
                
                # Recursively process nested blocks
                for item in block:
                    process_block(item)
        
        # Start processing from the root
        process_block(doc.body)
        
        # Save tables to JSON file
        with open(output_path, 'w', encoding='utf-8') as f:
            json.dump(tables_data, f, ensure_ascii=False, indent=2)
            
        print(f"\nFound {len(tables_data)} tables")
        print(f"Table data has been saved to {output_path}")
        
        # Print preview of first table
        if tables_data:
            print("\nPreview of first table:")
            print("-" * 80)
            for row in tables_data[0][:3]:  # Show first 3 rows
                print(row)
            if len(tables_data[0]) > 3:
                print("...")
            print("-" * 80)
        
    except Exception as e:
        print(f"Error processing file: {str(e)}", file=sys.stderr)
        sys.exit(1)

if __name__ == '__main__':
    if len(sys.argv) != 2:
        print("Usage: python extract_tables.py <input_file>")
        sys.exit(1)
    
    input_file = sys.argv[1]
    output_file = 'tables_data.json'
    extract_tables(input_file, output_file) 