#!/usr/bin/env python3
"""
CV Template Generator

This script takes a CV template with placeholders and fills them with provided values.
Placeholders should be in the format: {{PLACEHOLDER_NAME}}

Usage:
    python cv_generator.py --template template.txt --output output.txt --data data.json
    python cv_generator.py --template template.txt --output output.txt --name "John Doe" --email "john@example.com"
"""

import argparse
import json
import sys
import re
from pathlib import Path
from copy import deepcopy
from docx import Document
from docx.oxml import OxmlElement
from docx.shared import Pt, RGBColor


def load_template(template_path):
    """Load the template file (supports .txt and .docx)."""
    template_path = Path(template_path)
    
    if not template_path.exists():
        print(f"Error: Template file '{template_path}' not found.", file=sys.stderr)
        sys.exit(1)
    
    try:
        if template_path.suffix.lower() == '.docx':
            return Document(template_path)
        else:
            with open(template_path, 'r', encoding='utf-8') as f:
                return f.read()
    except Exception as e:
        print(f"Error reading template: {e}", file=sys.stderr)
        sys.exit(1)


def load_data_from_json(json_path):
    """Load replacement data from JSON file."""
    try:
        with open(json_path, 'r', encoding='utf-8') as f:
            return json.load(f)
    except FileNotFoundError:
        print(f"Error: Data file '{json_path}' not found.", file=sys.stderr)
        sys.exit(1)
    except json.JSONDecodeError as e:
        print(f"Error parsing JSON: {e}", file=sys.stderr)
        sys.exit(1)


def parse_markdown_formatting(text):
    """
    Parse markdown-style formatting in text.
    Returns a list of tuples: (text, is_bold)
    
    Example: "*bold* text" -> [("bold", True), (" text", False)]
    """
    parts = []
    pattern = r'\*([^*]+)\*'
    last_end = 0
    
    for match in re.finditer(pattern, text):
        # Add text before the bold part
        if match.start() > last_end:
            parts.append((text[last_end:match.start()], False))
        # Add the bold part
        parts.append((match.group(1), True))
        last_end = match.end()
    
    # Add remaining text
    if last_end < len(text):
        parts.append((text[last_end:], False))
    
    return parts if parts else [(text, False)]


def set_paragraph_text_with_formatting(paragraph, text, run_font_properties=None):
    """Set text of a paragraph applying markdown formatting and preserving run properties."""
    # Clear existing runs
    for run in paragraph.runs:
        run.clear()
    
    # Apply new text with formatting
    parts = parse_markdown_formatting(text)
    for i, (part_text, is_bold) in enumerate(parts):
        run = paragraph.add_run(part_text)
        if run_font_properties:
            run.bold = run_font_properties.bold
            run.italic = run_font_properties.italic
            run.underline = run_font_properties.underline
            if run_font_properties.name:
                run.font.name = run_font_properties.name
            if run_font_properties.size:
                run.font.size = run_font_properties.size
            if run_font_properties.color.rgb:
                run.font.color.rgb = run_font_properties.color.rgb
        run.bold = is_bold # Apply or override bold based on markdown


def add_bullet_style(paragraph_element):
    """Add bullet point style to a paragraph element."""
    nsmap = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
    pPr = paragraph_element.find('w:pPr', namespaces=nsmap)
    if pPr is None:
        pPr = OxmlElement('w:pPr')
        paragraph_element.insert(0, pPr)
    
    numPr = pPr.find('w:numPr', namespaces=nsmap)
    if numPr is None:
        numPr = OxmlElement('w:numPr')
        pPr.append(numPr)
    
    ilvl = numPr.find('w:ilvl', namespaces=nsmap)
    if ilvl is None:
        ilvl = OxmlElement('w:ilvl')
        ilvl.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val', '0')
        numPr.append(ilvl)
        
    numId = numPr.find('w:numId', namespaces=nsmap)
    if numId is None:
        numId = OxmlElement('w:numId')
        numId.set('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val', '1') # Assuming '1' is a bullet list
        numPr.append(numId)


def fill_docx_template(doc, data):
    """Fill DOCX template placeholders with data values, supporting markdown formatting and repeating blocks."""
    
    # Process repeating blocks first
    
    all_paragraphs = list(doc.paragraphs)
    
    i = 0
    while i < len(all_paragraphs):
        para = all_paragraphs[i]
        
        match = re.search(r'{{#(\w+)}}', para.text)
        if not match:
            i += 1
            continue

        block_name = match.group(1)
        
        if not (block_name in data and isinstance(data[block_name], list)):
            i+= 1
            continue

        # Find the closing paragraph
        closing_idx = -1
        for j in range(i + 1, len(all_paragraphs)):
            if f'{{{{/{block_name}}}}}' in all_paragraphs[j].text:
                closing_idx = j
                break
        
        if closing_idx == -1:
            i += 1
            continue
            
        opening_para = all_paragraphs[i]
        template_paras = all_paragraphs[i+1:closing_idx]
        
        # Remove template paragraphs including opening and closing markers
        # Store a reference to the element before the opening paragraph for insertion
        anchor_element = opening_para._element.getprevious()
        
        for p_idx in range(i, closing_idx + 1):
            p = all_paragraphs[p_idx]
            p._element.getparent().remove(p._element)

        # Collect all generated paragraphs for this block
        generated_paragraph_elements = []
        for item_data in data[block_name]:
            for template_para in template_paras:
                new_p = doc.add_paragraph()
                # Copy properties using docx.oxml elements for robustness
                new_p._p = deepcopy(template_para._p)
                
                # Replace placeholders in the copied paragraph
                current_para_content = "".join([run.text for run in template_para.runs])
                filled_content_with_placeholders = current_para_content # Start with template text
                
                for key, value in item_data.items():
                    placeholder = f"{{{{{key}}}}}"
                    if placeholder in filled_content_with_placeholders:
                        # Convert list values to newline separated string for initial replacement
                        # This will be further processed by expand_list_placeholders
                        if isinstance(value, list):
                            # This should not happen with our flattened data structure for repeating blocks
                            # But as a fallback, join them
                            filled_content_with_placeholders = filled_content_with_placeholders.replace(placeholder, "\n".join(value))
                        else:
                            filled_content_with_placeholders = filled_content_with_placeholders.replace(placeholder, str(value))
                
                # Set text and reapply formatting
                # Note: `set_paragraph_text_with_formatting` clears all runs and re-adds them.
                # If we want to preserve complex run-level formatting from template_para,
                # we'd need a more intricate copy or run-by-run replacement.
                # For now, we assume primary formatting is at paragraph level (style, alignment, tab stops)
                # and font properties are consistent within the runs of a template_para.
                set_paragraph_text_with_formatting(new_p, filled_content_with_placeholders, template_para.runs[0].font if template_para.runs else None)

                generated_paragraph_elements.append(new_p._p)
        
        # Insert all generated paragraphs after the anchor element
        # If anchor_element is None, it means the opening_para was the first in the document
        # In this case, prepend to the document body.
        if anchor_element is None:
            # Need to find the document body element
            body = doc._body
            for p_elem in reversed(generated_paragraph_elements): # Insert in reverse to maintain order
                body.insert(0, p_elem)
        else:
            current_anchor_for_insertion = anchor_element
            for p_elem in generated_paragraph_elements:
                current_anchor_for_insertion.addnext(p_elem)
                current_anchor_for_insertion = p_elem
            
        # Rebuild paragraphs list and restart scan
        all_paragraphs = list(doc.paragraphs)
        i = 0
        continue # Continue outer loop from the start after modifying doc.paragraphs

    # Now, expand list placeholders (KEY_ACHIEVEMENTS, TECHNICAL_STACK, DESCRIPTION)
    # This must be done after repeating blocks are processed, as these can contain lists too
    
    paragraphs_to_process = list(doc.paragraphs) # Get a fresh list
    for para in paragraphs_to_process:
        for key in data:
            if isinstance(data[key], list) and (
               (key == 'KEY_ACHIEVEMENTS' and f"{{{{{key}}}}}" in para.text) or
               (key == 'TECHNICAL_STACK' and f"{{{{{key}}}}}" in para.text) or
               (key == 'DESCRIPTION' and f"{{{{{key}}}}}" in para.text)
            ):
                placeholder = f"{{{{{key}}}}}"
                if placeholder in para.text:
                    list_items = data[key]
                    if list_items:
                        # Get original paragraph style and properties
                        original_style = para.style
                        original_pPr_element = deepcopy(para._p.pPr) if para._p.pPr is not None else OxmlElement('w:pPr')
                        original_font_props = para.runs[0].font if para.runs else None
                        
                        # Replace the placeholder paragraph with the first item
                        # Need to clear existing runs and add new ones for the first item
                        for run in para.runs:
                            run.clear()
                        set_paragraph_text_with_formatting(para, str(list_items[0]), original_font_props)

                        para.style = original_style # Reapply original style
                        
                        # Clear existing pPr if any and re-append the original pPr element
                        existing_pPr = para._p.pPr
                        if existing_pPr is not None:
                            para._p.remove(existing_pPr)
                        para._p.insert(0, original_pPr_element)
                        
                        # Insert remaining items as new paragraphs with same style and properties
                        current_anchor_elem = para._p
                        for item in list_items[1:]:
                            new_para = doc.add_paragraph()
                            new_para.style = original_style # Copy style
                            
                            # Copy pPr element
                            new_para._p.insert(0, deepcopy(original_pPr_element))
                            
                            set_paragraph_text_with_formatting(new_para, str(item), original_font_props)
                            current_anchor_elem.addnext(new_para._p)
                            current_anchor_elem = new_para._p
                    else:
                        # If list is empty, remove placeholder
                        set_paragraph_text_with_formatting(para, para.text.replace(placeholder, ""))
    
    # Finally, fill any remaining simple top-level placeholders
    # Need to re-list paragraphs as they might have changed due to block processing
    for para in doc.paragraphs:
        for key, value in data.items():
            if not isinstance(value, (list, dict)): # Only fill simple placeholders
                set_paragraph_text_with_formatting(para, para.text.replace(f"{{{{{key}}}}}", str(value)), para.runs[0].font if para.runs else None)

    for table in doc.tables:
        for row in table.rows:
            for cell in row.cells:
                for para in cell.paragraphs:
                    for key, value in data.items():
                        if not isinstance(value, (list, dict)): # Only fill simple placeholders
                            set_paragraph_text_with_formatting(para, para.text.replace(f"{{{{{key}}}}}", str(value)), para.runs[0].font if para.runs else None)
    
    return doc


def main():
    parser = argparse.ArgumentParser(
        description='Fill CV template with provided data',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # Using JSON data file
  python cv_generator.py --template cv_template.txt --output cv_filled.txt --data data.json
  
  # Using command-line arguments
  python cv_generator.py --template cv_template.txt --output cv_filled.txt --name "John Doe" --email "john@example.com"
  
  # Mix JSON file and command-line arguments (command-line takes precedence)
  python cv_generator.py --template cv_template.txt --output cv_filled.txt --data data.json --name "Jane Smith"
        """
    )
    
    parser.add_argument('--template', required=True, help='Path to the CV template file')
    parser.add_argument('--output', required=True, help='Path to save the filled CV')
    parser.add_argument('--data', help='Path to JSON file with replacement data')
    
    args = parser.parse_args()
    
    # Load template
    template_doc = load_template(args.template)
    
    # Initialize data dictionary
    data = {}
    
    # Load from JSON file if provided
    if args.data:
        data = load_data_from_json(args.data)
    
    # Fill the template
    filled_doc = fill_docx_template(template_doc, data)
    
    # Output result
    try:
        output_path = Path(args.output)
        output_path.parent.mkdir(parents=True, exist_ok=True)
        filled_doc.save(str(output_path))
        print(f"✓ CV filled successfully: {args.output}")
    except Exception as e:
        print(f"Error writing output: {e}", file=sys.stderr)
        sys.exit(1)


if __name__ == '__main__':
    main()