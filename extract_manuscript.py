#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Extract manuscript text from .docx file (docx is a ZIP archive with XML)
"""
import zipfile
import xml.etree.ElementTree as ET
import os
from pathlib import Path

def extract_docx_text(docx_path):
    """Extract all text from a .docx file"""
    try:
        with zipfile.ZipFile(docx_path, 'r') as zip_ref:
            # Read the main document XML
            xml_content = zip_ref.read('word/document.xml')
            root = ET.fromstring(xml_content)

            # Define namespace
            ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}

            # Extract all text nodes
            text_elements = root.findall('.//w:t', ns)
            text_parts = [elem.text for elem in text_elements if elem.text]

            full_text = ''.join(text_parts)
            return full_text
    except Exception as e:
        print(f"Error extracting text: {e}")
        return None

def extract_docx_structure(docx_path):
    """Extract document structure with paragraphs"""
    try:
        with zipfile.ZipFile(docx_path, 'r') as zip_ref:
            xml_content = zip_ref.read('word/document.xml')
            root = ET.fromstring(xml_content)

            ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}

            # Get all paragraphs
            paragraphs = root.findall('.//w:p', ns)

            content = []
            for para in paragraphs:
                text_nodes = para.findall('.//w:t', ns)
                para_text = ''.join([t.text for t in text_nodes if t.text])
                if para_text:
                    content.append(para_text)

            return content
    except Exception as e:
        print(f"Error extracting structure: {e}")
        return []

# Extract from target file
source_file = "d:/Workplace/SCS/revised_manuscript_2006_2024_noER_SCS_deep_revised.docx"
output_file = "d:/Workplace/SCS/manuscript_extracted.txt"

if os.path.exists(source_file):
    print(f"Extracting from: {source_file}")

    # Extract structured content
    paragraphs = extract_docx_structure(source_file)

    # Save to file
    with open(output_file, 'w', encoding='utf-8') as f:
        for i, para in enumerate(paragraphs):
            f.write(para + '\n')

    print(f"✓ Extracted {len(paragraphs)} paragraphs")
    print(f"✓ Saved to: {output_file}")

    # Print first 20 paragraphs as preview
    print("\n--- PREVIEW (First 20 paragraphs) ---")
    for i, para in enumerate(paragraphs[:20]):
        print(f"{i+1}: {para[:100]}..." if len(para) > 100 else f"{i+1}: {para}")
else:
    print(f"File not found: {source_file}")
