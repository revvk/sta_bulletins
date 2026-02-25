#!/usr/bin/env python3
"""
Analyze .docx bulletin files - extract VML text box content and document structure.
"""

import zipfile
import xml.etree.ElementTree as ET
import sys
import re

def extract_document_xml(docx_path):
    """Extract word/document.xml from a .docx file."""
    with zipfile.ZipFile(docx_path, 'r') as z:
        # List all files in the zip
        print(f"Files in {docx_path}:")
        for name in z.namelist():
            if 'word/' in name:
                print(f"  {name}")
        print()

        with z.open('word/document.xml') as f:
            return f.read()

def get_all_namespaces(xml_bytes):
    """Extract all namespace declarations from the XML."""
    namespaces = {}
    # Parse namespace declarations
    xml_str = xml_bytes.decode('utf-8')
    ns_pattern = re.compile(r'xmlns:(\w+)="([^"]+)"')
    for match in ns_pattern.finditer(xml_str[:5000]):  # Check header area
        prefix, uri = match.groups()
        namespaces[prefix] = uri
    return namespaces

def extract_text_from_element(elem, namespaces):
    """Recursively extract all text from an element, preserving paragraph structure."""
    w_ns = namespaces.get('w', '')

    texts = []
    # Find all w:t elements
    for t in elem.iter(f'{{{w_ns}}}t'):
        if t.text:
            texts.append(t.text)
    return ''.join(texts)

def get_paragraph_properties(para, namespaces):
    """Extract paragraph properties like style, alignment, etc."""
    w_ns = namespaces.get('w', '')
    props = {}

    pPr = para.find(f'{{{w_ns}}}pPr')
    if pPr is not None:
        # Style
        pStyle = pPr.find(f'{{{w_ns}}}pStyle')
        if pStyle is not None:
            props['style'] = pStyle.get(f'{{{w_ns}}}val', '')

        # Alignment
        jc = pPr.find(f'{{{w_ns}}}jc')
        if jc is not None:
            props['alignment'] = jc.get(f'{{{w_ns}}}val', '')

        # Bold in run properties
        rPr = pPr.find(f'{{{w_ns}}}rPr')
        if rPr is not None:
            b = rPr.find(f'{{{w_ns}}}b')
            if b is not None:
                props['bold'] = True
            i = rPr.find(f'{{{w_ns}}}i')
            if i is not None:
                props['italic'] = True
            sz = rPr.find(f'{{{w_ns}}}sz')
            if sz is not None:
                props['fontSize'] = sz.get(f'{{{w_ns}}}val', '')
            color = rPr.find(f'{{{w_ns}}}color')
            if color is not None:
                props['color'] = color.get(f'{{{w_ns}}}val', '')

    return props

def get_run_properties(run, namespaces):
    """Extract run-level properties."""
    w_ns = namespaces.get('w', '')
    props = {}

    rPr = run.find(f'{{{w_ns}}}rPr')
    if rPr is not None:
        b = rPr.find(f'{{{w_ns}}}b')
        if b is not None:
            val = b.get(f'{{{w_ns}}}val', 'true')
            if val != '0' and val != 'false':
                props['bold'] = True

        i = rPr.find(f'{{{w_ns}}}i')
        if i is not None:
            val = i.get(f'{{{w_ns}}}val', 'true')
            if val != '0' and val != 'false':
                props['italic'] = True

        sz = rPr.find(f'{{{w_ns}}}sz')
        if sz is not None:
            props['fontSize'] = sz.get(f'{{{w_ns}}}val', '')

        color = rPr.find(f'{{{w_ns}}}color')
        if color is not None:
            props['color'] = color.get(f'{{{w_ns}}}val', '')

        u = rPr.find(f'{{{w_ns}}}u')
        if u is not None:
            props['underline'] = True

        caps = rPr.find(f'{{{w_ns}}}caps')
        if caps is not None:
            props['caps'] = True

        smallCaps = rPr.find(f'{{{w_ns}}}smallCaps')
        if smallCaps is not None:
            props['smallCaps'] = True

    return props

def analyze_paragraph_detail(para, namespaces):
    """Get detailed info about a paragraph including per-run formatting."""
    w_ns = namespaces.get('w', '')

    runs_info = []
    for run in para.findall(f'{{{w_ns}}}r'):
        run_props = get_run_properties(run, namespaces)
        run_text = ''
        for t in run.findall(f'{{{w_ns}}}t'):
            if t.text:
                run_text += t.text
        if run_text:
            runs_info.append({
                'text': run_text,
                'props': run_props
            })

    return runs_info

def check_for_page_break(para, namespaces):
    """Check if paragraph contains a page break."""
    w_ns = namespaces.get('w', '')

    # Check for page break in run
    for run in para.iter(f'{{{w_ns}}}r'):
        for br in run.findall(f'{{{w_ns}}}br'):
            br_type = br.get(f'{{{w_ns}}}type', '')
            if br_type == 'page':
                return True

    # Check for page break before in paragraph properties
    pPr = para.find(f'{{{w_ns}}}pPr')
    if pPr is not None:
        pageBreakBefore = pPr.find(f'{{{w_ns}}}pageBreakBefore')
        if pageBreakBefore is not None:
            return True
        # Also check sectPr for section breaks
        sectPr = pPr.find(f'{{{w_ns}}}sectPr')
        if sectPr is not None:
            return 'section_break'

    return False

def process_textbox_content(txbxContent, namespaces, textbox_index):
    """Process the content of a VML text box."""
    w_ns = namespaces.get('w', '')

    elements = []
    para_count = 0

    for para in txbxContent.findall(f'{{{w_ns}}}p'):
        para_count += 1
        text = extract_text_from_element(para, namespaces)
        props = get_paragraph_properties(para, namespaces)
        runs = analyze_paragraph_detail(para, namespaces)
        page_break = check_for_page_break(para, namespaces)

        if page_break:
            elements.append({
                'type': 'PAGE_BREAK' if page_break == True else 'SECTION_BREAK',
                'text': '',
                'props': {},
                'runs': []
            })

        element = {
            'type': 'paragraph',
            'text': text.strip() if text else '',
            'props': props,
            'runs': runs
        }
        elements.append(element)

    return elements, para_count

def classify_element(element):
    """Attempt to classify what role an element plays in the bulletin."""
    text = element.get('text', '')
    props = element.get('props', {})
    runs = element.get('runs', [])

    if not text:
        return 'empty_line'

    # Check run-level formatting
    has_bold = any(r.get('props', {}).get('bold') for r in runs)
    has_italic = any(r.get('props', {}).get('italic') for r in runs)
    has_caps = any(r.get('props', {}).get('caps') for r in runs)
    has_small_caps = any(r.get('props', {}).get('smallCaps') for r in runs)

    # Check for common patterns
    if text.startswith('Celebrant') or text.startswith('Priest') or text.startswith('Bishop'):
        return 'celebrant_line'
    if text.startswith('People'):
        return 'people_line'
    if text.startswith('Deacon'):
        return 'deacon_line'
    if text.startswith('Reader'):
        return 'reader_line'
    if text.startswith('Lector'):
        return 'lector_line'

    if props.get('italic') or has_italic:
        if has_bold:
            return 'bold_italic_text'
        return 'rubric/italic'

    if props.get('bold') or has_bold:
        if has_caps or text.isupper():
            return 'heading_caps'
        return 'heading/bold'

    if has_caps or text.isupper():
        return 'caps_text'

    return 'body_text'

def analyze_bulletin(docx_path):
    """Main analysis function for a bulletin .docx file."""
    print(f"\n{'='*100}")
    print(f"ANALYZING: {docx_path}")
    print(f"{'='*100}\n")

    xml_bytes = extract_document_xml(docx_path)
    namespaces = get_all_namespaces(xml_bytes)

    # Register namespaces for parsing
    for prefix, uri in namespaces.items():
        ET.register_namespace(prefix, uri)

    root = ET.fromstring(xml_bytes)

    w_ns = namespaces.get('w', '')
    wps_ns = namespaces.get('wps', '')
    mc_ns = namespaces.get('mc', '')
    wp_ns = namespaces.get('wp', '')

    print(f"Key namespaces:")
    for k in ['w', 'wps', 'mc', 'wp', 'v', 'r']:
        if k in namespaces:
            print(f"  {k}: {namespaces[k]}")
    print()

    # Find all text boxes - they can be in various locations
    # Method 1: Find wps:txbx elements
    textbox_count = 0
    all_textbox_elements = []

    # Search for txbxContent anywhere in the document
    txbx_contents = list(root.iter(f'{{{w_ns}}}txbxContent'))
    print(f"Found {len(txbx_contents)} w:txbxContent elements\n")

    for idx, txbxContent in enumerate(txbx_contents):
        textbox_count += 1
        elements, para_count = process_textbox_content(txbxContent, namespaces, idx)

        print(f"\n--- TEXT BOX #{idx + 1} ({para_count} paragraphs) ---")

        line_num = 0
        for elem in elements:
            if elem['type'] in ('PAGE_BREAK', 'SECTION_BREAK'):
                print(f"\n  *** {elem['type']} ***\n")
                continue

            line_num += 1
            text = elem['text']
            classification = classify_element(elem)
            props = elem['props']
            runs = elem['runs']

            # Build formatting description
            fmt_parts = []
            if props.get('style'):
                fmt_parts.append(f"style={props['style']}")
            if props.get('alignment'):
                fmt_parts.append(f"align={props['alignment']}")

            # Run-level formatting summary
            run_fmts = set()
            for r in runs:
                rp = r.get('props', {})
                if rp.get('bold'): run_fmts.add('BOLD')
                if rp.get('italic'): run_fmts.add('ITALIC')
                if rp.get('caps'): run_fmts.add('CAPS')
                if rp.get('smallCaps'): run_fmts.add('SMALLCAPS')
                if rp.get('underline'): run_fmts.add('UNDERLINE')
                if rp.get('color'): run_fmts.add(f"color={rp['color']}")
                if rp.get('fontSize'): run_fmts.add(f"sz={rp['fontSize']}")

            if run_fmts:
                fmt_parts.append('|'.join(sorted(run_fmts)))

            fmt_str = f" [{', '.join(fmt_parts)}]" if fmt_parts else ""

            # Truncate very long text for display but show enough
            display_text = text if len(text) <= 200 else text[:200] + '...'

            if text:
                print(f"  {line_num:3d}. [{classification:20s}]{fmt_str}")
                print(f"       \"{display_text}\"")
            else:
                print(f"  {line_num:3d}. [empty_line]")

        all_textbox_elements.append(elements)

    # Also check for content in the main document body (outside text boxes)
    body = root.find(f'{{{w_ns}}}body')
    if body is not None:
        print(f"\n\n--- MAIN DOCUMENT BODY (outside text boxes) ---")
        body_para_count = 0
        for para in body.findall(f'{{{w_ns}}}p'):
            text = extract_text_from_element(para, namespaces)
            if text and text.strip():
                body_para_count += 1
                props = get_paragraph_properties(para, namespaces)
                runs = analyze_paragraph_detail(para, namespaces)
                classification = classify_element({'text': text.strip(), 'props': props, 'runs': runs})
                print(f"  Body para {body_para_count}: [{classification}] \"{text.strip()[:150]}\"")
        if body_para_count == 0:
            print("  (No text content in main body outside text boxes)")

    print(f"\n\nTotal text boxes found: {textbox_count}")
    return all_textbox_elements

# Analyze both bulletins
pentecost = analyze_bulletin("2025-06-08 - Pentecost C - 9 am (HEII-A) - Bulletin.docx")
print("\n\n" + "="*100 + "\n" + "="*100 + "\n")
lent = analyze_bulletin("2025-03-23 - Lent 3C - 9 am (HEII-A) - Bulletin.docx")
