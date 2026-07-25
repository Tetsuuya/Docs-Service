"""
DOCX Comparison Script
Compares content and design between two Word documents
"""

import zipfile
import xml.etree.ElementTree as ET
from pathlib import Path
import re

def extract_docx_xml(docx_path):
    """Extract XML content from DOCX file"""
    with zipfile.ZipFile(docx_path, 'r') as zip_ref:
        # Main document content
        document_xml = zip_ref.read('word/document.xml').decode('utf-8')
        
        # Try to get styles
        try:
            styles_xml = zip_ref.read('word/styles.xml').decode('utf-8')
        except:
            styles_xml = None
        
        # Try to get header
        try:
            header_xml = zip_ref.read('word/header1.xml').decode('utf-8')
        except:
            header_xml = None
        
        # Get list of media files (images)
        media_files = [f for f in zip_ref.namelist() if f.startswith('word/media/')]
        
    return document_xml, styles_xml, header_xml, media_files

def analyze_docx(docx_path):
    """Analyze DOCX structure and design"""
    print(f"\n{'='*80}")
    print(f"ANALYZING: {Path(docx_path).name}")
    print(f"{'='*80}\n")
    
    doc_xml, styles_xml, header_xml, media_files = extract_docx_xml(docx_path)
    
    # Parse XML
    root = ET.fromstring(doc_xml)
    ns = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
    
    # Analysis results
    analysis = {}
    
    # 1. PAGE COUNT
    page_breaks = doc_xml.count('<w:br w:type="page"/>')
    analysis['pages'] = page_breaks + 1
    
    # 2. PARAGRAPHS
    paragraphs = root.findall('.//w:p', ns)
    analysis['paragraphs'] = len(paragraphs)
    
    # 3. HEADINGS
    headings = []
    for p in paragraphs:
        pStyle = p.find('.//w:pStyle', ns)
        if pStyle is not None:
            style = pStyle.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val')
            if 'Heading' in style or 'Title' in style:
                # Get heading text
                texts = p.findall('.//w:t', ns)
                heading_text = ''.join([t.text for t in texts if t.text])
                headings.append(heading_text[:60])
    analysis['headings'] = headings
    
    # 4. TABLES - DETAILED ANALYSIS
    tables = root.findall('.//w:tbl', ns)
    analysis['tables'] = len(tables)
    
    # Extract table structures
    table_details = []
    for idx, table in enumerate(tables):
        rows = table.findall('.//w:tr', ns)
        table_info = {
            'index': idx + 1,
            'rows': len(rows),
            'cols': 0,
            'headers': [],
            'first_data_row': [],
            'has_shading': False
        }
        
        # Get column count from first row
        if rows:
            first_row = rows[0]
            cells = first_row.findall('.//w:tc', ns)
            table_info['cols'] = len(cells)
            
            # Extract header row text
            for cell in cells:
                texts = cell.findall('.//w:t', ns)
                cell_text = ''.join([t.text for t in texts if t.text]).strip()
                table_info['headers'].append(cell_text[:30])
            
            # Get first data row if exists
            if len(rows) > 1:
                second_row = rows[1]
                cells = second_row.findall('.//w:tc', ns)
                for cell in cells:
                    texts = cell.findall('.//w:t', ns)
                    cell_text = ''.join([t.text for t in texts if t.text]).strip()
                    table_info['first_data_row'].append(cell_text[:30])
        
        # Check for shading
        shading_elems = table.findall('.//w:shd', ns)
        table_info['has_shading'] = len(shading_elems) > 0
        
        table_details.append(table_info)
    
    analysis['table_details'] = table_details
    
    # 5. IMAGES
    analysis['images'] = len(media_files)
    
    # 6. BOLD TEXT
    bold_count = len(root.findall('.//w:b', ns))
    analysis['bold_instances'] = bold_count
    
    # 7. ITALIC TEXT
    italic_count = len(root.findall('.//w:i', ns))
    analysis['italic_instances'] = italic_count
    
    # 8. COLORS
    colors = set()
    for color_elem in root.findall('.//w:color', ns):
        color_val = color_elem.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val')
        if color_val and color_val != 'auto':
            colors.add(color_val)
    analysis['colors'] = list(colors)
    
    # 9. HEADER/FOOTER
    analysis['has_header'] = header_xml is not None
    if header_xml:
        analysis['header_text'] = ET.fromstring(header_xml).findall('.//w:t', ns)
        analysis['header_text'] = ''.join([t.text for t in analysis['header_text'] if t.text])
    
    # 10. FONT SIZES
    font_sizes = set()
    for sz_elem in root.findall('.//w:sz', ns):
        size = sz_elem.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}val')
        if size:
            font_sizes.add(int(size) // 2)  # Convert half-points to points
    analysis['font_sizes'] = sorted(list(font_sizes))
    
    # 11. TEXT CONTENT SAMPLE
    all_text = []
    for t in root.findall('.//w:t', ns):
        if t.text:
            all_text.append(t.text)
    full_text = ''.join(all_text)
    analysis['text_length'] = len(full_text)
    analysis['word_count'] = len(full_text.split())
    analysis['text_sample'] = full_text[:500]
    
    # 12. SHADING (Background colors)
    shading = set()
    for shd_elem in root.findall('.//w:shd', ns):
        fill = shd_elem.get('{http://schemas.openxmlformats.org/wordprocessingml/2006/main}fill')
        if fill and fill != 'auto':
            shading.add(fill)
    analysis['shading_colors'] = list(shading)
    
    return analysis

def print_analysis(analysis, doc_name):
    """Pretty print analysis results"""
    print(f"📄 DOCUMENT: {doc_name}")
    print(f"\n📊 STRUCTURE:")
    print(f"   Pages: {analysis['pages']}")
    print(f"   Paragraphs: {analysis['paragraphs']}")
    print(f"   Headings: {len(analysis['headings'])}")
    print(f"   Tables: {analysis['tables']}")
    print(f"   Images: {analysis['images']}")
    
    print(f"\n✍️  CONTENT:")
    print(f"   Text Length: {analysis['text_length']} characters")
    print(f"   Word Count: {analysis['word_count']} words")
    
    print(f"\n🎨 FORMATTING:")
    print(f"   Bold Instances: {analysis['bold_instances']}")
    print(f"   Italic Instances: {analysis['italic_instances']}")
    print(f"   Text Colors: {', '.join(analysis['colors']) if analysis['colors'] else 'None (default black)'}")
    print(f"   Shading Colors: {', '.join(analysis['shading_colors']) if analysis['shading_colors'] else 'None'}")
    print(f"   Font Sizes: {', '.join(map(str, analysis['font_sizes']))} pts")
    
    print(f"\n📋 DESIGN ELEMENTS:")
    print(f"   Has Header: {'✓ Yes' if analysis['has_header'] else '✗ No'}")
    if analysis['has_header']:
        print(f"   Header Text: {analysis.get('header_text', 'N/A')[:60]}")
    
    print(f"\n📝 HEADINGS FOUND:")
    for i, heading in enumerate(analysis['headings'][:10], 1):
        print(f"   {i}. {heading}")
    if len(analysis['headings']) > 10:
        print(f"   ... and {len(analysis['headings']) - 10} more")
    
    print(f"\n📊 TABLE DETAILS:")
    for table in analysis['table_details']:
        print(f"\n   Table {table['index']}: {table['rows']} rows × {table['cols']} cols")
        if table['headers']:
            print(f"      Headers: {' | '.join(table['headers'])}")
        if table['first_data_row']:
            print(f"      First row: {' | '.join(table['first_data_row'])}")
        print(f"      Has shading: {'Yes' if table['has_shading'] else 'No'}")
    
    print(f"\n📖 TEXT SAMPLE (first 500 chars):")
    print(f"   {analysis['text_sample'][:500]}...")
    print()

def compare_documents(doc1_analysis, doc2_analysis, doc1_name, doc2_name):
    """Compare two document analyses"""
    print(f"\n{'='*80}")
    print(f"COMPARISON: {doc1_name} vs {doc2_name}")
    print(f"{'='*80}\n")
    
    comparisons = [
        ('Pages', 'pages'),
        ('Paragraphs', 'paragraphs'),
        ('Headings', lambda a: len(a['headings'])),
        ('Tables', 'tables'),
        ('Images', 'images'),
        ('Bold Instances', 'bold_instances'),
        ('Italic Instances', 'italic_instances'),
        ('Word Count', 'word_count'),
        ('Has Header', 'has_header'),
        ('Font Sizes', lambda a: len(a['font_sizes'])),
        ('Colors Used', lambda a: len(a['colors'])),
    ]
    
    for label, key in comparisons:
        if callable(key):
            val1 = key(doc1_analysis)
            val2 = key(doc2_analysis)
        else:
            val1 = doc1_analysis[key]
            val2 = doc2_analysis[key]
        
        diff = ""
        if isinstance(val1, (int, float)) and isinstance(val2, (int, float)):
            if val1 > val2:
                diff = f"({doc1_name} has {val1 - val2} more)"
            elif val2 > val1:
                diff = f"({doc2_name} has {val2 - val1} more)"
            else:
                diff = "(Same)"
        elif val1 == val2:
            diff = "(Same)"
        else:
            diff = "(Different)"
        
        print(f"{label:20} | {str(val1):20} | {str(val2):20} | {diff}")
    
    print(f"\n🎯 KEY DIFFERENCES:")
    
    # Check for header difference
    if doc1_analysis['has_header'] != doc2_analysis['has_header']:
        if doc1_analysis['has_header']:
            print(f"   • {doc1_name} HAS headers/footers, {doc2_name} DOES NOT")
        else:
            print(f"   • {doc2_name} HAS headers/footers, {doc1_name} DOES NOT")
    
    # Check for color scheme
    if doc1_analysis['colors'] and not doc2_analysis['colors']:
        print(f"   • {doc1_name} uses custom colors, {doc2_name} uses default")
    elif doc2_analysis['colors'] and not doc1_analysis['colors']:
        print(f"   • {doc2_name} uses custom colors, {doc1_name} uses default")
    
    # Check for shading
    if doc1_analysis['shading_colors'] and not doc2_analysis['shading_colors']:
        print(f"   • {doc1_name} has background shading, {doc2_name} does not")
    elif doc2_analysis['shading_colors'] and not doc1_analysis['shading_colors']:
        print(f"   • {doc2_name} has background shading, {doc1_name} does not")
    
    # Structure similarity
    structure_score = 0
    if abs(doc1_analysis['pages'] - doc2_analysis['pages']) <= 2:
        structure_score += 1
    if abs(len(doc1_analysis['headings']) - len(doc2_analysis['headings'])) <= 5:
        structure_score += 1
    if doc1_analysis['tables'] > 0 and doc2_analysis['tables'] > 0:
        structure_score += 1
    
    print(f"\n📊 STRUCTURE SIMILARITY: {structure_score}/3")
    if structure_score >= 2:
        print("   ✓ Documents have similar structure")
    else:
        print("   ✗ Documents have different structures")
    
    # DETAILED TABLE COMPARISON
    print(f"\n📋 DETAILED TABLE COMPARISON:")
    print(f"   Template has {len(doc1_analysis['table_details'])} tables")
    print(f"   Output has {len(doc2_analysis['table_details'])} tables")
    
    if doc1_analysis['table_details'] and doc2_analysis['table_details']:
        print(f"\n   Comparing table structures:")
        max_tables = min(len(doc1_analysis['table_details']), len(doc2_analysis['table_details']))
        
        for i in range(max_tables):
            t1 = doc1_analysis['table_details'][i]
            t2 = doc2_analysis['table_details'][i]
            
            print(f"\n   TABLE {i+1}:")
            print(f"      Template: {t1['rows']} rows × {t1['cols']} cols")
            print(f"      Output:   {t2['rows']} rows × {t2['cols']} cols")
            
            # Compare headers
            if t1['headers'] and t2['headers']:
                print(f"      Template headers: {' | '.join(t1['headers'][:3])}")
                print(f"      Output headers:   {' | '.join(t2['headers'][:3])}")
                
                # Check if headers match
                if t1['headers'] == t2['headers']:
                    print(f"      ✓ Headers MATCH exactly!")
                else:
                    print(f"      ✗ Headers are DIFFERENT")
                    
            # Compare first data row
            if t1['first_data_row'] and t2['first_data_row']:
                print(f"      Template row 1: {' | '.join(t1['first_data_row'][:3])}")
                print(f"      Output row 1:   {' | '.join(t2['first_data_row'][:3])}")
                
                if t1['first_data_row'] == t2['first_data_row']:
                    print(f"      ✓ First row MATCHES exactly!")
                else:
                    print(f"      ✗ First row is DIFFERENT (AI generated new content)")

if __name__ == "__main__":
    # File paths
    doc1 = r"C:\Users\Rhenel Jhon Sajol\Downloads\Modele_Private_Equity.docx"
    doc2 = r"C:\Users\Rhenel Jhon Sajol\Downloads\global_private_equity_fund_iii.docx"
    
    # Check if files exist
    if not Path(doc1).exists():
        print(f"ERROR: File not found: {doc1}")
        exit(1)
    if not Path(doc2).exists():
        print(f"ERROR: File not found: {doc2}")
        exit(1)
    
    # Analyze both documents
    analysis1 = analyze_docx(doc1)
    print_analysis(analysis1, Path(doc1).stem)
    
    print("\n" + "="*80 + "\n")
    
    analysis2 = analyze_docx(doc2)
    print_analysis(analysis2, Path(doc2).stem)
    
    # Compare
    compare_documents(analysis1, analysis2, Path(doc1).stem, Path(doc2).stem)
    
    print(f"\n{'='*80}")
    print("ANALYSIS COMPLETE")
    print(f"{'='*80}\n")
