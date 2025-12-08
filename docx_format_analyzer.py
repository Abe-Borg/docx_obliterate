#!/usr/bin/env python3
"""
Word Document Format Analyzer

Extracts only the formatting-critical information from a .docx file:
- Styles (fonts, spacing, indentation, alignment)
- Numbering definitions
- Section properties (margins, page size)
- Headers/Footers
- Default settings affecting layout

This produces a focused report suitable for LLM understanding of document formatting.
"""

import zipfile
import os
from pathlib import Path
from datetime import datetime
import xml.etree.ElementTree as ET


class DocxFormatAnalyzer:
    def __init__(self, docx_path):
        """
        Initialize the analyzer with a path to a .docx file.
        
        Args:
            docx_path: Path to the input .docx file
        """
        self.docx_path = Path(docx_path)
        self.temp_dir = None
        self.report = []
        
    def analyze(self):
        """
        Extract and analyze formatting information from the .docx file.
        
        Returns:
            String containing the markdown report
        """
        # Extract to temp directory
        self.temp_dir = Path(f".temp_{self.docx_path.stem}")
        if self.temp_dir.exists():
            import shutil
            shutil.rmtree(self.temp_dir)
        
        with zipfile.ZipFile(self.docx_path, 'r') as zip_ref:
            zip_ref.extractall(self.temp_dir)
        
        self.report = []
        
        # Header
        self._add_header()
        
        # Styles - THE MOST IMPORTANT
        self._analyze_styles()
        
        # Numbering definitions
        self._analyze_numbering()
        
        # Section properties (margins, page size)
        self._analyze_sections()
        
        # Headers and footers
        self._analyze_headers_footers()
        
        # Document settings (tab stops, etc.)
        self._analyze_settings()
        
        # Cleanup
        import shutil
        shutil.rmtree(self.temp_dir)
        
        return "\n".join(self.report)
    
    def _add_header(self):
        """Add report header."""
        self.report.append(f"# Document Formatting Analysis")
        self.report.append(f"\n**Source:** `{self.docx_path.name}`")
        self.report.append(f"**Date:** {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        self.report.append("\nThis report contains only formatting-critical information for LLM understanding.")
        self.report.append("\n---\n")
    
    def _parse_xml(self, file_path):
        """Parse XML file and return tree with namespace handling."""
        if not file_path.exists():
            return None, None, {}
        
        tree = ET.parse(file_path)
        root = tree.getroot()
        
        # Extract namespaces
        namespaces = {}
        for event, elem in ET.iterparse(file_path, events=['start-ns']):
            prefix, uri = elem
            if prefix:
                namespaces[prefix] = uri
        
        return tree, root, namespaces
    
    def _get_val(self, element, ns_prefix='w'):
        """Get the w:val attribute from an element."""
        if element is None:
            return None
        ns = self._get_namespace(ns_prefix)
        return element.get(f'{{{ns}}}val') if ns else element.get('val')
    
    def _get_namespace(self, prefix):
        """Get namespace URI for a prefix."""
        common_ns = {
            'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
            'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
            'a': 'http://schemas.openxmlformats.org/drawingml/2006/main'
        }
        return common_ns.get(prefix)
    
    def _analyze_styles(self):
        """Analyze styles - the most critical formatting information."""
        styles_path = self.temp_dir / "word" / "styles.xml"
        tree, root, namespaces = self._parse_xml(styles_path)
        
        if root is None:
            self.report.append("## Styles\n\nNo styles.xml found.\n")
            return
        
        w_ns = self._get_namespace('w')
        ns = {'w': w_ns}
        
        styles = root.findall('.//w:style', ns)
        
        self.report.append(f"## Styles ({len(styles)} total)\n")
        self.report.append("Styles define the formatting DNA of the document.\n")
        
        for style in styles:
            style_id = style.get(f'{{{w_ns}}}styleId')
            style_type = style.get(f'{{{w_ns}}}type')
            is_default = style.get(f'{{{w_ns}}}default', '0') == '1'
            is_custom = style.get(f'{{{w_ns}}}customStyle', '0') == '1'
            
            # Get style name
            name_elem = style.find('w:name', ns)
            style_name = self._get_val(name_elem) if name_elem is not None else style_id
            
            self.report.append(f"### {style_name}")
            self.report.append(f"- **ID:** `{style_id}`")
            self.report.append(f"- **Type:** {style_type}")
            if is_default:
                self.report.append(f"- **Default Style:** Yes")
            if is_custom:
                self.report.append(f"- **Custom Style:** Yes")
            
            # Based on
            based_on = style.find('w:basedOn', ns)
            if based_on is not None:
                self.report.append(f"- **Based On:** `{self._get_val(based_on)}`")
            
            # Next style
            next_style = style.find('w:next', ns)
            if next_style is not None:
                self.report.append(f"- **Next Style:** `{self._get_val(next_style)}`")
            
            # UI Priority (helps understand importance)
            ui_priority = style.find('w:uiPriority', ns)
            if ui_priority is not None:
                self.report.append(f"- **UI Priority:** {self._get_val(ui_priority)}")
            
            # PARAGRAPH PROPERTIES
            pPr = style.find('w:pPr', ns)
            if pPr is not None:
                self.report.append("\n**Paragraph Formatting:**")
                
                # Alignment
                jc = pPr.find('w:jc', ns)
                if jc is not None:
                    self.report.append(f"- Alignment: {self._get_val(jc)}")
                
                # Spacing
                spacing = pPr.find('w:spacing', ns)
                if spacing is not None:
                    before = spacing.get(f'{{{w_ns}}}before')
                    after = spacing.get(f'{{{w_ns}}}after')
                    line = spacing.get(f'{{{w_ns}}}line')
                    line_rule = spacing.get(f'{{{w_ns}}}lineRule')
                    
                    if before:
                        self.report.append(f"- Space Before: {before} twips ({int(before)/1440:.2f} inches)")
                    if after:
                        self.report.append(f"- Space After: {after} twips ({int(after)/1440:.2f} inches)")
                    if line:
                        line_desc = f"{line} ({line_rule})" if line_rule else line
                        if line_rule == 'auto':
                            line_spacing = int(line) / 240
                            self.report.append(f"- Line Spacing: {line_spacing:.1f} lines")
                        else:
                            self.report.append(f"- Line Spacing: {line_desc}")
                
                # Indentation
                ind = pPr.find('w:ind', ns)
                if ind is not None:
                    left = ind.get(f'{{{w_ns}}}left')
                    right = ind.get(f'{{{w_ns}}}right')
                    first_line = ind.get(f'{{{w_ns}}}firstLine')
                    hanging = ind.get(f'{{{w_ns}}}hanging')
                    
                    if left:
                        self.report.append(f"- Left Indent: {left} twips ({int(left)/1440:.2f} inches)")
                    if right:
                        self.report.append(f"- Right Indent: {right} twips ({int(right)/1440:.2f} inches)")
                    if first_line:
                        self.report.append(f"- First Line Indent: {first_line} twips ({int(first_line)/1440:.2f} inches)")
                    if hanging:
                        self.report.append(f"- Hanging Indent: {hanging} twips ({int(hanging)/1440:.2f} inches)")
                
                # Numbering reference
                numPr = pPr.find('w:numPr', ns)
                if numPr is not None:
                    numId = numPr.find('w:numId', ns)
                    ilvl = numPr.find('w:ilvl', ns)
                    if numId is not None:
                        num_id_val = self._get_val(numId)
                        ilvl_val = self._get_val(ilvl) if ilvl is not None else '0'
                        self.report.append(f"- Numbering: ID={num_id_val}, Level={ilvl_val}")
                
                # Keep with next
                keep_next = pPr.find('w:keepNext', ns)
                if keep_next is not None:
                    self.report.append(f"- Keep With Next: Yes")
                
                # Keep lines together
                keep_lines = pPr.find('w:keepLines', ns)
                if keep_lines is not None:
                    self.report.append(f"- Keep Lines Together: Yes")
                
                # Page break before
                page_break = pPr.find('w:pageBreakBefore', ns)
                if page_break is not None:
                    self.report.append(f"- Page Break Before: Yes")
                
                # Widow/orphan control
                widow_control = pPr.find('w:widowControl', ns)
                if widow_control is not None:
                    self.report.append(f"- Widow/Orphan Control: Yes")
                
                # Outline level
                outline_lvl = pPr.find('w:outlineLvl', ns)
                if outline_lvl is not None:
                    self.report.append(f"- Outline Level: {self._get_val(outline_lvl)}")
            
            # CHARACTER (RUN) PROPERTIES
            rPr = style.find('w:rPr', ns)
            if rPr is not None:
                self.report.append("\n**Character Formatting:**")
                
                # Font
                rFonts = rPr.find('w:rFonts', ns)
                if rFonts is not None:
                    ascii_font = rFonts.get(f'{{{w_ns}}}ascii')
                    hAnsi = rFonts.get(f'{{{w_ns}}}hAnsi')
                    cs = rFonts.get(f'{{{w_ns}}}cs')
                    
                    if ascii_font:
                        self.report.append(f"- Font (ASCII): {ascii_font}")
                    if hAnsi and hAnsi != ascii_font:
                        self.report.append(f"- Font (High ANSI): {hAnsi}")
                    if cs:
                        self.report.append(f"- Font (Complex Script): {cs}")
                
                # Font size
                sz = rPr.find('w:sz', ns)
                if sz is not None:
                    size = int(self._get_val(sz)) / 2  # Half-points to points
                    self.report.append(f"- Font Size: {size:.1f} pt")
                
                # Bold
                b = rPr.find('w:b', ns)
                if b is not None:
                    self.report.append(f"- Bold: Yes")
                
                # Italic
                i = rPr.find('w:i', ns)
                if i is not None:
                    self.report.append(f"- Italic: Yes")
                
                # Underline
                u = rPr.find('w:u', ns)
                if u is not None:
                    u_val = self._get_val(u)
                    self.report.append(f"- Underline: {u_val}")
                
                # Strike
                strike = rPr.find('w:strike', ns)
                if strike is not None:
                    self.report.append(f"- Strikethrough: Yes")
                
                # Small caps
                smallCaps = rPr.find('w:smallCaps', ns)
                if smallCaps is not None:
                    self.report.append(f"- Small Caps: Yes")
                
                # All caps
                caps = rPr.find('w:caps', ns)
                if caps is not None:
                    self.report.append(f"- All Caps: Yes")
                
                # Color
                color = rPr.find('w:color', ns)
                if color is not None:
                    color_val = self._get_val(color)
                    if color_val and color_val != 'auto':
                        self.report.append(f"- Color: #{color_val}")
                
                # Highlight
                highlight = rPr.find('w:highlight', ns)
                if highlight is not None:
                    self.report.append(f"- Highlight: {self._get_val(highlight)}")
                
                # Character spacing
                spacing = rPr.find('w:spacing', ns)
                if spacing is not None:
                    self.report.append(f"- Character Spacing: {self._get_val(spacing)} twips")
            
            self.report.append("")
    
    def _analyze_numbering(self):
        """Analyze numbering definitions."""
        numbering_path = self.temp_dir / "word" / "numbering.xml"
        tree, root, namespaces = self._parse_xml(numbering_path)
        
        if root is None:
            self.report.append("## Numbering\n\nNo numbering.xml found.\n")
            return
        
        w_ns = self._get_namespace('w')
        ns = {'w': w_ns}
        
        abstract_nums = root.findall('.//w:abstractNum', ns)
        num_defs = root.findall('.//w:num', ns)
        
        self.report.append(f"## Numbering Definitions\n")
        
        # Abstract numbering (the templates)
        self.report.append(f"### Abstract Numbering ({len(abstract_nums)} definitions)\n")
        
        for abs_num in abstract_nums:
            abs_num_id = abs_num.get(f'{{{w_ns}}}abstractNumId')
            
            self.report.append(f"#### Abstract Num ID: {abs_num_id}\n")
            
            # Get levels
            levels = abs_num.findall('w:lvl', ns)
            
            for level in levels:
                ilvl = level.get(f'{{{w_ns}}}ilvl')
                
                self.report.append(f"**Level {ilvl}:**")
                
                # Start value
                start = level.find('w:start', ns)
                if start is not None:
                    self.report.append(f"- Start: {self._get_val(start)}")
                
                # Number format
                numFmt = level.find('w:numFmt', ns)
                if numFmt is not None:
                    self.report.append(f"- Format: {self._get_val(numFmt)}")
                
                # Level text (the actual numbering pattern)
                lvlText = level.find('w:lvlText', ns)
                if lvlText is not None:
                    self.report.append(f"- Text: `{self._get_val(lvlText)}`")
                
                # Alignment
                lvlJc = level.find('w:lvlJc', ns)
                if lvlJc is not None:
                    self.report.append(f"- Alignment: {self._get_val(lvlJc)}")
                
                # Paragraph properties for this level
                pPr = level.find('w:pPr', ns)
                if pPr is not None:
                    ind = pPr.find('w:ind', ns)
                    if ind is not None:
                        left = ind.get(f'{{{w_ns}}}left')
                        hanging = ind.get(f'{{{w_ns}}}hanging')
                        if left:
                            self.report.append(f"- Left Indent: {left} twips ({int(left)/1440:.2f} inches)")
                        if hanging:
                            self.report.append(f"- Hanging Indent: {hanging} twips ({int(hanging)/1440:.2f} inches)")
                
                # Run properties (font for numbers)
                rPr = level.find('w:rPr', ns)
                if rPr is not None:
                    rFonts = rPr.find('w:rFonts', ns)
                    if rFonts is not None:
                        font = rFonts.get(f'{{{w_ns}}}ascii')
                        if font:
                            self.report.append(f"- Font: {font}")
                
                self.report.append("")
        
        # Numbering instances (references to abstract nums)
        self.report.append(f"### Numbering Instances ({len(num_defs)} instances)\n")
        
        for num in num_defs:
            num_id = num.get(f'{{{w_ns}}}numId')
            
            abstract_num_id = num.find('w:abstractNumId', ns)
            if abstract_num_id is not None:
                self.report.append(f"- **Num ID {num_id}** → Abstract Num {self._get_val(abstract_num_id)}")
        
        self.report.append("")
    
    def _analyze_sections(self):
        """Analyze section properties (margins, page size)."""
        doc_path = self.temp_dir / "word" / "document.xml"
        tree, root, namespaces = self._parse_xml(doc_path)
        
        if root is None:
            self.report.append("## Section Properties\n\nNo document.xml found.\n")
            return
        
        w_ns = self._get_namespace('w')
        ns = {'w': w_ns}
        
        sections = root.findall('.//w:sectPr', ns)
        
        self.report.append(f"## Section Properties ({len(sections)} sections)\n")
        
        for i, section in enumerate(sections, 1):
            self.report.append(f"### Section {i}\n")
            
            # Page size
            pgSz = section.find('w:pgSz', ns)
            if pgSz is not None:
                width = pgSz.get(f'{{{w_ns}}}w')
                height = pgSz.get(f'{{{w_ns}}}h')
                orient = pgSz.get(f'{{{w_ns}}}orient', 'portrait')
                
                if width and height:
                    w_inches = int(width) / 1440
                    h_inches = int(height) / 1440
                    self.report.append(f"**Page Size:**")
                    self.report.append(f"- Width: {width} twips ({w_inches:.2f} inches)")
                    self.report.append(f"- Height: {height} twips ({h_inches:.2f} inches)")
                    self.report.append(f"- Orientation: {orient}")
                    self.report.append("")
            
            # Margins
            pgMar = section.find('w:pgMar', ns)
            if pgMar is not None:
                self.report.append(f"**Margins:**")
                
                margins = {
                    'top': 'Top',
                    'bottom': 'Bottom',
                    'left': 'Left',
                    'right': 'Right',
                    'header': 'Header',
                    'footer': 'Footer',
                    'gutter': 'Gutter'
                }
                
                for attr, label in margins.items():
                    value = pgMar.get(f'{{{w_ns}}}{attr}')
                    if value:
                        inches = int(value) / 1440
                        self.report.append(f"- {label}: {value} twips ({inches:.2f} inches)")
                
                self.report.append("")
            
            # Columns
            cols = section.find('w:cols', ns)
            if cols is not None:
                num = cols.get(f'{{{w_ns}}}num', '1')
                space = cols.get(f'{{{w_ns}}}space')
                self.report.append(f"**Columns:** {num}")
                if space:
                    self.report.append(f"- Column Spacing: {space} twips ({int(space)/1440:.2f} inches)")
                self.report.append("")
            
            # Header references
            headerReference = section.findall('w:headerReference', ns)
            if headerReference:
                self.report.append(f"**Headers:** {len(headerReference)}")
                for hdr in headerReference:
                    hdr_type = hdr.get(f'{{{w_ns}}}type')
                    hdr_id = hdr.get(f'{{{self._get_namespace("r")}}}id')
                    self.report.append(f"- Type: {hdr_type}, ID: {hdr_id}")
                self.report.append("")
            
            # Footer references
            footerReference = section.findall('w:footerReference', ns)
            if footerReference:
                self.report.append(f"**Footers:** {len(footerReference)}")
                for ftr in footerReference:
                    ftr_type = ftr.get(f'{{{w_ns}}}type')
                    ftr_id = ftr.get(f'{{{self._get_namespace("r")}}}id')
                    self.report.append(f"- Type: {ftr_type}, ID: {ftr_id}")
                self.report.append("")
    
    def _analyze_headers_footers(self):
        """Analyze headers and footers."""
        word_dir = self.temp_dir / "word"
        
        # Find all header and footer files
        headers = list(word_dir.glob('header*.xml'))
        footers = list(word_dir.glob('footer*.xml'))
        
        if not headers and not footers:
            self.report.append("## Headers & Footers\n\nNo headers or footers found.\n")
            return
        
        self.report.append(f"## Headers & Footers\n")
        
        w_ns = self._get_namespace('w')
        ns = {'w': w_ns}
        
        # Analyze headers
        if headers:
            self.report.append(f"### Headers ({len(headers)} files)\n")
            
            for header_file in sorted(headers):
                tree, root, _ = self._parse_xml(header_file)
                
                if root is None:
                    continue
                
                self.report.append(f"#### {header_file.name}\n")
                
                # Get all paragraphs
                paragraphs = root.findall('.//w:p', ns)
                
                self.report.append(f"**Paragraphs:** {len(paragraphs)}")
                
                for j, para in enumerate(paragraphs, 1):
                    # Get text content
                    texts = para.findall('.//w:t', ns)
                    text_content = ''.join([t.text for t in texts if t.text])
                    
                    if text_content.strip():
                        self.report.append(f"{j}. Content: `{text_content[:100]}`")
                    
                    # Paragraph properties
                    pPr = para.find('w:pPr', ns)
                    if pPr is not None:
                        # Alignment
                        jc = pPr.find('w:jc', ns)
                        if jc is not None:
                            self.report.append(f"   - Alignment: {self._get_val(jc)}")
                        
                        # Style reference
                        pStyle = pPr.find('w:pStyle', ns)
                        if pStyle is not None:
                            self.report.append(f"   - Style: {self._get_val(pStyle)}")
                
                self.report.append("")
        
        # Analyze footers
        if footers:
            self.report.append(f"### Footers ({len(footers)} files)\n")
            
            for footer_file in sorted(footers):
                tree, root, _ = self._parse_xml(footer_file)
                
                if root is None:
                    continue
                
                self.report.append(f"#### {footer_file.name}\n")
                
                # Get all paragraphs
                paragraphs = root.findall('.//w:p', ns)
                
                self.report.append(f"**Paragraphs:** {len(paragraphs)}")
                
                for j, para in enumerate(paragraphs, 1):
                    # Get text content
                    texts = para.findall('.//w:t', ns)
                    text_content = ''.join([t.text for t in texts if t.text])
                    
                    if text_content.strip():
                        self.report.append(f"{j}. Content: `{text_content[:100]}`")
                    
                    # Paragraph properties
                    pPr = para.find('w:pPr', ns)
                    if pPr is not None:
                        # Alignment
                        jc = pPr.find('w:jc', ns)
                        if jc is not None:
                            self.report.append(f"   - Alignment: {self._get_val(jc)}")
                        
                        # Style reference
                        pStyle = pPr.find('w:pStyle', ns)
                        if pStyle is not None:
                            self.report.append(f"   - Style: {self._get_val(pStyle)}")
                
                self.report.append("")
    
    def _analyze_settings(self):
        """Analyze document settings that affect layout."""
        settings_path = self.temp_dir / "word" / "settings.xml"
        tree, root, namespaces = self._parse_xml(settings_path)
        
        if root is None:
            self.report.append("## Document Settings\n\nNo settings.xml found.\n")
            return
        
        w_ns = self._get_namespace('w')
        ns = {'w': w_ns}
        
        self.report.append(f"## Document Settings (Layout-Affecting)\n")
        
        # Default tab stop
        defaultTabStop = root.find('.//w:defaultTabStop', ns)
        if defaultTabStop is not None:
            val = self._get_val(defaultTabStop)
            if val:
                inches = int(val) / 1440
                self.report.append(f"**Default Tab Stop:** {val} twips ({inches:.2f} inches)")
        
        # Compatibility settings (can affect layout)
        compat = root.find('.//w:compat', ns)
        if compat is not None:
            self.report.append(f"\n**Compatibility Settings:**")
            for setting in compat:
                tag_name = setting.tag.split('}')[-1] if '}' in setting.tag else setting.tag
                val = self._get_val(setting)
                if val:
                    self.report.append(f"- {tag_name}: {val}")
                else:
                    self.report.append(f"- {tag_name}")
        
        # Character spacing control
        characterSpacingControl = root.find('.//w:characterSpacingControl', ns)
        if characterSpacingControl is not None:
            self.report.append(f"\n**Character Spacing Control:** {self._get_val(characterSpacingControl)}")
        
        # Even and odd headers
        evenAndOddHeaders = root.find('.//w:evenAndOddHeaders', ns)
        if evenAndOddHeaders is not None:
            self.report.append(f"\n**Even and Odd Headers:** Yes")
        
        self.report.append("")
    
    def save_report(self, output_path=None):
        """
        Save the formatting analysis to a file.
        
        Args:
            output_path: Path to save the markdown file. If None, uses default name.
        
        Returns:
            Path to the saved markdown file
        """
        if output_path is None:
            output_path = self.docx_path.parent / f"{self.docx_path.stem}_format_analysis.md"
        else:
            output_path = Path(output_path)
        
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write("\n".join(self.report))
        
        return output_path


def main():
    """Main function demonstrating usage."""
    import sys
    
    if len(sys.argv) < 2:
        print("Usage: python docx_format_analyzer.py <path_to_docx>")
        print("\nExample: python docx_format_analyzer.py sample.docx")
        sys.exit(1)
    
    docx_path = sys.argv[1]
    
    if not os.path.exists(docx_path):
        print(f"Error: File not found: {docx_path}")
        sys.exit(1)
    
    # Create analyzer
    analyzer = DocxFormatAnalyzer(docx_path)
    
    # Analyze formatting
    print(f"Analyzing formatting in {docx_path}...")
    report = analyzer.analyze()
    
    # Save report
    output_path = analyzer.save_report()
    
    print(f"\n{'='*60}")
    print("ANALYSIS COMPLETE")
    print(f"{'='*60}")
    print(f"Report saved to: {output_path}")
    print(f"Report size: {len(report):,} characters")
    print("\nThis focused report contains only formatting-critical information")
    print("suitable for LLM understanding of document structure.")


if __name__ == "__main__":
    main()
