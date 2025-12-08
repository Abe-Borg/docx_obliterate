#!/usr/bin/env python3
"""
Word Document Editor - Simple Modifications

Provides two editing operations:
1. Change all fonts to Helvetica
2. Delete all content and replace with "EMPTY" (36pt bold Times New Roman)

Usage:
    python docx_editor.py input.docx --helvetica output.docx
    python docx_editor.py input.docx --empty output.docx
"""

import zipfile
import shutil
import stat
import os
from pathlib import Path
import xml.etree.ElementTree as ET
import argparse


class DocxEditor:
    def __init__(self, docx_path):
        """Initialize with input .docx file."""
        self.docx_path = Path(docx_path)
        self.extract_dir = None
        
    def extract(self):
        """Extract .docx to temporary directory."""
        self.extract_dir = Path(f".temp_{self.docx_path.stem}")
        
        # Remove if exists
        if self.extract_dir.exists():
            shutil.rmtree(self.extract_dir)
        
        # Extract
        with zipfile.ZipFile(self.docx_path, 'r') as zip_ref:
            zip_ref.extractall(self.extract_dir)
        
        print(f"Extracted to: {self.extract_dir}")
    
    def change_all_fonts_to_helvetica(self):
        """Change all fonts in styles.xml and document.xml to Helvetica."""
        if self.extract_dir is None:
            raise ValueError("Must extract first")
        
        print("Changing all fonts to Helvetica...")
        
        # Modify styles.xml
        styles_path = self.extract_dir / "word" / "styles.xml"
        if styles_path.exists():
            self._replace_fonts_in_file(styles_path, "Helvetica")
            print(f"  Modified: {styles_path.name}")
        
        # Modify document.xml (for any inline font specifications)
        doc_path = self.extract_dir / "word" / "document.xml"
        if doc_path.exists():
            self._replace_fonts_in_file(doc_path, "Helvetica")
            print(f"  Modified: {doc_path.name}")
        
        print("Font change complete!")
    
    def _replace_fonts_in_file(self, xml_path, new_font):
        """Replace all font specifications in an XML file."""
        # Read the file
        with open(xml_path, 'r', encoding='utf-8') as f:
            content = f.read()
        
        # Parse XML
        tree = ET.parse(xml_path)
        root = tree.getroot()
        
        # Find all rFonts elements (w:rFonts)
        # These specify fonts for runs
        namespaces = {
            'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
        }
        
        # Register namespace to preserve prefixes
        ET.register_namespace('w', namespaces['w'])
        ET.register_namespace('r', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships')
        ET.register_namespace('mc', 'http://schemas.openxmlformats.org/markup-compatibility/2006')
        ET.register_namespace('m', 'http://schemas.openxmlformats.org/officeDocument/2006/math')
        ET.register_namespace('v', 'urn:schemas-microsoft-com:vml')
        ET.register_namespace('wp', 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing')
        ET.register_namespace('w10', 'urn:schemas-microsoft-com:office:word')
        ET.register_namespace('o', 'urn:schemas-microsoft-com:office:office')
        
        # Find and modify all font references
        w_ns = namespaces['w']
        
        for rFonts in root.iter(f'{{{w_ns}}}rFonts'):
            # Set all font attributes to Helvetica
            if f'{{{w_ns}}}ascii' in rFonts.attrib:
                rFonts.attrib[f'{{{w_ns}}}ascii'] = new_font
            if f'{{{w_ns}}}hAnsi' in rFonts.attrib:
                rFonts.attrib[f'{{{w_ns}}}hAnsi'] = new_font
            if f'{{{w_ns}}}cs' in rFonts.attrib:
                rFonts.attrib[f'{{{w_ns}}}cs'] = new_font
            if f'{{{w_ns}}}eastAsia' in rFonts.attrib:
                rFonts.attrib[f'{{{w_ns}}}eastAsia'] = new_font
        
        # Write back
        tree.write(xml_path, encoding='utf-8', xml_declaration=True)
    
    def replace_with_empty(self):
        """Delete all content and replace with 'EMPTY' in 36pt bold Times New Roman."""
        if self.extract_dir is None:
            raise ValueError("Must extract first")
        
        print("Replacing all content with 'EMPTY'...")
        
        doc_path = self.extract_dir / "word" / "document.xml"
        
        if not doc_path.exists():
            print("  Error: document.xml not found")
            return
        
        # Parse the document
        tree = ET.parse(doc_path)
        root = tree.getroot()
        
        # Namespaces
        w_ns = 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'
        
        # Register namespaces
        ET.register_namespace('w', w_ns)
        ET.register_namespace('r', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships')
        ET.register_namespace('mc', 'http://schemas.openxmlformats.org/markup-compatibility/2006')
        ET.register_namespace('m', 'http://schemas.openxmlformats.org/officeDocument/2006/math')
        ET.register_namespace('v', 'urn:schemas-microsoft-com:vml')
        ET.register_namespace('wp', 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing')
        ET.register_namespace('w10', 'urn:schemas-microsoft-com:office:word')
        ET.register_namespace('o', 'urn:schemas-microsoft-com:office:office')
        ET.register_namespace('wp14', 'http://schemas.microsoft.com/office/word/2010/wordprocessingDrawing')
        ET.register_namespace('w14', 'http://schemas.microsoft.com/office/word/2010/wordml')
        
        # Find the body element
        body = root.find(f'.//{{{w_ns}}}body')
        
        if body is None:
            print("  Error: body element not found")
            return
        
        # Clear all content from body (but keep sectPr if it exists)
        sectPr = body.find(f'{{{w_ns}}}sectPr')
        
        # Remove everything
        body.clear()
        
        # Create a new paragraph with "EMPTY"
        # <w:p>
        p = ET.SubElement(body, f'{{{w_ns}}}p')
        
        #   <w:pPr> (paragraph properties)
        pPr = ET.SubElement(p, f'{{{w_ns}}}pPr')
        
        #   <w:r> (run)
        r = ET.SubElement(p, f'{{{w_ns}}}r')
        
        #     <w:rPr> (run properties)
        rPr = ET.SubElement(r, f'{{{w_ns}}}rPr')
        
        #       <w:rFonts> (font specification)
        rFonts = ET.SubElement(rPr, f'{{{w_ns}}}rFonts')
        rFonts.set(f'{{{w_ns}}}ascii', 'Times New Roman')
        rFonts.set(f'{{{w_ns}}}hAnsi', 'Times New Roman')
        rFonts.set(f'{{{w_ns}}}cs', 'Times New Roman')
        
        #       <w:b> (bold)
        b_elem = ET.SubElement(rPr, f'{{{w_ns}}}b')
        
        #       <w:sz> (font size in half-points, so 36pt = 72)
        sz = ET.SubElement(rPr, f'{{{w_ns}}}sz')
        sz.set(f'{{{w_ns}}}val', '72')
        
        #       <w:szCs> (complex script font size)
        szCs = ET.SubElement(rPr, f'{{{w_ns}}}szCs')
        szCs.set(f'{{{w_ns}}}val', '72')
        
        #     <w:t> (text content)
        t = ET.SubElement(r, f'{{{w_ns}}}t')
        t.text = 'EMPTY'
        
        # Restore sectPr if it existed
        if sectPr is not None:
            body.append(sectPr)
        
        # Write back
        tree.write(doc_path, encoding='utf-8', xml_declaration=True)
        
        print("  Content replaced with 'EMPTY' (36pt bold Times New Roman)")
    
    def reconstruct(self, output_path):
        """Reconstruct the .docx file from modified components."""
        if self.extract_dir is None:
            raise ValueError("Must extract first")
        
        output_path = Path(output_path)
        
        print(f"Reconstructing document to: {output_path}")
        
        # Create ZIP
        with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as docx:
            for file_path in self.extract_dir.rglob('*'):
                if file_path.is_file():
                    arcname = file_path.relative_to(self.extract_dir)
                    docx.write(file_path, arcname)
        
        print("Reconstruction complete!")
    
    def cleanup(self):
        """Delete the temporary extraction directory with Windows-compatible handling."""
        if self.extract_dir and self.extract_dir.exists():
            import time
            import stat
            
            def handle_remove_readonly(func, path, exc):
                """Error handler for Windows readonly files."""
                if func in (os.unlink, os.rmdir):
                    # Change the file to be writable and try again
                    os.chmod(path, stat.S_IWRITE)
                    func(path)
                else:
                    raise
            
            try:
                shutil.rmtree(self.extract_dir, onerror=handle_remove_readonly)
                print(f"Cleaned up: {self.extract_dir}")
            except PermissionError:
                # Windows file locking - wait briefly and retry
                print(f"Retrying cleanup (Windows file lock)...")
                time.sleep(0.5)
                try:
                    shutil.rmtree(self.extract_dir, onerror=handle_remove_readonly)
                    print(f"Cleaned up: {self.extract_dir}")
                except Exception as e:
                    print(f"Warning: Could not delete temp directory: {self.extract_dir}")
                    print(f"         You can manually delete it later. Error: {e}")
            except Exception as e:
                print(f"Warning: Could not delete temp directory: {self.extract_dir}")
                print(f"         You can manually delete it later. Error: {e}")


def main():
    """Main CLI interface."""
    parser = argparse.ArgumentParser(
        description='Edit Word documents: change fonts or replace content',
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # Change all fonts to Helvetica
  python docx_editor.py input.docx --helvetica output.docx
  
  # Replace all content with "EMPTY"
  python docx_editor.py input.docx --empty output.docx
        """
    )
    
    parser.add_argument('input', help='Input .docx file')
    parser.add_argument('output', help='Output .docx file')
    
    # Mutually exclusive group for operations
    group = parser.add_mutually_exclusive_group(required=True)
    group.add_argument('--helvetica', action='store_true',
                      help='Change all fonts to Helvetica')
    group.add_argument('--empty', action='store_true',
                      help='Replace all content with "EMPTY" (36pt bold Times New Roman)')
    
    args = parser.parse_args()
    
    # Validate input file
    if not Path(args.input).exists():
        print(f"Error: Input file not found: {args.input}")
        return 1
    
    # Create editor
    editor = DocxEditor(args.input)
    
    try:
        # Extract
        editor.extract()
        
        # Perform operation
        if args.helvetica:
            editor.change_all_fonts_to_helvetica()
        elif args.empty:
            editor.replace_with_empty()
        
        # Reconstruct
        editor.reconstruct(args.output)
        
        print(f"\n{'='*60}")
        print("SUCCESS")
        print(f"{'='*60}")
        print(f"Input:  {args.input}")
        print(f"Output: {args.output}")
        if args.helvetica:
            print("Operation: Changed all fonts to Helvetica")
        elif args.empty:
            print("Operation: Replaced content with 'EMPTY'")
        
    finally:
        # Always cleanup
        editor.cleanup()
    
    return 0


if __name__ == "__main__":
    exit(main())
