#!/usr/bin/env python3
"""
Word Document Decomposer and Reconstructor

This tool extracts the internal components of a .docx file (which is a ZIP archive
containing XML and other files), documents the structure in markdown, and can
reconstruct the original document from the extracted components.
"""

import zipfile
import os
import shutil
from pathlib import Path
from datetime import datetime
import xml.etree.ElementTree as ET


class DocxDecomposer:
    def __init__(self, docx_path):
        """
        Initialize the decomposer with a path to a .docx file.
        
        Args:
            docx_path: Path to the input .docx file
        """
        self.docx_path = Path(docx_path)
        self.extract_dir = None
        self.markdown_report = []
        
    def extract(self, output_dir=None):
        """
        Extract the .docx file to a directory.
        
        Args:
            output_dir: Directory to extract to. If None, creates a directory
                       based on the docx filename.
        
        Returns:
            Path to the extraction directory
        """
        if output_dir is None:
            base_name = self.docx_path.stem
            output_dir = Path(f"{base_name}_extracted")
        else:
            output_dir = Path(output_dir)
        
        # Remove existing directory if it exists
        if output_dir.exists():
            shutil.rmtree(output_dir)
        
        # Extract the ZIP archive
        print(f"Extracting {self.docx_path} to {output_dir}...")
        with zipfile.ZipFile(self.docx_path, 'r') as zip_ref:
            zip_ref.extractall(output_dir)
        
        self.extract_dir = output_dir
        print(f"Extraction complete: {len(list(output_dir.rglob('*')))} items extracted")
        return output_dir
    
    def analyze_structure(self):
        """
        Analyze the extracted directory structure and generate a markdown report.
        
        Returns:
            String containing the markdown report
        """
        if self.extract_dir is None:
            raise ValueError("Must call extract() before analyze_structure()")
        
        self.markdown_report = []
        
        # Header
        self._add_header()
        
        # Directory structure
        self._add_directory_tree()
        
        # Key file descriptions
        self._add_file_descriptions()
        
        # Content types
        self._add_content_types()
        
        # Relationships
        self._add_relationships()
        
        # Document structure
        self._add_document_structure()
        
        return "\n".join(self.markdown_report)
    
    def _add_header(self):
        """Add markdown header."""
        self.markdown_report.append(f"# Word Document Structure Analysis")
        self.markdown_report.append(f"\n**Source Document:** `{self.docx_path.name}`")
        self.markdown_report.append(f"**Analysis Date:** {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
        self.markdown_report.append(f"**Extraction Directory:** `{self.extract_dir}`")
        self.markdown_report.append("\n---\n")
    
    def _add_directory_tree(self):
        """Add directory tree structure."""
        self.markdown_report.append("## Directory Structure\n")
        self.markdown_report.append("```")
        self._print_tree(self.extract_dir, prefix="")
        self.markdown_report.append("```\n")
    
    def _print_tree(self, directory, prefix="", is_last=True):
        """Recursively print directory tree."""
        items = sorted(directory.iterdir(), key=lambda x: (not x.is_dir(), x.name))
        
        for i, item in enumerate(items):
            is_last_item = (i == len(items) - 1)
            current_prefix = "└── " if is_last_item else "├── "
            self.markdown_report.append(f"{prefix}{current_prefix}{item.name}")
            
            if item.is_dir():
                extension = "    " if is_last_item else "│   "
                self._print_tree(item, prefix + extension, is_last_item)
    
    def _add_file_descriptions(self):
        """Add descriptions of key files."""
        self.markdown_report.append("## Key Components\n")
        
        key_files = {
            "[Content_Types].xml": "Defines the content types for parts in the package",
            "_rels/.rels": "Package-level relationships (connections between parts)",
            "word/document.xml": "Main document content (paragraphs, tables, etc.)",
            "word/_rels/document.xml.rels": "Relationships for the main document",
            "word/styles.xml": "Style definitions (paragraph styles, character styles)",
            "word/settings.xml": "Document settings and properties",
            "word/fontTable.xml": "Font definitions used in the document",
            "word/numbering.xml": "Numbering definitions for lists",
            "word/theme/": "Theme files (colors, fonts, effects)",
            "docProps/core.xml": "Core document properties (author, created date, etc.)",
            "docProps/app.xml": "Application-specific properties"
        }
        
        for file_path, description in key_files.items():
            full_path = self.extract_dir / file_path
            exists = "✓" if full_path.exists() else "✗"
            self.markdown_report.append(f"- **{file_path}** {exists}")
            self.markdown_report.append(f"  - {description}")
            
            if full_path.exists() and full_path.is_file():
                size = full_path.stat().st_size
                self.markdown_report.append(f"  - Size: {size:,} bytes")
        
        self.markdown_report.append("")
    
    def _add_content_types(self):
        """Analyze and document content types."""
        content_types_path = self.extract_dir / "[Content_Types].xml"
        
        if not content_types_path.exists():
            return
        
        self.markdown_report.append("## Content Types\n")
        
        try:
            tree = ET.parse(content_types_path)
            root = tree.getroot()
            
            # Remove namespace for easier parsing
            for elem in root.iter():
                if '}' in elem.tag:
                    elem.tag = elem.tag.split('}', 1)[1]
            
            defaults = root.findall('.//Default')
            overrides = root.findall('.//Override')
            
            if defaults:
                self.markdown_report.append("### Default Content Types by Extension\n")
                for default in defaults:
                    ext = default.get('Extension')
                    content_type = default.get('ContentType')
                    self.markdown_report.append(f"- `.{ext}` → `{content_type}`")
                self.markdown_report.append("")
            
            if overrides:
                self.markdown_report.append("### Content Type Overrides by Part\n")
                for override in overrides:
                    part_name = override.get('PartName')
                    content_type = override.get('ContentType')
                    self.markdown_report.append(f"- `{part_name}` → `{content_type}`")
                self.markdown_report.append("")
        
        except Exception as e:
            self.markdown_report.append(f"Error parsing content types: {e}\n")
    
    def _add_relationships(self):
        """Analyze and document relationships."""
        rels_path = self.extract_dir / "_rels" / ".rels"
        
        if not rels_path.exists():
            return
        
        self.markdown_report.append("## Package Relationships\n")
        
        try:
            tree = ET.parse(rels_path)
            root = tree.getroot()
            
            # Remove namespace
            for elem in root.iter():
                if '}' in elem.tag:
                    elem.tag = elem.tag.split('}', 1)[1]
            
            relationships = root.findall('.//Relationship')
            
            if relationships:
                for rel in relationships:
                    rel_id = rel.get('Id')
                    rel_type = rel.get('Type', '').split('/')[-1]
                    target = rel.get('Target')
                    self.markdown_report.append(f"- **{rel_id}** ({rel_type})")
                    self.markdown_report.append(f"  - Target: `{target}`")
                self.markdown_report.append("")
        
        except Exception as e:
            self.markdown_report.append(f"Error parsing relationships: {e}\n")
    
    def _add_document_structure(self):
        """Analyze the main document structure."""
        doc_path = self.extract_dir / "word" / "document.xml"
        
        if not doc_path.exists():
            return
        
        self.markdown_report.append("## Document Content Structure\n")
        
        try:
            tree = ET.parse(doc_path)
            root = tree.getroot()
            
            # Count key elements
            namespaces = {'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
            
            paragraphs = root.findall('.//w:p', namespaces)
            tables = root.findall('.//w:tbl', namespaces)
            sections = root.findall('.//w:sectPr', namespaces)
            
            self.markdown_report.append(f"- **Paragraphs:** {len(paragraphs)}")
            self.markdown_report.append(f"- **Tables:** {len(tables)}")
            self.markdown_report.append(f"- **Sections:** {len(sections)}")
            
            # Sample first few paragraphs
            if paragraphs:
                self.markdown_report.append("\n### First 5 Paragraphs (text content):\n")
                for i, para in enumerate(paragraphs[:5], 1):
                    texts = para.findall('.//w:t', namespaces)
                    text_content = ''.join([t.text for t in texts if t.text])
                    if text_content.strip():
                        preview = text_content[:100] + "..." if len(text_content) > 100 else text_content
                        self.markdown_report.append(f"{i}. {preview}")
                self.markdown_report.append("")
        
        except Exception as e:
            self.markdown_report.append(f"Error parsing document structure: {e}\n")
    
    def save_analysis(self, output_path=None):
        """
        Save the markdown analysis to a file.
        
        Args:
            output_path: Path to save the markdown file. If None, uses default name.
        
        Returns:
            Path to the saved markdown file
        """
        if not self.markdown_report:
            self.analyze_structure()
        
        if output_path is None:
            output_path = self.extract_dir.parent / f"{self.extract_dir.name}_analysis.md"
        else:
            output_path = Path(output_path)
        
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write("\n".join(self.markdown_report))
        
        print(f"Analysis saved to: {output_path}")
        return output_path
    
    def reconstruct(self, output_path=None):
        """
        Reconstruct the .docx file from the extracted components.
        
        Args:
            output_path: Path for the reconstructed .docx file. If None, uses default name.
        
        Returns:
            Path to the reconstructed .docx file
        """
        if self.extract_dir is None:
            raise ValueError("Must call extract() before reconstruct()")
        
        if output_path is None:
            output_path = self.extract_dir.parent / f"{self.extract_dir.name}_reconstructed.docx"
        else:
            output_path = Path(output_path)
        
        print(f"Reconstructing document from {self.extract_dir}...")
        
        # Create a new ZIP file
        with zipfile.ZipFile(output_path, 'w', zipfile.ZIP_DEFLATED) as docx:
            # Walk through all files in the extracted directory
            for file_path in self.extract_dir.rglob('*'):
                if file_path.is_file():
                    # Get the relative path for the archive
                    arcname = file_path.relative_to(self.extract_dir)
                    docx.write(file_path, arcname)
        
        print(f"Reconstruction complete: {output_path}")
        return output_path


def main():
    """Main function demonstrating usage."""
    import sys
    
    if len(sys.argv) < 2:
        print("Usage: python docx_decomposer.py <path_to_docx>")
        print("\nExample: python docx_decomposer.py sample.docx")
        sys.exit(1)
    
    docx_path = sys.argv[1]
    
    if not os.path.exists(docx_path):
        print(f"Error: File not found: {docx_path}")
        sys.exit(1)
    
    # Create decomposer
    decomposer = DocxDecomposer(docx_path)
    
    # Extract the document
    extract_dir = decomposer.extract()
    
    # Analyze and save report
    analysis_path = decomposer.save_analysis()
    
    # Reconstruct the document
    reconstructed_path = decomposer.reconstruct()
    
    print("\n" + "="*60)
    print("SUMMARY")
    print("="*60)
    print(f"Original document:     {docx_path}")
    print(f"Extracted to:          {extract_dir}")
    print(f"Analysis report:       {analysis_path}")
    print(f"Reconstructed document: {reconstructed_path}")
    print("\nVerify the reconstructed document opens correctly in Word!")


if __name__ == "__main__":
    main()
