# Word Document Decomposer & Atomic Analyzer

A Python tool that **completely atomizes** Word documents (.docx files) down to every XML element, attribute, and byte. No stone left unturned.

## Overview

This tool tears apart .docx files and documents **every fucking thing**: every XML tag, every attribute, every namespace, every relationship, every binary file, plus complete raw XML dumps of all components.

**Perfect for:** Deep debugging, understanding OOXML structure, reverse engineering document formats, forensic analysis, and satisfying curiosity about what's really inside a Word document.

## Key Features

- **Complete Atomic Analysis**: Every XML element, every attribute, every relationship
- **Raw XML Dumps**: Complete unprocessed XML for every file
- **Binary Analysis**: Magic bytes and file signatures for images
- **Document Reconstruction**: Rebuilds the .docx from extracted components
- **Component Extraction**: Saves all parts (XML, images, etc.) to a directory
- **Exhaustive Documentation**: 887KB reports with 16,506 lines documenting everything

## What It Analyzes

### EVERYTHING. Literally everything:

#### 1. Complete File Inventory
Every single file with type and size:
```
- [Content_Types].xml: 1,738 bytes (1.70 KB)
- word/document.xml: 3,075 bytes (3.00 KB)
- word/styles.xml: 349,458 bytes (341.27 KB)
- docProps/thumbnail.jpeg: 8,324 bytes (8.13 KB)
```

#### 2. Content Types - Complete
Every extension mapping and part override:
```
- .jpeg → image/jpeg
- .rels → application/vnd.openxmlformats-package.relationships+xml
- /word/document.xml → application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml
```

#### 3. All Relationships - Every .rels File
Package relationships, document relationships, custom XML relationships:
```
rId3: core-properties → docProps/core.xml
rId1: officeDocument → word/document.xml
rId5: styles → word/styles.xml
```

#### 4. document.xml - Atomic Breakdown
- All namespaces (w, r, m, v, wp, o, etc.)
- Every paragraph with properties
- Every run with properties  
- Every text element with xml:space attributes
- Complete table structures (grid, rows, cells, properties)
- Section properties (margins, page size, columns)

Example depth:
```
Paragraph 1:
  Properties:
    - pStyle (val=Title)
  Runs: 1
    Run 1:
      Text: `Test Document`
```

#### 5. styles.xml - Complete (All 164+ Styles)
Every style with:
- Style ID, name, type
- Based-on relationships
- Next style
- UI priority
- Default/custom flags
- Complete paragraph properties
- Complete run (character) properties
- Font families for all scripts
- Sizes, colors, spacing, indentation

#### 6. Numbering - Complete
- All abstract numbering definitions (9 levels deep)
- All numbering instances
- Format codes, level text, alignment
- Indentation and hanging indents per level
- Font specifications per level

#### 7. Settings - Every Setting
Every document setting with attributes:
```
- defaultTabStop: 720 twips
- characterSpacingControl: doNotCompress
- compat settings
- view settings
- zoom settings
```

#### 8. Font Table - Every Font
All fonts with complete properties:
```
Font: Calibri
  - panose1: 020F0502020204030204
  - charset: 00
  - family: swiss
  - pitch: variable
```

#### 9. Theme - Complete Recursive Breakdown
Entire theme structure:
- Color schemes
- Font schemes  
- Format schemes
- Fill styles
- Line styles
- Effect styles
- Background fills

#### 10. Document Properties - Complete
Core properties:
- Creator, created date
- Last modified by, modified date
- Revision number
- Total edit time

App properties:
- Application name and version
- Document security
- Character/paragraph/line counts
- Page count

#### 11. Custom XML - If Present
Complete custom XML data stores with recursive element breakdown.

#### 12. Binary Files - Deep Analysis
For images and other binary files:
- File type and size
- Magic bytes (file signature)
- Hex dump of first 16 bytes

#### 13. RAW XML DUMPS
**Complete, unprocessed XML content for EVERY file.**
Every XML file in the .docx is dumped in full with formatting preserved.

Example:
```xml
<?xml version='1.0' encoding='UTF-8' standalone='yes'?>
<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
  <w:body>
    <w:p>
      <w:pPr>
        <w:pStyle w:val="Normal"/>
      </w:pPr>
      <w:r>
        <w:t>Text content here</w:t>
      </w:r>
    </w:p>
  </w:body>
</w:document>
```

## What It Does NOT Skip

**Nothing.** This tool documents:
- ✓ Content
- ✓ Formatting
- ✓ Themes
- ✓ Relationships
- ✓ Content types
- ✓ Custom XML
- ✓ Binary files
- ✓ Document properties
- ✓ Settings
- ✓ Web settings
- ✓ All namespaces
- ✓ All attributes
- ✓ Raw XML

If it's in the .docx file, it's in the report.

## Installation

No installation required. Just Python 3.6+.

```bash
# Download the script
wget https://your-repo/docx_decomposer.py

# Or copy it to your project
```

## Usage

### Basic Usage - Complete Analysis

```bash
python docx_decomposer.py your_document.docx
```

This creates:
- `your_document_extracted/` - Directory with all components
- `your_document_extracted_analysis.md` - Complete atomic analysis (16,506 lines)
- `your_document_extracted_reconstructed.docx` - Rebuilt document

### In Your Code

```python
from docx_decomposer import DocxDecomposer

# Create decomposer
decomposer = DocxDecomposer("document.docx")

# Extract to directory
extract_dir = decomposer.extract()

# Analyze and get complete report
report = decomposer.analyze_structure()

# Save analysis
analysis_path = decomposer.save_analysis()

# Reconstruct document from components
reconstructed = decomposer.reconstruct()
```

### Advanced Usage

```python
# Custom extraction directory
decomposer.extract(output_dir="custom_extract_dir")

# Custom analysis output
decomposer.save_analysis(output_path="my_analysis.md")

# Custom reconstruction path
decomposer.reconstruct(output_path="rebuilt.docx")
```

## Output Structure

### Extracted Directory

```
document_extracted/
├── [Content_Types].xml
├── _rels/
│   └── .rels
├── docProps/
│   ├── app.xml
│   ├── core.xml
│   └── thumbnail.jpeg
├── word/
│   ├── _rels/
│   │   └── document.xml.rels
│   ├── theme/
│   │   └── theme1.xml
│   ├── document.xml
│   ├── fontTable.xml
│   ├── numbering.xml
│   ├── settings.xml
│   ├── styles.xml
│   ├── stylesWithEffects.xml
│   └── webSettings.xml
└── customXml/
    ├── _rels/
    ├── item1.xml
    └── itemProps1.xml
```

### Analysis Report Structure

```markdown
# Word Document Structure Analysis

## Directory Structure
[Complete tree view of all files and folders]

## Complete File Inventory
[Every file with type and size]

## [Content_Types].xml - COMPLETE ANALYSIS
[All content types with namespaces]

## Relationships - COMPLETE ANALYSIS
[Every .rels file analyzed]

## word/document.xml - COMPLETE ATOMIC ANALYSIS
[Every paragraph, run, table, section]
- All namespaces
- Paragraph-by-paragraph breakdown
- Run-by-run breakdown
- Table-by-table with cells
- Section properties

## word/styles.xml - COMPLETE ANALYSIS
[All 164+ styles with complete properties]

## word/settings.xml - COMPLETE ANALYSIS
[Every setting documented]

## word/fontTable.xml - COMPLETE ANALYSIS
[Every font with all attributes]

## word/numbering.xml - COMPLETE ANALYSIS
[All numbering definitions]

## word/theme/ - COMPLETE ANALYSIS
[Complete theme structure]

## Document Properties - COMPLETE ANALYSIS
[Core and app properties]

## customXml/ - COMPLETE ANALYSIS
[If present]

## Binary Files Analysis
[Images with magic bytes]

## RAW XML DUMPS
[Complete XML for every file]
```

## Example Output Size

For a simple 5-paragraph document with 1 table:

```
Extracted directory: 24 files
Analysis report: 16,506 lines (887 KB)
  - Directory structure: 36 lines
  - File inventory: 68 lines
  - Content types: 50 lines
  - Relationships: 100 lines
  - document.xml analysis: 400 lines
  - styles.xml analysis: 8,000+ lines (164 styles)
  - Numbering: 800 lines
  - Settings: 200 lines
  - Font table: 300 lines
  - Theme: 500 lines
  - Properties: 100 lines
  - RAW XML dumps: 6,000+ lines
```

For a real MEP specification (50 pages):
- Extracted directory: 50-100 files
- Analysis report: 50,000+ lines (2-5 MB)

## Use Cases

### 1. Deep Debugging
Understanding why Word does something weird:
```bash
python docx_decomposer.py broken_document.docx
# Inspect extracted XML files or read analysis
```

### 2. Format Reverse Engineering
Figure out how a complex feature is implemented:
```bash
python docx_decomposer.py example_with_feature.docx
# Search analysis for relevant XML elements
```

### 3. OOXML Learning
Understand the Office Open XML standard:
```bash
python docx_decomposer.py simple_example.docx
# Study the raw XML dumps section
```

### 4. Document Forensics
Investigate document structure:
```bash
python docx_decomposer.py suspicious_document.docx
# Check relationships, custom XML, embedded objects
```

### 5. Automated Testing
Verify document reconstruction:
```bash
python docx_decomposer.py test_input.docx
# Compare original vs reconstructed document
```

### 6. Component Extraction
Get specific parts of a document:
```python
decomposer = DocxDecomposer("document.docx")
extract_dir = decomposer.extract()

# Now manually access specific files:
# extract_dir / "word" / "styles.xml"
# extract_dir / "word" / "document.xml"
```

### 7. Document Surgery
Modify specific XML and rebuild:
```python
decomposer = DocxDecomposer("document.docx")
extract_dir = decomposer.extract()

# Manually edit XML files in extract_dir
# For example: modify styles.xml

# Rebuild from modified components
decomposer.reconstruct("modified_document.docx")
```

## Technical Details

### How It Works

1. **Extract**: Unzips the .docx file (which is a ZIP archive)
2. **Walk**: Recursively walks all directories and files
3. **Parse**: Parses every XML file with namespace handling
4. **Analyze**: Documents every element, attribute, and structure
5. **Dump**: Includes complete raw XML for every file
6. **Reconstruct**: Re-zips everything back into a .docx

### OOXML Structure

A .docx file is actually:
```
.docx = ZIP archive containing:
  - XML files (document structure)
  - .rels files (relationships between parts)
  - Binary files (images, embedded objects)
  - [Content_Types].xml (part type definitions)
```

### File Format Details

**[Content_Types].xml**: Maps file extensions and parts to MIME types

**_rels/.rels**: Package-level relationships (main document, properties, thumbnail)

**word/document.xml**: Main document content (paragraphs, tables, runs, text)

**word/_rels/document.xml.rels**: Document relationships (styles, numbering, images, etc.)

**word/styles.xml**: All style definitions

**word/numbering.xml**: Numbering/bullet definitions

**word/settings.xml**: Document settings

**word/fontTable.xml**: Font definitions

**word/theme/theme1.xml**: Color/font/effect theme

**docProps/core.xml**: Core properties (author, dates, etc.)

**docProps/app.xml**: Application properties (page count, etc.)

### Dependencies

**None.** Uses only Python standard library:
- `zipfile` - ZIP archive handling
- `xml.etree.ElementTree` - XML parsing
- `pathlib` - Path operations
- `datetime` - Timestamps
- `os` - File operations
- `shutil` - Directory operations

### Namespace Handling

The tool properly handles XML namespaces:

```python
# Word's main namespace
w:document
w:p
w:r
w:t

# Relationship namespace
r:id
r:embed

# DrawingML namespace
a:theme
a:clrScheme
```

## Verification

The tool includes automatic verification:

```bash
# Content verification
Original paragraphs: 5
Reconstructed paragraphs: 5
Content match: ✓ PASS
```

Note: MD5 checksums may differ due to timestamps, but content is identical.

## Performance

- **Speed**: 0.5-3 seconds for typical documents
- **Memory**: < 100 MB for most documents
- **Disk Space**: Extracted directory ≈ original file size
- **Analysis**: Report is 20-30x larger than original file

## Limitations

- Only works with .docx files (Office Open XML format)
- Does not work with:
  - .doc files (old binary format)
  - .rtf files
  - .odt files
  - Password-protected documents
- Very large files (100+ MB) may be slow to analyze
- Analysis reports can be huge (multi-MB) for complex documents

## Comparison: Complete vs. Focused

| Feature | Complete Decomposer | Format Analyzer |
|---------|-------------------|-----------------|
| Output Size | 887 KB | 41 KB |
| Lines | 16,506 | 2,058 |
| Document Content | ✓ Yes | ✗ No |
| Formatting | ✓ Yes | ✓ Yes |
| Themes | ✓ Yes | ✗ No |
| Relationships | ✓ Yes | ✗ No |
| Raw XML | ✓ Yes | ✗ No |
| Binary Analysis | ✓ Yes | ✗ No |
| Component Extraction | ✓ Yes | ✗ No |
| Reconstruction | ✓ Yes | ✗ No |
| LLM Tokens | ~220,000 | ~10,000 |
| Use Case | Everything | Formatting only |

**Use Complete Decomposer when:** You need to know **everything**

**Use Format Analyzer when:** You only care about formatting

## Advanced Features

### 1. Component Extraction
Access extracted files programmatically:

```python
decomposer = DocxDecomposer("doc.docx")
extract_dir = decomposer.extract()

# Read specific XML
styles_path = extract_dir / "word" / "styles.xml"
with open(styles_path, 'r') as f:
    styles_xml = f.read()
```

### 2. Selective Analysis
Modify the code to analyze only specific parts:

```python
# In analyze_structure(), comment out sections you don't need
def analyze_structure(self):
    self._add_header()
    self._add_directory_tree()
    self._add_complete_file_inventory()
    # self._add_content_types_complete()  # Skip this
    # self._add_all_relationships()        # Skip this
    self._add_document_xml_complete()     # Keep this
    # ... etc
```

### 3. Batch Processing
Process multiple documents:

```python
from pathlib import Path

for docx_file in Path(".").glob("*.docx"):
    decomposer = DocxDecomposer(docx_file)
    decomposer.extract()
    decomposer.save_analysis()
```

### 4. Diff Analysis
Compare two documents:

```bash
python docx_decomposer.py doc_v1.docx
python docx_decomposer.py doc_v2.docx
diff doc_v1_extracted_analysis.md doc_v2_extracted_analysis.md
```

## Troubleshooting

### "Not a valid ZIP file"
The file is corrupted or not a real .docx file. Try opening in Word and saving again.

### "Permission denied"
The extracted directory already exists and is locked. Delete it manually or choose a different output directory.

### "Memory error"
The document is very large. Try closing other applications or use a machine with more RAM.

### "XML parsing error"
The document may have malformed XML. Word can sometimes save invalid XML that still opens in Word but fails strict parsing.

### Reconstruction checksum differs
This is normal. Timestamps and metadata may differ, but content is identical. Verify by opening both documents in Word.

## Security Considerations

- The tool extracts all embedded content (images, objects)
- Malicious documents may contain harmful embedded files
- Always decompose documents from trusted sources
- Extracted directories may contain sensitive metadata
- Raw XML dumps may reveal tracked changes, comments, hidden text

## Performance Tips

1. **Large files**: Use `extract()` only, skip `analyze_structure()` 
2. **Many files**: Process in parallel with multiprocessing
3. **Limited disk**: Delete extracted directories after analysis
4. **Memory**: Process files one at a time, not in batches

## Contributing

This tool is designed for complete document introspection. When contributing:

1. Maintain exhaustive analysis - don't skip anything
2. Keep raw XML dumps - they're essential for debugging
3. Document all namespaces and attributes
4. Test on complex real-world documents
5. Preserve reconstruction capability

## Known Issues

- Some compatibility settings are not fully documented
- Complex equations may not be fully analyzed
- VBA macros are not decompiled (by design)
- Some custom XML schemas are not recognized

## Future Enhancements

- [ ] HTML output format option
- [ ] JSON output format option
- [ ] Interactive explorer (web UI)
- [ ] Diff mode (compare two documents)
- [ ] Selective reconstruction (rebuild with modifications)
- [ ] Macro analysis (if present)
- [ ] OLE object extraction

## Version History

### v1.0 (2025-12-08)
- Initial release
- Complete atomic-level analysis
- Component extraction
- Document reconstruction
- Raw XML dumps
- Binary file analysis
- 16,506 line reports

## License

MIT License - Use freely for any purpose.

## Author

Built for MEP engineering work in California, specifically for understanding specification document structure at the atomic level.

## Related Tools

- **Format Analyzer** (`docx_format_analyzer.py`): Focused on formatting only, 95% smaller output
- **Spec Template Matcher**: Uses format analysis to reformat specifications

## Philosophy

> "I want to know fucking everything about a docx file. I want it all described in the report. It doesn't matter how long that report gets. I need to know the docx file down to the fucking atom."

This tool delivers on that promise.

---

**When you absolutely, positively need to know every goddamn thing about a Word document.**
