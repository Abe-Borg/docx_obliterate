# Word Document Format Analyzer

A Python tool that extracts the **formatting DNA** of Word documents (.docx files) for LLM consumption and specification template analysis.

## Overview

This tool analyzes .docx files and produces a focused report containing only formatting-critical information: styles, fonts, spacing, indentation, numbering, margins, headers/footers, and layout settings. 

**Perfect for:** Understanding how specification templates are formatted, comparing formatting between documents, and feeding formatting information to LLMs for document generation or analysis.

## Key Features

- **Focused Analysis**: Extracts only formatting-relevant information (styles, numbering, margins, etc.)
- **95% Reduction**: Produces 41KB reports instead of 887KB complete dumps
- **LLM-Ready Output**: Structured markdown format optimized for language model consumption
- **Specification-Focused**: Designed for MEP specification templates and architectural documents
- **No Dependencies**: Uses only Python standard library (zipfile, xml.etree.ElementTree)

## What It Analyzes

### 1. Styles (Most Important)
Every style in the document with complete formatting details:
- **Fonts**: Family, size, bold, italic, underline, color, highlighting
- **Paragraph Spacing**: Before, after, line spacing (in lines or exact)
- **Indentation**: Left, right, first-line, hanging (in inches and twips)
- **Alignment**: Left, right, center, justified
- **Numbering**: References to numbering definitions
- **Inheritance**: Based-on relationships between styles
- **Behavior**: Keep-with-next, page breaks, widow/orphan control, outline levels

### 2. Numbering Definitions
Complete numbering and bullet formatting:
- Abstract numbering patterns (templates)
- Level-by-level definitions (9 levels deep)
- Number formats: decimal (1, 2, 3), roman (i, ii, iii), letters (a, b, c), bullets
- Indentation and hanging indent per level
- Font specifications for numbers/bullets
- Numbering instances (mappings to abstract definitions)

### 3. Section Properties
Page layout and margins:
- Page size (width, height, orientation)
- Margins (top, bottom, left, right, header, footer, gutter)
- Column settings
- Header/footer references

### 4. Headers & Footers
If present in the document:
- Content structure
- Paragraph formatting
- Style references
- Alignment

### 5. Document Settings
Layout-affecting settings:
- Default tab stops
- Compatibility settings
- Character spacing control
- Even/odd header settings

## What It Ignores

The tool deliberately excludes information that doesn't affect formatting:
- ❌ Document content (actual paragraph text)
- ❌ Theme definitions (colors, effects)
- ❌ Relationships and content types
- ❌ Custom XML
- ❌ Binary file details
- ❌ Document properties (author, created date, etc.)
- ❌ Raw XML dumps

## Installation

No installation required. Just Python 3.6+.

```bash
# Clone or download the script
wget https://your-repo/docx_format_analyzer.py

# Or copy the script to your project
```

## Usage

### Basic Usage

```bash
python docx_format_analyzer.py your_document.docx
```

This creates: `your_document_format_analysis.md`

### In Your Code

```python
from docx_format_analyzer import DocxFormatAnalyzer

# Create analyzer
analyzer = DocxFormatAnalyzer("spec_template.docx")

# Analyze and get report as string
report = analyzer.analyze()

# Or save to file
output_path = analyzer.save_report("custom_output.md")
```

## Output Format

The tool produces a structured markdown report:

```markdown
# Document Formatting Analysis

**Source:** `template.docx`
**Date:** 2025-12-08 14:45:46

## Styles (164 total)

### Normal
- **ID:** `Normal`
- **Type:** paragraph
- **Default Style:** Yes

**Paragraph Formatting:**
- Space After: 120 twips (0.08 inches)
- Line Spacing: 1.15 lines

**Character Formatting:**
- Font (ASCII): Calibri
- Font Size: 11.0 pt

### Heading 1
- **ID:** `Heading1`
- **Type:** paragraph
- **Based On:** `Normal`

**Paragraph Formatting:**
- Space Before: 480 twips (0.33 inches)
- Left Indent: 0 twips (0.00 inches)
- Keep With Next: Yes
- Outline Level: 0

**Character Formatting:**
- Font Size: 16.0 pt
- Bold: Yes
- Color: #2F5496

[... continues with all styles, numbering, margins, etc.]
```

## Use Cases

### 1. Spec Template Analysis
Understand exactly how an architect's spec template is formatted:
```bash
python docx_format_analyzer.py architects_template.docx
```
Feed the output to an LLM to understand formatting conventions.

### 2. Template Comparison
Compare formatting between different templates:
```bash
python docx_format_analyzer.py template_A.docx
python docx_format_analyzer.py template_B.docx
# Use diff or LLM to compare the two reports
```

### 3. Format Matching
Extract formatting from a target document, then use it to format your own specs:
```bash
# Get target format
python docx_format_analyzer.py target_format.docx

# Feed the analysis to your spec generation pipeline
# Your code reads the analysis and applies the formatting
```

### 4. LLM-Powered Document Generation
Provide the format analysis to an LLM along with content:
```python
format_analysis = DocxFormatAnalyzer("template.docx").analyze()

prompt = f"""
Using this document format:
{format_analysis}

Generate a specification section for fire sprinklers...
"""
```

## Understanding the Output

### Measurement Units

The tool reports measurements in both **twips** and **inches**:
- **Twips**: Word's internal unit (1440 twips = 1 inch)
- **Inches**: Human-readable conversion

Example:
```
- Left Indent: 720 twips (0.50 inches)
```

### Style Inheritance

Styles can be based on other styles:
```
### Heading 2
- **Based On:** `Normal`
```
This means Heading 2 inherits all formatting from Normal, then applies its own overrides.

### Numbering System

Numbering has two parts:
1. **Abstract Numbering**: The template/pattern (e.g., "1., 2., 3.")
2. **Numbering Instance**: References an abstract numbering

Example:
```
- **Num ID 5** → Abstract Num 7
```
When a paragraph uses Num ID 5, it gets the format defined in Abstract Num 7.

## Technical Details

### How It Works

1. **Unzips** the .docx file (which is a ZIP archive)
2. **Parses XML** from these files:
   - `word/styles.xml` - Style definitions
   - `word/numbering.xml` - Numbering definitions  
   - `word/document.xml` - Section properties
   - `word/header*.xml` - Headers
   - `word/footer*.xml` - Footers
   - `word/settings.xml` - Document settings
3. **Extracts** only formatting-relevant elements
4. **Converts** measurements to readable units
5. **Generates** structured markdown report
6. **Cleans up** temporary files

### File Structure

```
your_project/
├── docx_format_analyzer.py    # Main script
├── your_document.docx          # Input
└── your_document_format_analysis.md  # Output
```

### Dependencies

**None.** Uses only Python standard library:
- `zipfile` - Extract .docx contents
- `xml.etree.ElementTree` - Parse XML
- `pathlib` - File path handling
- `datetime` - Timestamp for reports

## Comparison: Complete vs. Focused Analysis

| Metric | Complete Analyzer | Format Analyzer |
|--------|------------------|-----------------|
| Output Size | 887 KB | 41 KB |
| Line Count | 16,506 lines | 2,058 lines |
| Size Reduction | - | **95%** |
| Includes Content | Yes | No |
| Includes Themes | Yes | No |
| Includes Raw XML | Yes | No |
| LLM Token Usage | ~220,000 tokens | ~10,000 tokens |
| Focus | Everything | Formatting only |
| Use Case | Debugging | Template analysis |

## Example: Real-World MEP Spec

For a typical MEP specification template:

```bash
python docx_format_analyzer.py "Division 21 - Fire Sprinklers.docx"
```

Output includes:
- 50-80 custom styles (section headings, body text, lists, tables)
- 3-5 numbering definitions (section numbers, bullets, sub-items)
- Margins: typically 1" all around
- Headers: project info, section number
- Footers: page numbers, document control
- Font: usually Arial or Calibri, 11-12pt
- Line spacing: 1.0 or 1.15
- Indentation: 0.5" for sub-items, 0.25" hanging for lists

## Troubleshooting

### "No styles.xml found"
The .docx file may be corrupted. Try opening and re-saving in Word.

### "Error parsing XML"
The document may have custom XML that Word doesn't properly format. The tool will continue with other sections.

### Output seems incomplete
Some older .docx files (pre-2007) may not have all formatting stored in the expected locations. Re-save as a modern .docx.

## Performance

- **Speed**: Analyzes typical spec documents in 0.5-2 seconds
- **Memory**: < 50 MB for most documents
- **File Size**: Handles documents up to 100+ pages with ease

## Limitations

- Only works with .docx files (not .doc, .rtf, or other formats)
- Does not analyze:
  - Images or embedded objects
  - VBA macros
  - Track changes or comments
  - Custom XML beyond standard structure
  - Equation formatting
- Assumes standard Office Open XML structure

## Contributing

This tool is designed for MEP specification work in California. If you find issues or have suggestions:

1. Test your changes on real spec documents
2. Ensure output remains LLM-friendly
3. Keep dependencies minimal (stdlib only)
4. Document any new formatting elements extracted

## License

MIT License - Use freely for commercial or personal projects.

## Author

Built for MEP design workflow automation, specifically for California K-12, commercial, and healthcare projects.

## Related Tools

- **Complete Document Analyzer** (`docx_decomposer.py`): Extracts everything including content, themes, raw XML
- **Spec Template Matcher**: Uses this analyzer's output to reformat specifications

## Changelog

### v1.0 (2025-12-08)
- Initial release
- Extracts styles, numbering, sections, headers/footers, settings
- 95% size reduction compared to complete analysis
- Optimized for LLM consumption

---

**Questions?** This tool is part of a larger spec automation pipeline for MEP engineering work.
