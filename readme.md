# WordWriter

English | [简体中文](README_CN.md)

## Introduction

WordWriter is a Python library for processing Word document templates, making it easy to replace reserved tags in documents.

It supports replacing paragraph text, cell text, text boxes, headers and footers, and can also insert tables and images.

## ✨ What's New in v4.0

- **🎯 Object-Oriented API**: New `WordWriter` class with fluent interface
- **⛓️ Method Chaining**: Chain operations for cleaner code
- **🔍 Tag Inspection**: Get list of all tags with `get_tags()`
- **🛡️ Better Error Handling**: Clear exceptions and error messages
- **📦 Context Manager**: Support for `with` statement
- **🔄 Fully Backward Compatible**: All v3.x code still works!

## Requirements

```bash
pip install python-docx
pip install pandas
python3
```

## Quick Start (v4.0 OOP API) - Recommended

### Method 1: Fluent Interface (Recommended)

```python
from WordWriter import WordWriter

# The most concise way
WordWriter("template.docx") \
    .replace({
        "#[title]#": "My Report",
        "#[date]#": "2025-10-22"
    }) \
    .save("output.docx")
```

### Method 2: Step by Step

```python
from WordWriter import WordWriter

# Create instance
writer = WordWriter("template.docx")
writer.load()

# Inspect tags (v4.0 new feature)
tags = writer.get_tags()
print(f"Found {len(tags)} tags")

# Replace and save
writer.replace({"#[title]#": "My Report"})
writer.save("output.docx")
```

### Method 3: Context Manager

```python
from WordWriter import WordWriter

with WordWriter("template.docx") as writer:
    writer.replace({"#[title]#": "My Report"})
    writer.save("output.docx")
```

### Method 4: Class Method (One Step)

```python
from WordWriter import WordWriter

WordWriter.process("template.docx", "output.docx", 
                   {"#[title]#": "My Report"})
```

## Classic Usage (v3.x) - Still Supported

```python
from WordWriter import word_writer

resultsDict = {}
resultsDict["#[testheader1]#"] = "test header 1"
resultsDict["#[testheader2]#"] = "test header 2"
resultsDict["#[testString]#"] = "test strings"
resultsDict["#[testfooter]#"] = "test footer"
resultsDict["#[TX-testString2]#"] = "text box strings"
resultsDict["#[testTableString1]#"] = "cell text 1"
resultsDict["#[testTableString2]#"] = "cell text 2"
resultsDict["#[IMAGE-test1-(30,30)]#"] = "testPicture.png"
resultsDict["#[IMAGE-test2]#"] = "testPicture2.png"
resultsDict["#[IMAGE-test3-(10,10)]#"] = "testPicture.png"
resultsDict["#[TABLE-test1]#"] = "testTable.txt"

word_writer("test.docx", "output.docx", resultsDict)
```

## Tag Format

### Text Tags
```
#[tag_name]#
```
Used for replacing paragraph text, cell text, headers and footers, etc.

### Image Tags
```
#[IMAGE-image_name]#                    # Auto size
#[IMAGE-image_name-(width,height)]#     # Specified size (unit: cm)
#[TBIMG-image_name-(width,height)]#     # Image in table
```

### Table Tags
```
#[TABLE-table_name]#
```
Table data file should be a tab-separated text file (.txt).

### Text Box Tags
```
#[TX-textbox_name]#
```
Used for replacing content in text boxes.

## Special Values

- `#DELETETHISPARAGRAPH#` - Delete the paragraph containing the tag
- `#DELETETHISTABLE#` - Delete the table containing the tag

## Complete Example

```python
from WordWriter import WordWriter

# Prepare replacement data
replace_dict = {
    # Text replacement
    "#[title]#": "Annual Work Report",
    "#[author]#": "John Doe",
    "#[date]#": "October 22, 2025",
    
    # Headers and footers
    "#[header]#": "Confidential Document",
    "#[footer]#": "Page 1",
    
    # Image insertion
    "#[logo-(5,5)]#": "company_logo.png",
    "#[chart]#": "sales_chart.png",
    
    # Table insertion
    "#[TABLE-sales]#": "sales_data.txt",
    
    # Text box
    "#[TX-note]#": "Important: This document is for internal use only",
    
    # Delete paragraph
    "#[draft_watermark]#": "#DELETETHISPARAGRAPH#"
}

# Process using method chaining
WordWriter("template.docx") \
    .replace(replace_dict) \
    .save("annual_report.docx")

print("✓ Report generated successfully!")
```

## Advanced Usage

### Batch Processing Multiple Documents

```python
from WordWriter import WordWriter

templates = ["template1.docx", "template2.docx", "template3.docx"]
data_list = [data1, data2, data3]

for template, data in zip(templates, data_list):
    WordWriter(template) \
        .replace(data) \
        .save(f"output_{template}")
```

### Conditional Replacement

```python
writer = WordWriter("template.docx")
writer.load()

# Decide replacement content based on found tags
tags = writer.get_tags()
replace_dict = {}

if "#[date]#" in tags:
    from datetime import datetime
    replace_dict["#[date]#"] = datetime.now().strftime("%B %d, %Y")
    
if "#[title]#" in tags:
    replace_dict["#[title]#"] = "Auto-generated Report"

writer.replace(replace_dict).save("output.docx")
```

### Error Handling

```python
from WordWriter import WordWriter

try:
    writer = WordWriter("template.docx")
    writer.load()
    writer.replace(replace_dict)
    writer.save("output.docx")
    print("✓ Processing successful!")
except FileNotFoundError as e:
    print(f"✗ File not found: {e}")
except RuntimeError as e:
    print(f"✗ Runtime error: {e}")
except Exception as e:
    print(f"✗ Unknown error: {e}")
```

## Table Merging

WordWriter also provides table row merging functionality:

```python
from WordWriter import merge_table_row
from docx import Document

doc = Document("document.docx")
table = doc.tables[0]

# Merge rows with the same content in the first column
merge_table_row(table, 0)

doc.save("merged.docx")
```

## API Reference

### WordWriter Class

#### Constructor
```python
WordWriter(template_path: str)
```

#### Methods

- `load() -> WordWriter` - Load template (supports method chaining)
- `replace(replace_dict: Dict[str, str], logs: bool = True) -> WordWriter` - Replace tags (supports method chaining)
- `save(output_path: str) -> None` - Save document
- `get_tags() -> List[str]` - Get list of all tags
- `process(template_path, output_path, replace_dict, logs=True)` - Class method, one-step completion

#### Special Methods
- `__enter__` / `__exit__` - Context manager support
- `__repr__` - Object string representation

### Functional API (Backward Compatible)

```python
word_writer(input_docx: str, output_docx: str, 
            replace_dict: Dict[str, str], logs: bool = True) -> None
```

```python
merge_table_row(table: Table, col_index: int, 
                remove_other_row_text: bool = True) -> None
```

## Migration Guide

Upgrading from v3.x to v4.0? Check out the [Migration Guide](MIGRATION_GUIDE_v4.md).

## FAQ

### Q: What's the difference between v4.0 and v3.x?
**A**: v4.0 introduces an object-oriented API for a more modern usage pattern. However, it's fully backward compatible - all v3.x code works without modification.

### Q: Which API should I use?
**A**: For new projects, we recommend using the v4.0 OOP API (method chaining). Existing projects can continue using the functional API.

### Q: How do I check my version?
```python
import WordWriter
print(WordWriter.__version__)  # Output: 4.0.0
```

### Q: What if tags are not being replaced?
1. Check if the tag format is correct (`#[tag_name]#`)
2. Use `get_tags()` to see which tags were actually found
3. Ensure the tag actually exists in the template file

### Q: What image formats are supported?
All formats supported by python-docx: PNG, JPG, JPEG, GIF, BMP, TIFF, etc.

### Q: What is the table file format?
Tab-separated text file (.txt), one record per line, fields separated by Tab characters.

## License

MIT License
