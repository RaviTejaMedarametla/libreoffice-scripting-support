# UNO (Universal Network Objects) Basics

## What is UNO?

**UNO** is LibreOffice's component model that allows different programming languages to interact with LibreOffice objects and documents. It's the bridge between your scripts and LibreOffice's core.

### Key Components

1. **XSCRIPTCONTEXT** (Python only)
   - Global object available in Python macros
   - Provides access to the active document
   - Example: `doc = XSCRIPTCONTEXT.getDocument()`

2. **ThisComponent** (Basic)
   - Equivalent to XSCRIPTCONTEXT in StarBasic
   - Example: `oDoc = ThisComponent`

3. **Services & Objects**
   - UNO provides objects like `Desktop`, `Document`, `Sheet`, `Cell`
   - Each object has properties and methods

## Common UNO Objects

### Document Objects

```python
# Get active document
doc = XSCRIPTCONTEXT.getDocument()

# Get current sheet (Calc)
sheet = doc.CurrentController.ActiveSheet

# Get current text (Writer)
text = doc.Text
```

### Cell Operations (Calc)

```python
# Get a single cell
cell = sheet.getCellRangeByName("A1")

# Set cell value
cell.Value = 42
cell.String = "Hello"

# Get cell value
value = cell.Value
```

### Text Manipulation (Writer)

```python
# Get text and cursor
text = doc.Text
cursor = text.createTextCursor()

# Insert text
text.insertString(cursor, "Hello World", False)

# Format properties
cursor.CharWeight = 150  # Bold
cursor.ParaStyleName = "Heading 1"
```

## UNO API Documentation

For complete reference, visit: **https://api.libreoffice.org/**

Common interfaces:
- `com.sun.star.sheet.XSpreadsheet` – Spreadsheet operations
- `com.sun.star.text.XText` – Text operations
- `com.sun.star.frame.XComponent` – Document handling

## Examples in This Repository

- `python_uno_calc.py` – Demonstrates cell population and formatting
- `python_uno_writer.py` – Shows text insertion and styling
- `basic_macro_sample.bas` – Basic equivalent using ThisComponent

## Quick Reference

| Task | Python | Basic |
|------|--------|-------|
| Get document | `XSCRIPTCONTEXT.getDocument()` | `ThisComponent` |
| Get sheet | `doc.CurrentController.ActiveSheet` | `ThisComponent.CurrentController.ActiveSheet` |
| Set cell | `sheet.getCellRangeByName("A1").Value = 42` | `oSheet.getCellRangeByName("A1").Value = 42` |
| Insert text | `text.insertString(cursor, "text", False)` | `oText.insertString(oCursor, "text", False)` |

---

**Next Steps:** Run the examples in this repo and modify them for your own documents!
