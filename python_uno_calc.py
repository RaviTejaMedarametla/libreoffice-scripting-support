"""
Python UNO Macro for LibreOffice Calc
Demonstrates document automation via UNO API

To use:
1. Save this as: Tools > Macros > Organize Macros > Python > New Macro
2. Run via: Tools > Macros > Run Macro
3. Select 'populate_calc' function
"""

import uno

def populate_calc():
    """
    Populate LibreOffice Calc with sample data using Python UNO.
    Creates a simple table with headers and data.
    """
    try:
        # Get the active document and sheet
        doc = XSCRIPTCONTEXT.getDocument()
        sheet = doc.CurrentController.ActiveSheet
        
        # Write headers (row 1)
        sheet.getCellRangeByName("A1").String = "Name"
        sheet.getCellRangeByName("B1").String = "Value"
        sheet.getCellRangeByName("C1").String = "Status"
        
        # Make headers bold (optional UNO property)
        header_range = sheet.getCellRangeByName("A1:C1")
        header_range.CharWeight = 150  # Bold
        
        # Write sample data (rows 2-4)
        data = [
            ["Alice", 42, "Active"],
            ["Bob", 87, "Inactive"],
            ["Charlie", 56, "Active"]
        ]
        
        for row_idx, row_data in enumerate(data, start=2):
            sheet.getCellRangeByName(f"A{row_idx}").String = row_data[0]
            sheet.getCellRangeByName(f"B{row_idx}").Value = row_data[1]
            sheet.getCellRangeByName(f"C{row_idx}").String = row_data[2]
        
        return "✓ Calc populated successfully! Check sheet for data."
    
    except Exception as e:
        return f"✗ Error: {str(e)}"

def clear_calc():
    """Clear the sheet (optional utility)."""
    doc = XSCRIPTCONTEXT.getDocument()
    sheet = doc.CurrentController.ActiveSheet
    sheet.getCellRangeByName("A1:C4").clearContents()
    return "Sheet cleared."
