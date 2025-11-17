# Setup Guide for LibreOffice Scripting Examples

## Prerequisites

- **LibreOffice 25.8+** or latest version
- **Operating System:** Windows, macOS, or Linux
- **Python Support:** LibreOffice must have Python scripting enabled

## Installation & Running

### Step 1: Install LibreOffice (if not already installed)

**Windows/macOS/Linux:**
Download from: https://www.libreoffice.org/download/

### Step 2: Enable Python Scripting

1. Open LibreOffice Calc or Writer
2. Go to **Tools > Options > LibreOffice > Security > Macro Security**
3. Set to "Medium" or "Low" (to allow macros)
4. Click **OK**

### Step 3: Add Python Macros

#### For Calc Macros:

1. Open LibreOffice Calc
2. Go to **Tools > Macros > Organize Macros > Python**
3. Click **New** and create a new library (e.g., "MyLibrary")
4. Select the library and click **New Module**
5. Copy the code from `python_uno_calc.py` into the editor
6. Save and close

#### To Run the Macro:

1. In Calc, go to **Tools > Macros > Run Macro**
2. Select your macro (e.g., "populate_calc")
3. Click **Run**
4. Result: Cells will be populated with data!

### Step 4: Add Writer Macros (Optional)

Follow the same process, but in **LibreOffice Writer** instead of Calc.

### Step 5: Try StarBasic Macros

1. Go to **Tools > Macros > Organize Macros > LibreOffice Basic**
2. Create a new module
3. Paste code from `basic_macro_sample.bas`
4. Save and run

## Troubleshooting

| Issue | Solution |
|-------|----------|
| "Macro not found" | Ensure macro is saved in correct library; refresh macro list |
| "XSCRIPTCONTEXT is not defined" | This is normal for Basic; only Python uses XSCRIPTCONTEXT |
| Macros disabled | Check Tools > Options > Security > Macro Security setting |
| Python module not found | LibreOffice may need Python environment setup; check Tools > Options > LibreOffice > Python |

## Resources

- LibreOffice UNO API Docs: https://api.libreoffice.org/
- Scripting Guide: https://docs.libreoffice.org/
- Community Support: https://ask.libreoffice.org/

---

**Questions?** Refer to `docs/UNO_BASICS.md` for more details on the UNO API.
