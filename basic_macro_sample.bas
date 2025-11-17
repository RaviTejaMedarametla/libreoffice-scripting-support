REM LibreOffice StarBasic Macro Example
REM Demonstrates classic macro automation in Calc

Sub AutomateCalcBasic
    Dim oDoc As Object
    Dim oSheet As Object
    Dim oCell As Object
    
    REM Get active document and sheet
    oDoc = ThisComponent
    oSheet = oDoc.CurrentController.ActiveSheet
    
    REM Write to cells using Basic
    oSheet.getCellRangeByName("A1").String = "Basic Macro Demo"
    oSheet.getCellRangeByName("B1").Value = 123
    
    REM Display message
    MsgBox "Basic macro executed! Check cells A1 and B1.", 0, "Success"
End Sub

Sub HelloWorldBasic
    REM Simple hello world macro
    MsgBox "Hello from LibreOffice StarBasic!", 0, "Greeting"
End Sub
