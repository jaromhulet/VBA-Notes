# VBA-Notes
Notes of Common VBA Code I Commonly Use

# set variable as number of worksheets in workbook
wrksheet_num = ThisWorkbook.Sheets.Count

# If statement for AutoFilterMode (Auto Filter)
If ActiveSheet.AutoFilterMode Then
    Else
End If

# Clear Contents
Selection.ClearContents

#Format interior of cells
    With Selection.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With

#Select all cells
Cells.Select

# Run this before opening another workbook
Application.Run "ConnectChartEvents"

# Open workbook
Workbooks.Open Filename:= _
        "C:\excel_file\excel.xls"

# Select another workbook
Windows("excel.xls").Activate

# Close Workbook
Workbooks("excel.xls").Close

# save active workbook
ActiveWorkbook.Save

# Put a formula in the active cell
ActiveCell.FormulaR1C1= "=VLOOKUP(RC[-1],C[-6]:C[-4],2,FALSE)"

# Turn off screen flashing
Application.ScreenUpdating = False

# Paste Values ( paste special menu)
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False

# Autofit all columns on a worksheet
    Cells.EntireColumn.AutoFit

# Change format (style) of cells
Selection.Style = "Percent"

# Select down, equivilent of ctrl + shft + down arrow
    Selection.End(xlDown).Select

# Turn off alert boxes
Application.DisplayAlerts = False

