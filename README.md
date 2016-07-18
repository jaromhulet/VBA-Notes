# VBA-Notes
Notes of Common VBA Code I Commonly Use

# set variable as number of worksheets in workbook
wrksheet_num = ThisWorkbook.Sheets.Count

# Set variable equal to the row number of the last row (the last non-null cell)
LastRow = Cells.Find(What:="*", After:=Range("A1"), SearchOrder:=xlByRows, SearchDirection:=xlPrevious, MatchCase:=False).Row

# set varaible equal to the number of adjacent cells in column 1
HomeLoop = Worksheets(1).Cells(Rows.Count, 1).End(xlUp).Row

# Set variable equal to a value in a cell.
Num = Cells(1, 10).Value

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

# Select down, equivilent of ctrl + down arrow
    Selection.End(xlDown).Select

# Selection down, equivalent of ctrl + Shift + down arrow (right arrow)
    Range(Selection, Selection.End(xlDown)).Select
    Range(Selection, Selection.End(xlToRight)).Select

# Turn off alert boxes
Application.DisplayAlerts = False

# Disable events
Application.EnableEvents = False

# Unprotect worksheet with password
    ActiveSheet.Unprotect "password123"
    
# Run another macro
Application.Run ("Macro_Name")

# Countif Function
Application.WorksheetFunction.CountIf(Sheets(1).Range("A:A"), Sheet6.Cells(x + 1, 7))

# test for if active cell is null
If ActiveCell.Value = vbNullString Then
Else
End If

# Auto fill based on another column - based on column B, starting at range D3:J3 as far down as B is.
Range("D3:J3").autofill Destination:=Range("D3:J" & Range("B" & Rows.Count).End(xlUp).Row)

# Find if string contains certain characters will result in boolean
InStr(Cells(x, 2), "string")

# Refresh workbook formulas and links
Application.Run "RefreshEntireWorkbook"

# make columns values
Columns("c").Value = Columns("C").Value

# inbedding code in VBA
       pass = InputBox("Enter Password")
        If pass <> "password123" Then
            MsgBox "Password is Not Correct"
            Exit Sub
        End If
        
# Event code for leaving the worksheet (this code asks to to protect worksheet before leaving the worksheet)
Private Sub Worksheet_Deactivate()
       If Sheets(1).ProtectContents = False Then
           MsgBox "Please Protect 'Sheet 1' Worksheet Before Moving to Another"
           Sheets(1).Select
       End If
End Sub

# Delete and shift up
Selection.Delete Shift:=xlUp

# Call macro when cell is changed (A1 in this case)
Private Sub Worksheet_Change(ByVal Target As Range)

If Target.Address = "$A$1" Then

Call called_macro

End If

End Sub
