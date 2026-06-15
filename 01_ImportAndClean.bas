============================================================
'  STAGE 1 - Import & Cleanse
'  PullPOlines: import raw ERP export, strip cancelled/excluded
'  lines, fix text-formatted numbers, look up vendor names.
'  DeleteFilteredRows: shared helper used throughout this stage.
' ============================================================
Option Explicit

Sub PullPOlines()

    Dim wbSource As Workbook
    Dim wsSource As Worksheet
    Dim wsDestination As Worksheet
    Dim sourcePath As String
    Dim lastRow As Long

    ' Step 1. Optimise performance
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.CutCopyMode = False

    ' Step 2. Open the source workbook (path read from a config cell, never hard-coded)
    sourcePath = ThisWorkbook.Worksheets("directory").Range("B1").Value
    Set wbSource = Workbooks.Open(sourcePath)
    Set wsSource = wbSource.Sheets("Sheet1")
    Set wsDestination = ThisWorkbook.Sheets("ERP_PO lines raw")

    ' Step 3. Clear previous data in the destination sheet
    wsDestination.Range("A3:W10000").ClearContents

    ' Step 4. Copy source values into the destination sheet
    wsSource.Range("A1:U10000").Copy
    wsDestination.Range("A3").PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False

    ' Step 5. Close the source workbook without saving
    wbSource.Close SaveChanges:=False

    ' Step 6. Reset the header filter on row 3
    wsDestination.AutoFilterMode = False
    wsDestination.Rows("3:3").AutoFilter

    ' Step 7. Remove "Canceled" PO lines (column A = line status)
    DeleteFilteredRows wsDestination, 1, "Canceled"

    ' Step 8. Remove INKJET / COPACK items (column H = item number)
    DeleteFilteredRows wsDestination, 8, "INKJET", "COPACK"

    ' Step 9. Remove rows with a blank item number (column H)
    DeleteFilteredRows wsDestination, 8, "="

    ' Step 10. Reset the header filter
    wsDestination.AutoFilterMode = False
    wsDestination.Rows("3:3").AutoFilter

    ' Step 11. Last populated row after the deletions
    lastRow = wsDestination.Cells(wsDestination.Rows.Count, "A").End(xlUp).Row

    ' Step 12. Convert text-formatted numbers in columns H and E to real numbers
    With wsDestination
        .Range("W4:W" & lastRow).Formula = "=IFERROR(IF(H4="""","""",VALUE(H4)),H4)"
        .Range("X4:X" & lastRow).Formula = "=IFERROR(IF(E4="""","""",VALUE(E4)),E4)"
        .Range("H4:H" & lastRow).Value = .Range("W4:W" & lastRow).Value
        .Range("E4:E" & lastRow).Value = .Range("X4:X" & lastRow).Value
        .Range("W4:X" & lastRow).ClearContents
    End With

    ' Step 13. Look up vendor name per vendor code, then convert to static values
    With wsDestination
        .Range("V4").Formula = "=XLOOKUP(E4,VendorInfo!A:A,VendorInfo!B:B,""not found"")"
        .Range("V4:V" & lastRow).FillDown
        .Range("V4:V" & lastRow).Value = .Range("V4:V" & lastRow).Value
    End With

    ' Restore application settings
    Application.Calculation = xlCalculationAutomatic
    Application.ScreenUpdating = True

End Sub

' Filters a sheet on one field and deletes every visible data row (row 4 down).
' Handles single-value, OR, and blank ("=") criteria from one place.
Private Sub DeleteFilteredRows(ws As Worksheet, fieldIndex As Long, _
                               crit1 As String, Optional crit2 As String = "")

    Dim lastRow As Long
    Dim visibleData As Range

    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    If lastRow < 4 Then Exit Sub          ' only the header is present

    ws.AutoFilterMode = False

    If Len(crit2) > 0 Then
        ws.Range("A3:V" & lastRow).AutoFilter Field:=fieldIndex, _
            Criteria1:=crit1, Operator:=xlOr, Criteria2:=crit2
    Else
        ws.Range("A3:V" & lastRow).AutoFilter Field:=fieldIndex, Criteria1:=crit1
    End If

    On Error Resume Next
    Set visibleData = ws.Range("A4:A" & lastRow).SpecialCells(xlCellTypeVisible)
    On Error GoTo 0

    If Not visibleData Is Nothing Then visibleData.EntireRow.Delete

    ws.AutoFilterMode = False

End Sub