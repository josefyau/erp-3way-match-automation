============================================================
'  STAGE 3 - Match & flag exceptions
'  MatchINV: copy working lines, sort, apply the conditional-
'  formatting exception rules.
'  CheckInvoiceAmount: refresh the validation pivot.
' ============================================================
Option Explicit

Sub MatchINV()

    Dim wsSource As Worksheet
    Dim wsDest As Worksheet
    Dim lastRow As Long

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationAutomatic

    Set wsSource = ThisWorkbook.Sheets("Arrange PO Lines")
    Set wsDest = ThisWorkbook.Sheets("ERP_PO_Match_INV")

    ' Step 1: Reset autofilters in destination
    wsDest.AutoFilterMode = False
    wsDest.Range("A3:U3").AutoFilter

    ' Step 2: Clear contents and conditional formatting in destination
    wsDest.Range("A4:W10000").ClearContents
    wsDest.Cells.FormatConditions.Delete

    ' Step 3: Last row of the source
    lastRow = wsSource.Cells(wsSource.Rows.Count, "A").End(xlUp).Row

    ' Step 4: Copy visible source rows (A:R) as values, then reset source filter
    wsSource.Range("A4:R" & lastRow).SpecialCells(xlCellTypeVisible).Copy
    wsDest.Range("A4").PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False

    wsSource.AutoFilterMode = False
    wsSource.Range("A3:R3").AutoFilter

    ' Step 5: Sort by delivery date (R) then item number (F)
    With wsDest.Sort
        .SortFields.Clear
        .SortFields.Add Key:=wsDest.Range("R4:R" & lastRow), Order:=xlAscending
        .SortFields.Add Key:=wsDest.Range("F4:F" & lastRow), Order:=xlAscending
        .SetRange wsDest.Range("A4:T" & lastRow)
        .Header = xlNo
        .Apply
    End With

    ' Step 6: Flag lines that are not "Received" or have no GR quantity
    With wsDest.Range("$A$4:$R$300").FormatConditions.Add(Type:=xlExpression, _
        Formula1:="=AND($B4<>"""", OR($A4<>""Received"", $I4=""""))")
        .Interior.Color = RGB(255, 0, 0)
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
    End With

    ' Step 7: Flag a received-amount mismatch (N <> O, unless "No GR")
    With wsDest.Range("$N$4:$O$300").FormatConditions.Add(Type:=xlExpression, _
        Formula1:="=AND($N4<>$O4, $O4<>""No GR"")")
        .Interior.Color = RGB(255, 0, 0)
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
    End With

    ' Step 8: Highlight any line carrying a remark (column U)
    With wsDest.Range("$A$4:$U$300").FormatConditions.Add(Type:=xlExpression, _
        Formula1:="=$U4<>""""")
        .Interior.Color = RGB(251, 226, 213)
        .Font.Color = RGB(0, 0, 0)
    End With

    ' Step 9: Land the cursor on the entry column
    wsDest.Activate
    wsDest.Range("S4").Select

    Application.ScreenUpdating = True

End Sub

Sub CheckInvoiceAmount()

    Dim ws As Worksheet
    Dim pt As PivotTable

    Set ws = ThisWorkbook.Sheets("Check_INV_Amount")
    ws.Activate

    On Error Resume Next
    Set pt = ws.PivotTables("INV_CHECK_PIVOTTABLE")
    On Error GoTo 0

    If Not pt Is Nothing Then pt.RefreshTable

End Sub