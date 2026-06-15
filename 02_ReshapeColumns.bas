============================================================
'  STAGE 2 - Reshape into working layout
'  ArrangeERP_POLinesColumns: remap raw columns into the
'  working layout, derive unit price / received amount, sort.
' ============================================================
Option Explicit

Sub ArrangeERP_POLinesColumns()

    Dim wsSource As Worksheet
    Dim wsDest As Worksheet
    Dim lastRow As Long

    Application.ScreenUpdating = False
    Application.CutCopyMode = False

    Set wsSource = ThisWorkbook.Sheets("ERP_PO lines raw")
    Set wsDest = ThisWorkbook.Sheets("Arrange PO Lines")

    ' Step 1: Remove all filters and reapply on the header row
    wsSource.AutoFilterMode = False
    wsSource.Range("A3:V3").AutoFilter
    wsDest.AutoFilterMode = False
    wsDest.Range("A3:R3").AutoFilter

    ' Step 2: Clear contents in the destination range
    wsDest.Range("A4:S10000").ClearContents

    lastRow = wsSource.Cells(wsSource.Rows.Count, "A").End(xlUp).Row

    ' Step 3: line status  (raw A -> arranged A)
    wsSource.Range("A4:A" & lastRow).Copy
    wsDest.Range("A4").PasteSpecial Paste:=xlPasteValues

    ' Step 4: vendor code  (raw E -> arranged B)
    wsSource.Range("E4:E" & lastRow).Copy
    wsDest.Range("B4").PasteSpecial Paste:=xlPasteValues

    ' Step 5: vendor name  (raw V -> arranged C)
    wsSource.Range("V4:V" & lastRow).Copy
    wsDest.Range("C4").PasteSpecial Paste:=xlPasteValues

    ' Step 6: PO, line no., item no., UOM, PO qty  (raw F:J -> arranged D:H)
    wsSource.Range("F4:J" & lastRow).Copy
    wsDest.Range("D4").PasteSpecial Paste:=xlPasteValues

    ' Step 7: received quantity  (raw L -> arranged I)
    wsSource.Range("L4:L" & lastRow).Copy
    wsDest.Range("I4").PasteSpecial Paste:=xlPasteValues

    ' Step 8: currency  (raw N -> arranged J)
    wsSource.Range("N4:N" & lastRow).Copy
    wsDest.Range("J4").PasteSpecial Paste:=xlPasteValues

    ' Step 9: purchase unit  (raw M -> arranged K)
    wsSource.Range("M4:M" & lastRow).Copy
    wsDest.Range("K4").PasteSpecial Paste:=xlPasteValues

    ' Step 10: price for x1 purchase unit  (raw O -> arranged L)
    wsSource.Range("O4:O" & lastRow).Copy
    wsDest.Range("L4").PasteSpecial Paste:=xlPasteValues

    ' Step 11: unit price (price of x1 Ea) -> arranged M
    wsDest.Range("M4").Formula = "=L4/K4"
    wsDest.Range("M4:M" & lastRow).FillDown
    wsDest.Range("M4:M" & lastRow).Value = wsDest.Range("M4:M" & lastRow).Value

    ' Step 12: ERP net amount  (raw P -> arranged N)
    wsSource.Range("P4:P" & lastRow).Copy
    wsDest.Range("N4").PasteSpecial Paste:=xlPasteValues

    ' Step 13: received amount -> arranged O
    wsDest.Range("O4").Formula = "=IF(I4="""",""No GR"",L4/K4*I4)"
    wsDest.Range("O4:O" & lastRow).FillDown
    wsDest.Range("O4:O" & lastRow).Value = wsDest.Range("O4:O" & lastRow).Value

    ' Step 14: warehouse and ASN  (raw T:U -> arranged P:Q)
    wsSource.Range("T4:U" & lastRow).Copy
    wsDest.Range("P4").PasteSpecial Paste:=xlPasteValues

    ' Step 15: delivery date  (raw S -> arranged R)
    wsSource.Range("S4:S" & lastRow).Copy
    wsDest.Range("R4").PasteSpecial Paste:=xlPasteValues

    Application.CutCopyMode = False

    ' Step 16: Sort by vendor (C), delivery date (R), item number (F)
    With wsDest.Sort
        .SortFields.Clear
        .SortFields.Add Key:=wsDest.Range("C4:C" & lastRow), Order:=xlAscending
        .SortFields.Add Key:=wsDest.Range("R4:R" & lastRow), Order:=xlAscending
        .SortFields.Add Key:=wsDest.Range("F4:F" & lastRow), Order:=xlAscending
        .SetRange wsDest.Range("A4:S" & lastRow)
        .Header = xlNo
        .Apply
    End With

    Application.ScreenUpdating = True

End Sub