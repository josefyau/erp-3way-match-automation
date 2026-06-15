' ============================================================
'  ERP PO ▸ Invoice 3-Way Match & Payment Automation
'  ------------------------------------------------------------
'  Pipeline:
'    1. PullPOlines ............ import + cleanse raw ERP export
'    2. ArrangeERP_POLinesColumns  reshape into working layout
'    3. MatchINV ............... copy + flag exceptions (cond. format)
'    4. CheckInvoiceAmount ..... refresh validation pivot
'    5. Over45days / Over70days  ageing of received-but-unpaid lines
'    6. Print / Payment-form / Perm_Rec routines
'  ------------------------------------------------------------
'  NOTE: rename worksheet tabs to match the names used here:
'    "ERP_PO lines raw", "ERP_PO_Match_INV", "Invoice_Lines_ERPMatch"
' ============================================================


' =========================== Module1 ===========================
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


' =========================== Module2 ===========================
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


' =========================== Module3 ===========================
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


' =========================== Module4 ===========================
Option Explicit

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


' =========================== Module5 ===========================
Option Explicit

Sub GoToPrintMatchingReport()

    Dim wsSource As Worksheet
    Dim wsDest As Worksheet
    Dim lastRowSource As Long
    Dim lastRowDest As Long

    Set wsSource = ThisWorkbook.Sheets("ERP_PO_Match_INV")
    Set wsDest = ThisWorkbook.Sheets("Print_Matching_Report")

    lastRowSource = wsSource.Cells(wsSource.Rows.Count, "B").End(xlUp).Row

    Application.ScreenUpdating = False

    ' Clear previous report data and standardise row height
    wsDest.Range("A4:T10000").ClearContents
    wsDest.Rows("4:500").RowHeight = 20

    ' Copy matched lines (B:U) into the report (A onward)
    wsSource.Range("B4:U" & lastRowSource).Copy
    wsDest.Range("A4").PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False

    lastRowDest = wsDest.Cells(wsDest.Rows.Count, "A").End(xlUp).Row

    ' Sort by delivery date (R), vendor (C), item (E)
    With wsDest.Sort
        .SortFields.Clear
        .SortFields.Add Key:=wsDest.Range("R4:R" & lastRowDest), Order:=xlAscending
        .SortFields.Add Key:=wsDest.Range("C4:C" & lastRowDest), Order:=xlAscending
        .SortFields.Add Key:=wsDest.Range("E4:E" & lastRowDest), Order:=xlAscending
        .SetRange wsDest.Range("A4:T" & lastRowDest)
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .Apply
    End With

    wsDest.Range("A4:A" & lastRowDest).Rows.AutoFit
    wsDest.Activate

    Application.ScreenUpdating = True

End Sub


' =========================== Module6 ===========================
Option Explicit

Sub ConfirmReportAndPrint()

    Dim wsDest As Worksheet
    Dim lastRowDest As Long
    Dim printRange As Range
    Dim visibleRange As Range

    Set wsDest = ThisWorkbook.Sheets("Print_Matching_Report")

    Application.ScreenUpdating = False

    If wsDest.AutoFilterMode Then wsDest.AutoFilterMode = False
    wsDest.Rows("3:3").AutoFilter

    lastRowDest = wsDest.Cells(wsDest.Rows.Count, "A").End(xlUp).Row
    Set printRange = wsDest.Range("A3:S" & lastRowDest)

    On Error Resume Next
    Set visibleRange = printRange.SpecialCells(xlCellTypeVisible)
    On Error GoTo 0

    wsDest.ResetAllPageBreaks

    If Not visibleRange Is Nothing Then
        wsDest.PageSetup.PrintArea = visibleRange.Address
        With wsDest.PageSetup
            .Orientation = xlLandscape
            .Zoom = False
            .FitToPagesWide = 1
            .FitToPagesTall = 1
            .LeftMargin = Application.InchesToPoints(0.25)
            .RightMargin = Application.InchesToPoints(0.25)
            .TopMargin = Application.InchesToPoints(0.3)
            .BottomMargin = Application.InchesToPoints(0.3)
            .HeaderMargin = Application.InchesToPoints(0.25)
            .FooterMargin = Application.InchesToPoints(0.25)
        End With
        wsDest.PrintOut
    Else
        MsgBox "No visible cells to print.", vbExclamation
    End If

    Application.ScreenUpdating = True

End Sub


' =========================== Module7 ===========================
Option Explicit

Sub PrintOutReport()

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim rng As Range

    Set ws = ThisWorkbook.Worksheets("Print_Matching_Report")
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    Set rng = ws.Range("A4:T" & lastRow)

    rng.Rows.AutoFit
    ws.PageSetup.PrintArea = "$A$3:$T$" & lastRow

    With ws.PageSetup
        .Orientation = xlLandscape
        .LeftMargin = Application.InchesToPoints(0.25)
        .RightMargin = Application.InchesToPoints(0.25)
        .TopMargin = Application.InchesToPoints(0.75)
        .BottomMargin = Application.InchesToPoints(0.75)
        .FitToPagesWide = 1
        .FitToPagesTall = False
    End With

    ws.PrintPreview

End Sub


' =========================== Module8 ===========================
Option Explicit

Sub PrintPaymentRequest()

    Dim lastRow As Long
    Dim emptyCell As Range
    Dim rng As Range
    Dim i As Long
    Dim hasPageBreak As Boolean

    Application.ScreenUpdating = False

    ' Reset / unhide rows from any previous payment form
    Rows("15:164").EntireRow.Hidden = False

    lastRow = Cells(Rows.Count, "B").End(xlUp).Row

    ' Hide the unused tail of the line-item block (B28:B164)
    Set rng = Range("B28:B164")
    Set emptyCell = rng.Find("", LookIn:=xlValues, LookAt:=xlWhole)

    If Not emptyCell Is Nothing Then
        If emptyCell.Row <= rng.Cells(rng.Rows.Count).Row Then
            Range(emptyCell, rng.Cells(rng.Rows.Count)).EntireRow.Hidden = True
        End If
    End If

    ' If a page break sits in rows 165-178, keep the footer rows visible
    With ActiveSheet
        hasPageBreak = False
        For i = 165 To 178
            If .Rows(i).PageBreak <> xlPageBreakNone Then
                hasPageBreak = True
                Exit For
            End If
        Next i
        If hasPageBreak Then .Rows("156:164").Hidden = False
    End With

    Application.ScreenUpdating = True
    Application.Dialogs(xlDialogPrintPreview).Show

End Sub


' =========================== Module9 ===========================
Option Explicit

Sub CopyMatchingToRec()

    Dim wsSource As Worksheet
    Dim wsDest As Worksheet
    Dim lastRowSource As Long
    Dim lastRowDest As Long

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual

    Set wsSource = ThisWorkbook.Sheets("ERP_PO_Match_INV")
    Set wsDest = ThisWorkbook.Sheets("Matching_Rec")

    lastRowSource = wsSource.Cells(wsSource.Rows.Count, "D").End(xlUp).Row
    lastRowDest = wsDest.Cells(wsDest.Rows.Count, "D").End(xlUp).Row + 1

    wsSource.Range("A4:W" & lastRowSource).Copy
    wsDest.Range("A" & lastRowDest).PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False

    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic

End Sub


' =========================== Module10 ===========================
Option Explicit

Sub TransformMatchedLinesFormat()

    Dim wsSource_ERP_POmatch As Worksheet
    Dim wsDest_Transformer As Worksheet
    Dim lastrowDestTransformer As Long

    Application.ScreenUpdating = False
    Application.CutCopyMode = False
    Application.Calculation = xlCalculationAutomatic

    Set wsSource_ERP_POmatch = ThisWorkbook.Sheets("ERP_PO_Match_INV")
    Set wsDest_Transformer = ThisWorkbook.Sheets("Invoice_Lines_ERPMatch")

    ' Guard: nothing matched yet
    If IsEmpty(wsSource_ERP_POmatch.Range("S4").Value) Or _
       IsEmpty(wsSource_ERP_POmatch.Range("D4").Value) Then
        MsgBox "No matched PO-to-Invoice lines! Check ERP_PO_Match_INV!"
        Application.ScreenUpdating = True
        Exit Sub
    End If

    ' Reset source filters / hidden rows; clear destination
    If wsSource_ERP_POmatch.AutoFilterMode Then wsSource_ERP_POmatch.AutoFilterMode = False
    wsSource_ERP_POmatch.Rows.Hidden = False
    wsSource_ERP_POmatch.Rows("3:3").AutoFilter
    wsDest_Transformer.Range("A4:I10000").ClearContents

    ' Unique invoice numbers (D) via spill, then hardcode
    wsDest_Transformer.Range("D4").Formula2 = _
        "=SORT(FILTER(UNIQUE(ERP_PO_Match_INV!$S$4:$S$500),UNIQUE(ERP_PO_Match_INV!$S$4:$S$500)<>""""))"

    lastrowDestTransformer = wsDest_Transformer.Cells(wsDest_Transformer.Rows.Count, "D").End(xlUp).Row

    With wsDest_Transformer
        .Range("D4:D" & lastrowDestTransformer).Value = .Range("D4:D" & lastrowDestTransformer).Value

        ' Vendor code
        .Range("A4:A" & lastrowDestTransformer).Formula = _
            "=XLOOKUP(D4, ERP_PO_Match_INV!$S$4:$S$500, ERP_PO_Match_INV!$B$4:$B$500, ""NF"")"
        .Range("A4:A" & lastrowDestTransformer).Value = .Range("A4:A" & lastrowDestTransformer).Value

        ' Vendor name
        .Range("B4:B" & lastrowDestTransformer).Formula = _
            "=XLOOKUP(D4, ERP_PO_Match_INV!$S$4:$S$500, ERP_PO_Match_INV!$C$4:$C$500, ""NF"")"
        .Range("B4:B" & lastrowDestTransformer).Value = .Range("B4:B" & lastrowDestTransformer).Value

        ' Currency
        .Range("C4:C" & lastrowDestTransformer).Formula = _
            "=XLOOKUP(D4,ERP_PO_Match_INV!$S$4:$S$500,ERP_PO_Match_INV!$J$4:$J$500,""NF"")"
        .Range("C4:C" & lastrowDestTransformer).Value = .Range("C4:C" & lastrowDestTransformer).Value

        ' Corresponding PO(s)
        .Range("E4:E" & lastrowDestTransformer).Formula2 = _
            "=""PO000""&TEXTJOIN("", "",TRUE,RIGHT(SORT(UNIQUE(FILTER(ERP_PO_Match_INV!$D$4:$D$500,ERP_PO_Match_INV!$S$4:$S$500=Invoice_Lines_ERPMatch!D4))),5))"
        .Range("E4:E" & lastrowDestTransformer).Value = .Range("E4:E" & lastrowDestTransformer).Value

        ' Invoice date
        .Range("F4:F" & lastrowDestTransformer).Formula = _
            "=XLOOKUP(D4, ERP_PO_Match_INV!$S$4:$S$500, ERP_PO_Match_INV!$T$4:$T$500, ""NF"")"
        .Range("F4:F" & lastrowDestTransformer).Value = .Range("F4:F" & lastrowDestTransformer).Value

        ' Invoice amount (sum)
        .Range("G4:G" & lastrowDestTransformer).Formula = _
            "=SUMIFS(ERP_PO_Match_INV!O:O, ERP_PO_Match_INV!S:S, Invoice_Lines_ERPMatch!D4)"
        .Range("G4:G" & lastrowDestTransformer).Value = .Range("G4:G" & lastrowDestTransformer).Value

        ' Received date
        .Range("H4:H" & lastrowDestTransformer).Formula = _
            "=XLOOKUP(D4, ERP_PO_Match_INV!$S$4:$S$500, ERP_PO_Match_INV!$V$4:$V$500, ""NF"")"
        .Range("H4:H" & lastrowDestTransformer).Value = .Range("H4:H" & lastrowDestTransformer).Value

        ' Processed date
        .Range("I4:I" & lastrowDestTransformer).Formula = _
            "=XLOOKUP(D4, ERP_PO_Match_INV!$S$4:$S$500, ERP_PO_Match_INV!$W$4:$W$500, ""NF"")"
        .Range("I4:I" & lastrowDestTransformer).Value = .Range("I4:I" & lastrowDestTransformer).Value
    End With

    wsDest_Transformer.Activate
    Application.CutCopyMode = False
    Application.ScreenUpdating = True

End Sub


' =========================== Module11 ===========================
Option Explicit

Sub SaveAndExit()

    Dim wb As Workbook
    Dim wbCount As Integer

    ThisWorkbook.Sheets("Arrange PO Lines").Activate
    ThisWorkbook.Save

    wbCount = 0
    For Each wb In Application.Workbooks
        wbCount = wbCount + 1
    Next wb

    If wbCount = 1 Then
        Application.Quit
    Else
        ThisWorkbook.Close
    End If

End Sub


' =========================== Module12 ===========================
Option Explicit

Sub CopyVendorToPayForm()

    Dim wsSource_Transformer As Worksheet
    Dim wsDest_PayForm As Worksheet
    Dim payvendor As String

    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationAutomatic

    Set wsSource_Transformer = ThisWorkbook.Sheets("Invoice_Lines_ERPMatch")
    Set wsDest_PayForm = ThisWorkbook.Sheets("New_Form")

    payvendor = wsSource_Transformer.Cells(wsSource_Transformer.Rows.Count, "B").End(xlUp).Value
    wsDest_PayForm.Range("C11").Value = payvendor

    Application.ScreenUpdating = True
    wsDest_PayForm.Activate

End Sub


' =========================== Module15 ===========================
Option Explicit

Sub SpecialVendorSortByPA()

    Dim ws As Worksheet
    Dim sortRange As Range

    Set ws = ThisWorkbook.Sheets("ERP_PO_Match_INV")

    Set sortRange = ws.Range("A3", _
        ws.Cells(ws.Rows.Count, "A").End(xlUp).Resize(, ws.UsedRange.Columns.Count))

    sortRange.Sort Key1:=ws.Columns("R"), Order1:=xlAscending, _
                   Key2:=ws.Columns("D"), Order2:=xlAscending, _
                   Key3:=ws.Columns("F"), Order3:=xlAscending, _
                   Header:=xlYes

End Sub


' =========================== Module16 ===========================
'  Overdue-receipt ageing. Over45days / Over70days are thin wrappers
'  over one shared routine so the logic lives in a single place.
Option Explicit

Sub Over45days(): ShowOverdueReceipts 45: End Sub
Sub Over70days(): ShowOverdueReceipts 70: End Sub

Private Sub ShowOverdueReceipts(daysOverdue As Long)

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim vendorCodes As Variant
    Dim i As Long
    Dim visibleRows As Range

    Application.ScreenUpdating = False
    Application.CutCopyMode = False

    Set ws = ThisWorkbook.Sheets("Arrange PO Lines")
    ws.AutoFilterMode = False
    ws.Cells.EntireColumn.Hidden = False
    ws.Cells.EntireRow.Hidden = False

    vendorCodes = GetExcludedVendors()
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    ' Strip out excluded vendors (column B), recomputing the last row each pass
    For i = LBound(vendorCodes) To UBound(vendorCodes)
        ws.Range("A3:R" & lastRow).AutoFilter Field:=2, Criteria1:=CStr(vendorCodes(i))
        On Error Resume Next
        Set visibleRows = ws.Range("A4:R" & lastRow).SpecialCells(xlCellTypeVisible)
        On Error GoTo 0
        If Not visibleRows Is Nothing Then visibleRows.EntireRow.Delete
        Set visibleRows = Nothing
        ws.AutoFilterMode = False
        lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row
    Next i

    ' Filter: received lines whose delivery date is older than the threshold
    ws.Range("A3:R" & lastRow).AutoFilter
    ws.Range("A3:R" & lastRow).AutoFilter Field:=1, Criteria1:="Received"
    ws.Range("A3:R" & lastRow).AutoFilter Field:=18, Criteria1:="<=" & Date - daysOverdue

    On Error Resume Next
    Set visibleRows = ws.Range("A4:R" & lastRow).SpecialCells(xlCellTypeVisible)
    On Error GoTo 0

    If visibleRows Is Nothing Then
        MsgBox "All shipments received within " & daysOverdue & " days have been paid!"
    Else
        With ws.Sort
            .SortFields.Clear
            .SortFields.Add Key:=ws.Range("R4:R" & lastRow), Order:=xlAscending  ' Delivery date
            .SortFields.Add Key:=ws.Range("C4:C" & lastRow), Order:=xlAscending  ' Vendor name
            .SortFields.Add Key:=ws.Range("D4:D" & lastRow), Order:=xlAscending  ' Item code
            .SetRange ws.Range("A4:R" & lastRow)
            .Header = xlNo
            .MatchCase = False
            .Orientation = xlTopToBottom
            .SortMethod = xlPinYin
            .Apply
        End With
    End If

    Application.ScreenUpdating = True

End Sub

' Intercompany / non-payable vendor codes excluded from ageing.
' PLACEHOLDERS for the portfolio copy - replace privately with real codes,
' or repoint this to a config range (e.g. directory!ExcludedVendors).
Private Function GetExcludedVendors() As Variant
    GetExcludedVendors = Array("0001", "0002", "0003", _
                               "1000001", "1000002", "1000003", "1000004")
End Function


' =========================== Module18 ===========================
Option Explicit

Sub CopyMatchedLinesToPermRecAndCheckDuplicates()

    Dim wsSource_ERP_POmatch As Worksheet
    Dim wsDest_PermRec As Worksheet
    Dim lastrowSourceERP_POMatch As Long
    Dim lastrowDestPermRecBeforePaste As Long
    Dim lastrowDestPermRecAfterPaste As Long
    Dim dataRange As Range
    Dim dataArray As Variant
    Dim dict As Object
    Dim duplicateKeys As Collection
    Dim key As String
    Dim k As Variant
    Dim foundDuplicate As Boolean
    Dim i As Long

    Application.ScreenUpdating = False
    Application.CutCopyMode = False
    Application.Calculation = xlCalculationManual

    Set wsSource_ERP_POmatch = ThisWorkbook.Sheets("ERP_PO_Match_INV")
    Set wsDest_PermRec = ThisWorkbook.Sheets("Perm_Rec")
    Set dict = CreateObject("Scripting.Dictionary")

    lastrowSourceERP_POMatch = wsSource_ERP_POmatch.Cells(wsSource_ERP_POmatch.Rows.Count, "B").End(xlUp).Row
    lastrowDestPermRecBeforePaste = wsDest_PermRec.Cells(wsDest_PermRec.Rows.Count, "B").End(xlUp).Row

    ' Unhide before pasting
    wsSource_ERP_POmatch.Rows.Hidden = False
    wsDest_PermRec.Rows.Hidden = False

    ' Append matched lines to the permanent record
    wsSource_ERP_POmatch.Range("A4:W" & lastrowSourceERP_POMatch).Copy
    wsDest_PermRec.Range("A" & lastrowDestPermRecBeforePaste + 1).PasteSpecial Paste:=xlPasteValues
    Application.CutCopyMode = False

    ' Load the full record into an array for duplicate detection
    lastrowDestPermRecAfterPaste = wsDest_PermRec.Cells(wsDest_PermRec.Rows.Count, "B").End(xlUp).Row
    Set dataRange = wsDest_PermRec.Range("A4:W" & lastrowDestPermRecAfterPaste)
    dataArray = dataRange.Value

    ' Build a composite key: vendor | PO | item | amount | invoice no.
    For i = LBound(dataArray, 1) To UBound(dataArray, 1)
        key = CStr(dataArray(i, 2)) & "|" & CStr(dataArray(i, 4)) & "|" & _
              CStr(dataArray(i, 6)) & "|" & Format(dataArray(i, 14), "0.00") & "|" & _
              CStr(dataArray(i, 19))
        If dict.exists(key) Then
            dict(key) = dict(key) + 1
        Else
            dict(key) = 1
        End If
    Next i

    ' Collect the keys that occur more than once
    Set duplicateKeys = New Collection
    For Each k In dict.Keys
        If dict(k) > 1 Then duplicateKeys.Add k
    Next k

    If duplicateKeys.Count > 0 Then
        foundDuplicate = True
        wsDest_PermRec.Activate
        wsDest_PermRec.Range("A3:W" & lastrowDestPermRecAfterPaste).AutoFilter

        ' Hide non-duplicate rows so only the offenders show
        For i = LBound(dataArray, 1) To UBound(dataArray, 1)
            key = CStr(dataArray(i, 2)) & "|" & CStr(dataArray(i, 4)) & "|" & _
                  CStr(dataArray(i, 6)) & "|" & Format(dataArray(i, 14), "0.00") & "|" & _
                  CStr(dataArray(i, 19))
            If Not dict.exists(key) Or dict(key) <= 1 Then
                wsDest_PermRec.Rows(i + 3).Hidden = True
            End If
        Next i

        MsgBox "DUPLICATE invoice lines! Check entries in Perm_Rec!"
    Else
        wsDest_PermRec.Activate
    End If

    ' Reapply header filter on the source and land the cursor on the next free row
    wsSource_ERP_POmatch.Range("A3:W3").AutoFilter
    wsDest_PermRec.Cells(wsDest_PermRec.Rows.Count, 1).End(xlUp).Offset(1, 0).Activate

    Application.CutCopyMode = False
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic

End Sub


' =========================== Module19 ===========================
Option Explicit

Sub ShowAllinPermRec()

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Perm_Rec")

    ws.Rows.Hidden = False

    If ws.AutoFilterMode Then ws.AutoFilterMode = False
    ws.Rows(3).AutoFilter

    ws.Activate
    ws.Range("A4").Select

End Sub
