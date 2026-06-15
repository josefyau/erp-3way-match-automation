============================================================
'  STAGE 5/6 - Reporting & printing
'  GoToPrintMatchingReport: build the printable report sheet.
'  ConfirmReportAndPrint: set print area for visible rows, print.
'  PrintOutReport: quick landscape print-preview.
'  PrintPaymentRequest: tidy and preview the payment form.
' ============================================================
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