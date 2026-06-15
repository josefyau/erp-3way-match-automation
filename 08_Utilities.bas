' ============================================================
'  Utilities
'  SpecialVendorSortByPA: re-sort the match sheet by date/PO/item.
'  SaveAndExit: save the workbook and close (or quit Excel).
' ============================================================
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