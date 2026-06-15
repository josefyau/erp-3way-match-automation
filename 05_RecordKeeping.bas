' ============================================================
'  Record keeping
'  CopyMatchingToRec: append matched lines to Matching_Rec.
'  CopyMatchedLinesToPermRecAndCheckDuplicates: append to the
'  permanent record and flag duplicate invoice lines.
'  ShowAllinPermRec: unhide / reset the Perm_Rec view.
' ============================================================
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

Sub ShowAllinPermRec()

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets("Perm_Rec")

    ws.Rows.Hidden = False

    If ws.AutoFilterMode Then ws.AutoFilterMode = False
    ws.Rows(3).AutoFilter

    ws.Activate
    ws.Range("A4").Select

End Sub