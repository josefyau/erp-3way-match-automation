' ============================================================
'  Invoice summary & payment form
'  TransformMatchedLinesFormat: collapse matched lines into one
'  row per invoice number (vendor, PO list, amount, dates).
'  CopyVendorToPayForm: push the vendor onto the payment form.
' ============================================================
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