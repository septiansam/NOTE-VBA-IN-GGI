
    Public WB1 As Workbook
    Public SH1_HOME As Worksheet
    Public SH1_IT_INV As Worksheet, SH1_MONITORING As Worksheet, SH1_REKAP_DETAIL As Worksheet, SH1_RWB As Worksheet
    Public SH1_ITEM_MASTER As Worksheet, SH1_TO_PVT As Worksheet, SH1_IC As Worksheet
    Public SH1_IC2 As Worksheet, SH1_PVT As Worksheet, SHI_PASTE_PVT As Worksheet
    Public SH1_OLAH As Worksheet, SH1_OLAH2 As Worksheet
    Public SH1 As Worksheet
    
    Public LR1 As Long, LC1 As Long, i As Long, j As Long
    Public LR1_2 As Long, FirstRow As Long
    Public SumLoop As Long, RowPaste As Long
    
    Public lookupValue As String
    Public result As String
    Public cell As Range, rng As Range
    
    Public pv_Rng As Range
    Public pv_Cache As PivotCache
    Public pv_Tb As PivotTable
    
    
Sub BUTTON_Execute()
    Application.DisplayAlerts = False
    Application.AskToUpdateLinks = False
    
    Call DeleteSheetsExcept("HOME", "IT INV", "MONITORING", "REKAP DETAIL", "RWB", "ITEM MASTER", "TO PVT", "IC")
    Call Initial_Sheets
    Call CreatePivot_ForGetKonker
    Call CreatePivotReport
    Call GetKonker_ToIC
    Call PengolahanUntukIC2
    Call PengolahanUtama

    Application.DisplayAlerts = True
    Application.AskToUpdateLinks = True
End Sub

Sub PengolahanUtama()

    SH1_PVT.Activate
    LR1 = SH1_PVT.Cells.Find("*", , xlFormulas, xlPart, xlByRows, xlPrevious).Row
    LC1 = SH1_PVT.Cells.Find("*", , xlFormulas, xlPart, xlByColumns, xlPrevious).Column
    SH1_PVT.Range(SH1_PVT.Cells(1, 1), SH1_PVT.Cells(LR1, LC1)).Copy
    
    SHI_PASTE_PVT.Activate
    SHI_PASTE_PVT.Range("A1").PasteSpecial xlPasteValuesAndNumberFormats: Application.CutCopyMode = False
    SHI_PASTE_PVT.Rows(1).ClearContents
    Cells(1, 1).Select
    
    LR1 = SHI_PASTE_PVT.Cells.Find("*", , xlFormulas, xlPart, xlByRows, xlPrevious).Row
    
    SHI_PASTE_PVT.Range("D:D").Insert
    SHI_PASTE_PVT.Range("D2").FormulaR1C1 = "2ND ITEM"
    SHI_PASTE_PVT.Range("D3:D" & LR1).FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-1],RWB!C3:C8,6,0),"""")"
    
    SHI_PASTE_PVT.Range("F:F").Insert
    SHI_PASTE_PVT.Range("F2").FormulaR1C1 = "DES 1"
    SHI_PASTE_PVT.Range("F3:F" & LR1).FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-1],'ITEM MASTER'!C5:C9,5,0),"""")"
    
    SHI_PASTE_PVT.Range("G:G").Insert
    SHI_PASTE_PVT.Range("G2").FormulaR1C1 = "DES 2"
    SHI_PASTE_PVT.Range("G3:G" & LR1).FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-2],'ITEM MASTER'!C5:C12,8,0),"""")"
    
    SHI_PASTE_PVT.Range("L2").FormulaR1C1 = "IT INVT VS MONT VS REK.DET"
    SHI_PASTE_PVT.Range("L3:L" & LR1).FormulaR1C1 = "=IF(RIGHT(RC[-10],5) = ""Total"",IF(AND(RC[-3]=RC[-2],RC[-2]=RC[-1]),TRUE,FALSE),"""")"
    
    On Error Resume Next
    SHI_PASTE_PVT.AutoFilterMode = False
    SHI_PASTE_PVT.Range("A2:L" & LR1).AutoFilter 12, "="
    SHI_PASTE_PVT.Range("L2:L" & LR1).Offset(1).ClearContents
    SHI_PASTE_PVT.AutoFilterMode = False
    SHI_PASTE_PVT.Range("L3:L" & LR1).SpecialCells(xlCellTypeBlanks).FormulaR1C1 = "=R[1]C"
    On Error GoTo 0
    
    SHI_PASTE_PVT.Range("M2").FormulaR1C1 = "ID TO IC"
    SHI_PASTE_PVT.Range("M3:M" & LR1).FormulaR1C1 = "=IF(RC[-10]=""(blank)"","""",TEXTJOIN(""-"",TRUE,RC[-12],RC[-11],RC[-10],RC[-5]))"
    
    SHI_PASTE_PVT.Range("N2").FormulaR1C1 = "CF IC"
    SHI_PASTE_PVT.Range("N3:N" & LR1).FormulaR1C1 = "=IF(RC[-1]="""","""",COUNTIF(R3C13:RC[-1],RC[-1]))"
    
    SHI_PASTE_PVT.Range("O1").FormulaR1C1 = "=SUBTOTAL(9,R[2]C:R[104]C)"
    SHI_PASTE_PVT.Range("O2").FormulaR1C1 = "IC"
    SHI_PASTE_PVT.Range("O3:O" & LR1).FormulaR1C1 = "=IF(RC[-1]="""","""",IF(RC[-1]>1,0,IF(RC[-1]=1,SUMIF('IC2'!C[-5],'PASTE PVT'!RC[-2],'IC2'!C),"""")))"
    
    SHI_PASTE_PVT.Range("O:Q").NumberFormat = "_-* #,##0.00_-;-* #,##0.00_-;_-* ""-""??_-;_-@_-"
    
    FirstRow = 3
    For i = 3 To LR1
        If Right(SHI_PASTE_PVT.Range("B" & i), 5) = "Total" Then
            SHI_PASTE_PVT.Range("O" & i).Formula = "=SUM(O" & FirstRow & ":O" & (i - 1) & ")"
            FirstRow = i + 1
        End If
    Next i
    
    SHI_PASTE_PVT.Range("P1").FormulaR1C1 = "=SUBTOTAL(9,R[2]C:R[104]C)"
    SHI_PASTE_PVT.Range("P2").FormulaR1C1 = "BLNC IC"
    SHI_PASTE_PVT.Range("P3:P" & LR1).FormulaR1C1 = "=IF(RC[-2]="""","""",RC[-1]-RC[-5])"
    
    FirstRow = 3
    For i = 3 To LR1
        If Right(SHI_PASTE_PVT.Range("B" & i), 5) = "Total" Then
            SHI_PASTE_PVT.Range("P" & i).Formula = "=SUM(P" & FirstRow & ":P" & (i - 1) & ")"
            FirstRow = i + 1
        End If
    Next i
    
    SHI_PASTE_PVT.Range("Q1").FormulaR1C1 = "=SUBTOTAL(9,R[2]C:R[104]C)"
    SHI_PASTE_PVT.Range("Q2").FormulaR1C1 = "BLNC KUMULATIF KONKER"
    SHI_PASTE_PVT.Range("Q3:Q" & LR1).FormulaR1C1 = "=IF(RIGHT(RC[-15],5) = ""Total"",RC[-2]-RC[-6],"""")"
    
    On Error Resume Next
    SHI_PASTE_PVT.AutoFilterMode = False
    SHI_PASTE_PVT.Range("A2:Q" & LR1).AutoFilter 17, "="
    SHI_PASTE_PVT.Range("Q2:Q" & LR1).Offset(1).ClearContents
    SHI_PASTE_PVT.AutoFilterMode = False
    On Error GoTo 0
    
    SHI_PASTE_PVT.Range("R2").FormulaR1C1 = "REMARK"
    
    LC1 = SHI_PASTE_PVT.Cells.Find("*", , xlFormulas, xlPart, xlByColumns, xlPrevious).Column
    
    On Error Resume Next
    SHI_PASTE_PVT.AutoFilterMode = False
    SHI_PASTE_PVT.Range("A2:R" & LR1).AutoFilter 2, "*Total"
    SHI_PASTE_PVT.Range("A3:R" & LR1).Interior.Color = RGB(180, 198, 231)
    SHI_PASTE_PVT.AutoFilterMode = False
    On Error GoTo 0
    
    Range("I3").Select
    ActiveWindow.FreezePanes = True
    Cells(1, 1).Select
    
    Cells.EntireColumn.AutoFit
    
    SH1_OLAH.Delete
    SH1_OLAH2.Delete
    
    SH1_HOME.Activate
    Cells(1, 1).Select
    
End Sub

Sub Initial_Sheets()
    Set WB1 = ThisWorkbook
    Set SH1_HOME = WB1.Worksheets("HOME")
    Set SH1_IT_INV = WB1.Worksheets("IT INV")
    Set SH1_MONITORING = WB1.Worksheets("MONITORING")
    Set SH1_REKAP_DETAIL = WB1.Worksheets("REKAP DETAIL")
    Set SH1_RWB = WB1.Worksheets("RWB")
    Set SH1_ITEM_MASTER = WB1.Worksheets("ITEM MASTER")
    Set SH1_TO_PVT = WB1.Worksheets("TO PVT")
    Set SH1_IC = WB1.Worksheets("IC")
    
    Set SH1_IC2 = Sheets.Add(AFTER:=SH1_IC): SH1_IC2.Name = "IC2"
    Set SH1_PVT = Sheets.Add(AFTER:=SH1_IC2): SH1_PVT.Name = "PVT"
    Set SHI_PASTE_PVT = Sheets.Add(AFTER:=SH1_PVT): SHI_PASTE_PVT.Name = "PASTE PVT"
    Set SH1_OLAH = Sheets.Add(AFTER:=SHI_PASTE_PVT): SH1_OLAH.Name = "OLAH"
    Set SH1_OLAH2 = Sheets.Add(AFTER:=SH1_OLAH): SH1_OLAH2.Name = "OLAH2"
    
End Sub

Sub CreatePivotReport()
    SH1_TO_PVT.Activate
    LR1 = SH1_TO_PVT.Range("A" & Rows.Count).End(xlUp).Row
    
    Set pv_Rng = Range("A1:K" & LR1)
    Set pv_Cache = WB1.PivotCaches.Create(SourceType:=xlDatabase, _
                        SourceData:=pv_Rng)
    Set pv_Tb = pv_Cache.CreatePivotTable _
                        (TableDestination:=SH1_PVT.Range("A1"))
    SH1_PVT.Activate
    
    With pv_Tb
        .RowAxisLayout xlTabularRow
        .RowGrand = False
        .NullString = ""
        .ColumnGrand = False
'        .ShowValuesRow = False
        .DisplayErrorString = False
        .RepeatAllLabels xlRepeatLabels
        .PivotCache.MissingItemsLimit = xlMissingItemsDefault
    End With
    
    With pv_Tb.PivotFields("BU")
        .Caption = "BU"
        .Orientation = xlRowField
        .Position = 1
        .Subtotals = _
            Array(False, False, False, False, False, False, False, False, False, False, False, False)
    End With
    
    With pv_Tb.PivotFields("KONKER")
        .Caption = "KONKER"
        .Orientation = xlRowField
        .Position = 2
'        .Subtotals = _
'            Array(False, False, False, False, False, False, False, False, False, False, False, False)
    End With
    
    With pv_Tb.PivotFields("WO")
        .Caption = "WO"
        .Orientation = xlRowField
        .Position = 3
        .Subtotals = _
            Array(False, False, False, False, False, False, False, False, False, False, False, False)
    End With
    
    With pv_Tb.PivotFields("ITEM")
        .Caption = "ITEM"
        .Orientation = xlRowField
        .Position = 4
        .Subtotals = _
            Array(False, False, False, False, False, False, False, False, False, False, False, False)
    End With
    
    With pv_Tb.PivotFields("UOM")
        .Caption = "UOM"
        .Orientation = xlRowField
        .Position = 5
        .Subtotals = _
            Array(False, False, False, False, False, False, False, False, False, False, False, False)
    End With
    
    With pv_Tb.PivotFields("SOURCE1")
        .Orientation = xlColumnField
        .Position = 1
    End With
    
    With pv_Tb.PivotFields("QTY")
        .Orientation = xlDataField
        .Position = 1
        .Function = xlSum
        .NumberFormat = "_(* #,##0.00_);[Red]_(* (#,##0.00);_(* ""-""??_)"
    End With
    
End Sub

Sub CreatePivot_ForGetKonker()
    SH1_TO_PVT.Activate
    
    SH1_TO_PVT.Range("N:O").ClearContents
    
    LR1 = SH1_TO_PVT.Range("A" & Rows.Count).End(xlUp).Row
    Set pv_Rng = Range("I1:J" & LR1)
    Set pv_Cache = WB1.PivotCaches.Create(SourceType:=xlDatabase, _
                        SourceData:=pv_Rng)
    Set pv_Tb = pv_Cache.CreatePivotTable _
                        (TableDestination:=SH1_TO_PVT.Range("N1"))
    
    With pv_Tb
        .RowAxisLayout xlTabularRow
        .RowGrand = False
        .NullString = ""
        .ColumnGrand = False
'        .ShowValuesRow = False
        .DisplayErrorString = False
        .RepeatAllLabels xlRepeatLabels
        .PivotCache.MissingItemsLimit = xlMissingItemsDefault
    End With
    
    With pv_Tb.PivotFields("KONKER")
        .Caption = "KONKER"
        .Orientation = xlRowField
        .Position = 1
        .Subtotals = _
            Array(False, False, False, False, False, False, False, False, False, False, False, False)
    End With
    
    With pv_Tb.PivotFields("WO")
        .Caption = "WO"
        .Orientation = xlRowField
        .Position = 2
        .PivotItems("(blank)").Visible = False
        .Subtotals = _
            Array(False, False, False, False, False, False, False, False, False, False, False, False)
    End With
    
    LR1 = SH1_TO_PVT.Range("N" & Rows.Count).End(xlUp).Row
    With SH1_TO_PVT.Range("N1:O" & LR1)
        .Copy: .PasteSpecial xlPasteValuesAndNumberFormats: Application.CutCopyMode = False
    End With
    
    SH1_TO_PVT.Cells(1, 1).Select

End Sub

Sub PengolahanUntukIC2()
    SH1_IC.Activate
    SH1_IC.AutoFilterMode = False
    LR1 = SH1_IC.Cells.Find("*", , xlFormulas, xlPart, xlByRows, xlPrevious).Row
    LC1 = SH1_IC.Cells.Find("*", , xlFormulas, xlPart, xlByColumns, xlPrevious).Column
    SH1_IC.Range(SH1_IC.Cells(7, 1), SH1_IC.Cells(LR1, LC1)).AutoFilter 10, "IC"
    If WorksheetFunction.CountA(SH1_IC.Range("F:F")) = 1 Then
        End
    End If
    SH1_IC.Range("F7:F" & LR1).Copy
    SH1_IC2.Activate
    SH1_IC2.Range("A1").PasteSpecial xlPasteValuesAndNumberFormats
    
    SH1_IC.Activate
    SH1_IC.Range("H7:H" & LR1).Copy
    SH1_IC2.Activate
    SH1_IC2.Range("B1").PasteSpecial xlPasteValuesAndNumberFormats
    
    LR1_2 = SH1_IC2.Range("A" & Rows.Count).End(xlUp).Row
    SH1_IC2.Range("B1") = "BU"
    SH1_IC2.Range("C1") = "DC"
    SH1_IC2.Range("C2:C" & LR1_2).Value = "IC"
    
    SH1_IC.Activate
    SH1_IC.Range("N7:N" & LR1).Copy
    SH1_IC2.Activate
    SH1_IC2.Range("D1").PasteSpecial xlPasteValuesAndNumberFormats
    SH1_IC2.Range("D1") = "LOT NUM"
    
    SH1_IC.Activate
    SH1_IC.Range("S7:S" & LR1).Copy
    SH1_IC2.Activate
    SH1_IC2.Range("E1").PasteSpecial xlPasteValuesAndNumberFormats
    SH1_IC2.Range("E1") = "QUANTITY"
    
    SH1_IC.Activate
    SH1_IC.Range("X7:X" & LR1).Copy
    
    SH1_IC2.Activate
    SH1_IC2.Range("F1").PasteSpecial xlPasteValuesAndNumberFormats
    SH1_IC2.Range("F1") = "UM"
    
    Application.CutCopyMode = False
    SH1_IC2.Cells.EntireColumn.AutoFit
    Cells(1, 1).Select
'    Stop
    SH1_IC2.AutoFilterMode = False
    LR1 = SH1_IC2.Cells.Find("*", , xlFormulas, xlPart, xlByRows, xlPrevious).Row
    SH1_IC2.Range("G1") = "LEN KK"
    SH1_IC2.Range("G2:G" & LR1).FormulaR1C1 = "=LEN(RC[-6])"
    SH1_IC2.Range("G2:G" & LR1).Copy
    SH1_IC2.Range("G2:G" & LR1).PasteSpecial xlPasteValuesAndNumberFormats: Application.CutCopyMode = False
    SH1_IC2.Range("A1:G" & LR1).AutoFilter 7, "8"
    SH1_IC2.Range("A1:G" & LR1).SpecialCells(xlCellTypeVisible).Copy
    SH1_OLAH.Activate
    SH1_OLAH.Range("A1").PasteSpecial xlPasteValuesAndNumberFormats: Application.CutCopyMode = False
    SH1_IC2.AutoFilterMode = False
    SH1_OLAH.Range("A1:F" & SH1_OLAH.Range("A10000").End(xlUp).Row).Copy
    SH1_IC2.Activate
    SH1_IC2.Range("K1").PasteSpecial xlPasteValuesAndNumberFormats: Application.CutCopyMode = False
    SH1_OLAH.Cells.ClearContents
    
    SH1_IC2.Range("A1:G" & LR1).AutoFilter 7, ">8"
    If WorksheetFunction.CountA(SH1_IC2.Range("A:A")) > 1 Then
        
        SH1_IC2.Range("A1:G" & LR1).SpecialCells(xlCellTypeVisible).Copy
        SH1_OLAH.Activate
        SH1_OLAH.Range("A1").PasteSpecial xlPasteValuesAndNumberFormats: Application.CutCopyMode = False
        SH1_OLAH.Cells.EntireColumn.AutoFit
        SH1_IC2.AutoFilterMode = False
        
        LR1 = SH1_OLAH.Range("A10000").End(xlUp).Row
        
        SH1_OLAH.Range("G1:G" & LR1).ClearContents
        SH1_OLAH.Range("G1").FormulaR1C1 = "KETERANGAN"
        SH1_OLAH.Range("G2:G" & LR1).FormulaR1C1 = "=""1 WO "" & LEN(RC[-6]) - LEN(SUBSTITUTE(RC[-6], ""|"", """")) + 1 & "" KONTRAK"""
        SH1_OLAH.Range("G2:G" & LR1).Copy
        SH1_OLAH.Range("G2:G" & LR1).PasteSpecial xlPasteValuesAndNumberFormats: Application.CutCopyMode = False
        
        SH1_OLAH.Range("A1:A" & LR1).Copy SH1_OLAH.Range("K1")
        SH1_OLAH.Range("K:K").TextToColumns Destination:=SH1_OLAH.Range("K1"), DataType:=xlDelimited, _
        Other:=True, OtherChar:="|"
        
        LC1 = SH1_OLAH.Cells.Find("*", , xlFormulas, xlPart, xlByColumns, xlPrevious).Column
        SumLoop = LC1 - 10
        For i = 1 To SumLoop
            SH1_OLAH.Activate
            SH1_OLAH.Range(SH1_OLAH.Cells(2, i + 10), SH1_OLAH.Cells(LR1, i + 10)).Copy
            SH1_OLAH2.Activate
            If i = 1 Then
                RowPaste = 1
            Else
                RowPaste = SH1_OLAH2.Range("A10000").End(xlUp).Offset(1).Row
            End If
            SH1_OLAH2.Range("A" & RowPaste).PasteSpecial xlPasteAll: Application.CutCopyMode = False
            SH1_OLAH.Range("B2:G" & LR1).Copy
            SH1_OLAH2.Range("B" & RowPaste).PasteSpecial xlPasteAll: Application.CutCopyMode = False
        Next i
        SH1_OLAH2.Activate
        LR1 = SH1_OLAH.Range("A10000").End(xlUp).Row
        For i = LR1 To 1 Step -1
            If Range("A" & i).Value = "" Then
                Rows(i).Delete
            End If
        Next i
        LR1 = SH1_OLAH2.Range("A10000").End(xlUp).Row
        SH1_OLAH2.Range("A1:G" & LR1).Copy
        SH1_IC2.Activate
        SH1_IC2.Range("K10000").End(xlUp).Offset(1).PasteSpecial xlPasteValuesAndNumberFormats
        Application.CutCopyMode = False
    End If
    SH1_IC2.Activate
    LR1 = SH1_IC2.Range("K10000").End(xlUp).Row
    SH1_IC2.Range("J1").FormulaR1C1 = "ID"
    SH1_IC2.Range("J2:J" & LR1).FormulaR1C1 = "=TEXTJOIN(""-"",TRUE,RC[2],RC[1],RC[4],RC[6])"
    SH1_IC2.Cells.EntireColumn.AutoFit
    SH1_IC2.Cells(1, 1).Select
End Sub

Sub GetKonker_ToIC()

    SH1_IC.Activate
    SH1_IC.AutoFilterMode = False
    SH1_IC.Range("F:F").ClearContents
    SH1_IC.Range("F7") = "KK"
    ' Tentukan baris terakhir di Sheet1 dan Sheet2
    
    LR1 = SH1_IC.Cells(SH1_IC.Rows.Count, "N").End(xlUp).Row
    LR1_2 = SH1_TO_PVT.Cells(SH1_TO_PVT.Rows.Count, "N").End(xlUp).Row

    ' Loop melalui setiap nilai di kolom N pada Sheet1
    For Each cell In SH1_IC.Range("N8:N" & LR1)
        If cell.Offset(0, -4) = "IC" Then
            lookupValue = cell.Value
            result = "" ' Reset result untuk setiap lookup value
    
            ' Loop melalui setiap cell di Sheet2 untuk mencari nilai yang cocok
            For Each rng In SH1_TO_PVT.Range("O2:O" & LR1_2)
                ' Jika nilai cocok, tambahkan ke hasil dengan pemisah "|"
                If rng.Value <> "" Then
                    If rng.Value = lookupValue Then
                        If result = "" Then
                            result = rng.Offset(0, -1).Value ' Ambil nilai di kolom sebelah kanan dari hasil yang cocok
                        Else
                            If rng.Offset(0, -1).Value <> Left(Results, 8) Then
                                result = result & " | " & rng.Offset(0, -1).Value
                            End If
                        End If
                    End If
                End If
            Next rng
    
            ' Simpan hasil di kolom F pada baris yang sama di Sheet1
            cell.Offset(0, -8).Value = result ' Sesuaikan Offset sesuai dengan posisi kolom F
        End If
    Next cell
End Sub













































