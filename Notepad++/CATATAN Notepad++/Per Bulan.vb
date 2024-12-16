Sub ReportPerBulan()
Application.DisplayAlerts = False

TES1.Activate
If TES1.AutoFilterMode = True Then TES1.AutoFilterMode = False
Cells.ClearContents: Cells(1, 1).Select

shOlahan.Activate
LR = Cells.Find(What:="*", LookAt:=xlPart, LookIn:=xlFormulas, SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
LC = Cells.Find(What:="*", LookAt:=xlPart, LookIn:=xlFormulas, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column

Range(Cells(1, 21), Cells(LR, 21)).Copy TES1.Range("A1") 'BULAN
Range(Cells(1, 2), Cells(LR, 2)).Copy TES1.Range("B1") 'FACTORY
Range(Cells(1, 22), Cells(LR, 22)).Copy TES1.Range("C1") 'PERIODE BULAN

TES1.Activate
Range("A:A").RemoveDuplicates 1, xlYes
Range("B:B").RemoveDuplicates 1, xlYes
Range("C:C").RemoveDuplicates 1, xlYes

sumBulan = TES1.Range("A" & Rows.Count).End(xlUp).Row - 1
Range("B1:B" & Range("B" & Rows.Count).End(xlUp).Row).Copy TES2.Range("A1")

'[*] Siapkan Kolom Factory Untuk Hasilnya
'[*] Setelah selesai pindahkan ke sheets TEMP2
'...........................................................................
Range("B1:B" & Range("B" & Rows.Count).End(xlUp).Row).Copy TEMP1.Range("A1")
TEMP1.Activate
Range("A" & Range("A" & Rows.Count).End(xlUp).Row + 1).Value = "Grand Total"

Rows(1).Insert
Range("B:G").HorizontalAlignment = xlCenter
Range("A:G").VerticalAlignment = xlCenter


'[*] OLAH DATANYA
'...........................................................................

TES2.Activate
Range("B:B,D:F").NumberFormat = _
    "_($* #,##0.00_);[Red]_($* (#,##0.00);_($* ""-""??_);_(@_)"
Range("D:F").NumberFormat = _
    "_($* #,##0.00_);[Red]_($* (#,##0.00);_($* ""-""??_);_(@_)"
Range("C:C").NumberFormat = "#,###"
Range("G:G").NumberFormat = "0.0%"

'#### JANGAN LUPA GANTI
'sumBulan = 1
'#### JANGAN LUPA GANTI

For i = 1 To sumBulan
    strBulan = TES1.Range("A" & i + 1).Value
    periodeBulan = TES1.Range("C" & i + 1).Value
    
    shReportBulan.Activate
    Set rgPeriode = Cells.Find(What:=periodeBulan, After:=ActiveCell, LookIn:=xlValues, _
        LookAt:=xlWhole, SearchOrder:=xlByRows, SearchDirection:=xlNext, _
        MatchCase:=False, SearchFormat:=False)
    
    cekData = Application.WorksheetFunction.CountA(Range("B:B"))
    
    If rgPeriode Is Nothing Then
    
        If cekData = 0 Then
            rowPaste = 1
            TES2.Activate
            LR = Cells.Find(What:="*", LookAt:=xlPart, LookIn:=xlFormulas, SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
        Else
            TES2.Activate
            LR = Cells.Find(What:="*", LookAt:=xlPart, LookIn:=xlFormulas, SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
            If Range("A" & LR).Value = "Grand Total" Then
                Rows(LR).Delete
                LR = LR - 1
            End If
            TEMP2.Activate
            cekData = Application.WorksheetFunction.CountA(Cells)
            If cekData > 0 Then
                rowPaste = Cells.Find(What:="*", LookAt:=xlPart, LookIn:=xlFormulas, SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row + 2
            Else
                rowPaste = 1
            End If
        
        End If
        TEMP1.Activate
        Cells.Borders.LineStyle = xlNone
        Range("A1").ClearContents
        Range("B:G").ClearContents
        
        TES2.Activate
        Range("B1:G" & LR).ClearContents
        
        Range("B1").FormulaR1C1 = "PLAN INCOME"
        Range("B2:B" & LR).FormulaR1C1 = _
            "=SUMIFS(KAPASITAS!C[16],KAPASITAS!C[2],'TES2'!RC[-1],KAPASITAS!C[19],""" & strBulan & """)"
        
        Range("C1").FormulaR1C1 = "Output (Pcs)"
        Range("C2:C" & LR).FormulaR1C1 = _
            "=SUMIFS(OLAHAN!C[8],OLAHAN!C[-1],'TES2'!RC[-2],OLAHAN!C[18],""" & strBulan & """)"
        
        Range("D1").FormulaR1C1 = "Income ($)"
        Range("D2:D" & LR).FormulaR1C1 = _
            "=SUMIFS(OLAHAN!C[9],OLAHAN!C[-2],'TES2'!RC[-3],OLAHAN!C[17],""" & strBulan & """)"
        
        Range("E1").FormulaR1C1 = "Cost ($)"
        Range("E2:E" & LR).FormulaR1C1 = _
            "=SUMIFS(OLAHAN!C[11],OLAHAN!C[-3],'TES2'!RC[-4],OLAHAN!C[16],""" & strBulan & """)"
        
        Range("F1").FormulaR1C1 = "Profit ($)"
        Range("F2:F" & LR).FormulaR1C1 = "=RC[-1]-RC[-2]"
        
        Range("G1").FormulaR1C1 = "REALISASI"
        Range("G2:G" & LR).FormulaR1C1 = "=RC[-3]/RC[-5]-1"
        
        Range("A" & LR + 1).Value = "Grand Total"
        Range(Cells(LR + 1, 2), Cells(LR + 1, 6)).FormulaR1C1 = "=SUM(R[-" & LR - 1 & "]C:R[-1]C)"
        Range("G" & LR + 1).FormulaR1C1 = "=RC[-3]/RC[-5]-1"
        
        Cells.EntireColumn.AutoFit
        With Range("B1:G" & LR + 1)
            .Copy
            .PasteSpecial xlPasteValuesAndNumberFormats
        End With
        Application.CutCopyMode = False: Cells(1, 1).Select
        
        Set rg = Range("B1:G" & LR + 1)
        rg.Copy
        
        TEMP1.Activate
        Cells(1, 1).Value = periodeBulan
        Cells(2, 2).PasteSpecial xlPasteAll: Application.CutCopyMode = False
        
        LR = Cells.Find(What:="*", LookAt:=xlPart, LookIn:=xlFormulas, SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
        LC = Cells.Find(What:="*", LookAt:=xlPart, LookIn:=xlFormulas, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column
        
        Range("A:A").Font.Bold = True
        Rows(2).Font.Bold = True
        Rows(LR).Font.Bold = True
        
        With Range(Cells(2, 1), Cells(LR, LC))
            .Borders.LineStyle = xlContinuous
        End With
        
        Cells.EntireColumn.AutoFit
        Cells(1, 1).Select
        
        Set rg = Range(Cells(1, 1), Cells(LR, LC))
        TEMP2.Activate
        
        rg.Copy
        Range("A" & rowPaste).PasteSpecial xlPasteAll
        Application.CutCopyMode = False
        Cells.EntireColumn.AutoFit: Cells(1, 1).Select
        
    End If
    
Next i

shReportBulan.Activate
cekData = Application.WorksheetFunction.CountA(Range("B:B"))

If cekData = 0 Then
    
    TEMP2.Activate
    Rows("1:3").Insert
    strTitle = "Resume Profit & Loss Th " & shBantuan.Range("A2").Value
    LC = Cells.Find(What:="*", LookAt:=xlPart, LookIn:=xlFormulas, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column
    
    With Range(Cells(2, 1), Cells(2, LC))
        .Merge
        .Value = strTitle
        .Font.Size = 18
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
    rowPaste = 1

Else
    shReportBulan.Activate
    rowPaste = Cells.Find(What:="*", LookAt:=xlPart, LookIn:=xlFormulas, SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row + 2

End If
TEMP2.Activate
Range("A:A").Insert: Cells(1, 1).Select

LR = Cells.Find(What:="*", LookAt:=xlPart, LookIn:=xlFormulas, SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
LC = Cells.Find(What:="*", LookAt:=xlPart, LookIn:=xlFormulas, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column

Set rg = Range(Cells(1, 1), Cells(LR, LC))
rg.Copy

shReportBulan.Activate
Range("A" & rowPaste).PasteSpecial xlPasteAll
Application.CutCopyMode = False
Cells.EntireColumn.AutoFit: Cells(1, 1).Select

TES2.Cells.Clear
TEMP1.Cells.Clear
TEMP2.Cells.Clear

Application.DisplayAlerts = True
End Sub
