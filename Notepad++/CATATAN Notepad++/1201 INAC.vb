
Public WB1 As Workbook
Public SH1_Home As Worksheet, SH1_RPA1 As Worksheet, SH As Worksheet, SH1_SetRef As Worksheet
Public SH1_ILDV As Worksheet, SH1_SHPMNT_REPORT As Worksheet, SH1_SALES_LOKAL As Worksheet, SH1_WO_BUYER As Worksheet, SH1_CC_INV_AGING As Worksheet, SH1_INV_AGING As Worksheet
Public SH1_OLAH1 As Worksheet, SH1_OLAH2 As Worksheet, SH1_OLAH3 As Worksheet, SH1_OLAH4 As Worksheet, SH1_OLAH5 As Worksheet
Public STR_BRANCH As String, STR_NAMA_SHEET As String

Public WB2 As Workbook

Public i As Long
Public LR1 As Long, LR2 As Long
Public LC1 As Long, LC2 As Long
Public RNG As Range

Public InitPath As String
Public PathAging As String

Public Int_BulanSekarang As Long
Public Int_NoBulan As Long
Public Int_Tahun As Long
Public Int_2NumberYear As Long

Public Str_NoBulan As String, Str_NamaBulan As String
Public StrBranch As String

Public pv_Rng As Range
Public pv_Cache As PivotCache
Public pv_Tb As PivotTable

Sub PROSES1()
    Application.DisplayAlerts = False
    Application.AskToUpdateLinks = False
    
    Call HideSeetRef
    Call Data_Validation
    Call DeleteSheetsExcept("Set_Ref", "HOME", "STATUS", "WA", "RPA1", "RPA2")
    Call Add_Sheets_Preprocessing("CC_ILDV", "CC_SALES_LOKAL", "CC_SHIPMENT_REPORT", "CC_WO_BUYER", "CC_INVENTORY_AGING", "INVENTORY_AGING", _
                                  "OLAH1", "OLAH2", "OLAH3", "OLAH4", "OLAH5")
    Call Init_Proses
    Call RemoveHyperlink
    Call ImportFile_TarikanJDE
    Call Olah_Data
    Call CreateSheetAndDataEmail
    
    Application.DisplayAlerts = True
    Application.DisplayAlerts = True
End Sub

Sub Olah_Data()

''[KE SHEET INVENTORY AGING -> AMBIL DATA YG AMOUNT>0, TAHUN BERJALAN]
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    If Month(Date) - 1 = 12 Then
        Int_Tahun = Year(Date) - 1
    Else
        Int_Tahun = Year(Date)
    End If
    
    SH1_CC_INV_AGING.Activate
    SH1_CC_INV_AGING.AutoFilterMode = False
    SH1_CC_INV_AGING.Cells.EntireColumn.Hidden = False
    LR1 = SH1_CC_INV_AGING.Range("A" & Rows.Count).End(xlUp).Row
    SH1_CC_INV_AGING.Range("A8:U" & LR1).AutoFilter 10, "<>0"
    SH1_CC_INV_AGING.Range("A8:U" & LR1).AutoFilter 14, Int_Tahun
    
    SH1_CC_INV_AGING.Range("A:U").SpecialCells(xlCellTypeVisible).Copy
    
    SH1_INV_AGING.Activate
    SH1_INV_AGING.Range("A1").PasteSpecial xlPasteValuesAndNumberFormats: Application.CutCopyMode = False
    Cells(1, 1).Select
    
''[DATA CLEANING]............................................................................................
' ------------------------------------------------------------------------------------------------------------
    SH1_ILDV.Activate
    SH1_ILDV.Cells.Copy
    
    SH1_OLAH1.Activate
    SH1_OLAH1.Range("A1").PasteSpecial xlPasteAll: Application.CutCopyMode = False
    SH1_OLAH1.Rows("1:5").Delete
    SH1_OLAH1.Range("A1") = "ITEM NUMBER"
    SH1_OLAH1.Range("H1") = "BUSINESS UNIT"
    SH1_OLAH1.Range("J1") = "DC. TY."
    SH1_OLAH1.Range("V1") = "REF.NUMBR"
    SH1_OLAH1.Rows(2).Delete
    SH1_OLAH1.Range("C:D").Delete
    SH1_OLAH1.Range("D:E").Delete
    SH1_OLAH1.Range("L:L").Delete
    SH1_OLAH1.Range("O:O").Delete
    SH1_OLAH1.Range("Q:Q").Delete
    
    LR1 = SH1_OLAH1.Cells.Find("*", , xlFormulas, xlPart, xlByRows, xlPrevious).Row
    LC1 = SH1_OLAH1.Cells.Find("*", , xlFormulas, xlPart, xlByColumns, xlPrevious).Column
    Set RNG = SH1_OLAH1.Range(SH1_OLAH1.Cells(1, 1), SH1_OLAH1.Cells(LR1, LC1))
    
    RNG.AutoFilter 13, "TOTAL"
    RNG.Offset(1).Resize(RNG.Rows.Count - 1, RNG.Columns.Count).SpecialCells(xlCellTypeVisible).Delete
    SH1_OLAH1.ShowAllData
    RNG.AutoFilter 6, "="
    RNG.Offset(1).Resize(RNG.Rows.Count - 1, RNG.Columns.Count).SpecialCells(xlCellTypeVisible).Delete
    SH1_OLAH1.AutoFilterMode = False
    
    LR1 = SH1_OLAH1.Cells.Find("*", , xlFormulas, xlPart, xlByRows, xlPrevious).Row
    
'    Range("I:I").Insert
'    Range("I1").FormulaR1C1 = "TAHUN"
'    With Range("I2:I" & LR1)
'        .FormulaR1C1 = "=TEXT(RC[-1],""YYYY"")"
'        .Copy: .PasteSpecial xlPasteValuesAndNumberFormats: Application.CutCopyMode = False
'    End With
'    Range("SAM1").Copy
'    Range("I:I").PasteSpecial xlPasteAll, xlPasteSpecialOperationAdd: Application.CutCopyMode = False
'
'    RNG.AutoFilter 9, "<>" & Int_Tahun
'    RNG.Offset(1).Resize(RNG.Rows.Count - 1, RNG.Columns.Count).SpecialCells(xlCellTypeVisible).Delete
'    Range("I:I").Delete
    
'    '-> [PILIH DATA YANG AMOUNT AGINGNYA > 0]
'    LR1 = SH1_OLAH1.Cells.Find("*", , xlFormulas, xlPart, xlByRows, xlPrevious).Row
'    SH1_OLAH1.Range("S1") = "AMOUNT"
'    With Range("S2:S" & LR1)
'        .FormulaR1C1 = _
'        "=IFERROR(VLOOKUP(RC[-18],INVENTORY_AGING!C1:C19,19,FALSE),"""")"
'        .Copy: .PasteSpecial xlPasteValuesAndNumberFormats: Application.CutCopyMode = False
'    End With
'    LC1 = SH1_OLAH1.Cells.Find("*", , xlFormulas, xlPart, xlByColumns, xlPrevious).Column
'    Set RNG = SH1_OLAH1.Range(SH1_OLAH1.Cells(1, 1), SH1_OLAH1.Cells(LR1, LC1))
'    RNG.AutoFilter 19, "="
'    RNG.Offset(1).Resize(RNG.Rows.Count - 1, RNG.Columns.Count).SpecialCells(xlCellTypeVisible).Delete
'
'    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

    SH1_OLAH1.AutoFilterMode = False
    Cells.EntireColumn.AutoFit
    Cells(1, 1).Select
    
    LR1 = SH1_OLAH1.Cells.Find("*", , xlFormulas, xlPart, xlByRows, xlPrevious).Row
    LC1 = SH1_OLAH1.Cells.Find("*", , xlFormulas, xlPart, xlByColumns, xlPrevious).Column
' -----------------------------------------------------------------------------------------------------------
    
''[DAPATKAN OR]..............................................................................................
' -----------------------------------------------------------------------------------------------------------
    SH1_OLAH1.Range("K:K").Insert
    SH1_OLAH1.Range("K1") = "LEN"
    SH1_OLAH1.Range("K2:K" & LR1).FormulaR1C1 = "=LEN(RC[-1])"
    SH1_OLAH1.Range("L:L").Insert
    SH1_OLAH1.Range("L1") = "OR"
    SH1_OLAH1.Range("S1") = "UOM"
    SH1_OLAH1.Range("L2:L" & LR1).FormulaR1C1 = "=IF(RC[-1]>7,RIGHT(RC[-2],8),"""")"
    With SH1_OLAH1.Range("K:L")
        .Copy: .PasteSpecial xlPasteValuesAndNumberFormats: Application.CutCopyMode = False
    End With
    SH1_OLAH1.Range("SAM1").Copy
    Range("L:L").PasteSpecial xlPasteAll, xlPasteSpecialOperationAdd
    Application.CutCopyMode = False
' -----------------------------------------------------------------------------------------------------------

''[DAPATKAN BUYER]...........................................................................................
' -----------------------------------------------------------------------------------------------------------
    SH1_ILDV.Activate
    LR1 = SH1_ILDV.Cells.Find("*", , xlFormulas, xlPart, xlByRows, xlPrevious).Row
    Range("K:K").Insert
    Range("K7").FormulaR1C1 = "[IM] -> BUYER"
    With Range("K8:K" & LR1)
        .FormulaR1C1 = _
            "=IFERROR(IF(AND(RC[-1]<>"""",RC[-1]=""IM""),VLOOKUP(RC[-2],CC_WO_BUYER!C3:C13,11,FALSE),""""),"""")"
        .Copy: .PasteSpecial xlPasteValuesAndNumberFormats: Application.CutCopyMode = False
    End With
    SH1_ILDV.AutoFilterMode = False
    SH1_ILDV.Rows("1:" & LR1).AutoFilter 11, "<>"
    SH1_ILDV.Cells.SpecialCells(xlCellTypeVisible).Copy
    SH1_OLAH3.Activate
    Range("A1").PasteSpecial xlPasteAll: Application.CutCopyMode = False
    Cells.EntireColumn.AutoFit
    
    SH1_OLAH1.Activate
'    LR1 = SH1_OLAH1.Cells.Find("*", , xlFormulas, xlPart, xlByRows, xlPrevious).Row
    LR1 = SH1_OLAH1.Range("A" & Rows.Count).End(xlUp).Row
    LC1 = SH1_OLAH1.Cells.Find("*", , xlFormulas, xlPart, xlByColumns, xlPrevious).Column
    
    SH1_OLAH1.Range("M:M").Insert
    SH1_OLAH1.Range("M1") = "OPSI1"
    SH1_OLAH1.Range("M2:M" & LR1).FormulaR1C1 = _
        "=IFERROR(IF(RC[-1]<>"""",VLOOKUP(RC[-1],CC_WO_BUYER!C4:C13,10,FALSE),""""),"""")"
    
    SH1_OLAH1.Range("N:N").Insert
    SH1_OLAH1.Range("N1") = "OPSI2"
    SH1_OLAH1.Range("N2:N" & LR1).FormulaR1C1 = _
        "=IFERROR(IF(RC[-1]="""",INDEX(CC_SHIPMENT_REPORT!C3,MATCH(OLAH1!RC[-2],CC_SHIPMENT_REPORT!C23,0)),""""),"""")"
    
    SH1_OLAH1.Range("O:O").Insert
    SH1_OLAH1.Range("O1") = "OPSI3"
    SH1_OLAH1.Range("O2:O" & LR1).FormulaR1C1 = _
        "=IFERROR(IF(AND(RC[-2]="""",RC[-1]="""",RC[-9]=""IM""),VLOOKUP(RC[-10],CC_WO_BUYER!C3:C13,11,FALSE),""""),"""")"
    
    SH1_OLAH1.Range("P:P").Insert
    SH1_OLAH1.Range("P1") = "BUYER TERISI V1"
    SH1_OLAH1.Range("P2:P" & LR1).FormulaR1C1 = "=CONCATENATE(RC[-3],RC[-2],RC[-1])"
        
    With SH1_OLAH1.Range("M1:P" & LR1)
        .Copy: .PasteSpecial xlPasteValuesAndNumberFormats: Application.CutCopyMode = False
    End With
    
''[SIMPAN YANG SUDAH ADA BUYERNYA, UNTUK DI LOOKUP BY ITEM YG BLM ADA BUYERNYA]
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    SH1_OLAH1.Cells.AutoFilter 16, "<>"
    SH1_OLAH1.Range("A1:P" & LR1).SpecialCells(xlCellTypeVisible).Copy
    
    SH1_OLAH2.Activate
    SH1_OLAH2.Range("A1").PasteSpecial xlPasteValuesAndNumberFormats: Application.CutCopyMode = False
    
    SH1_OLAH1.Activate
    SH1_OLAH1.AutoFilterMode = False
    
    SH1_OLAH1.Range("Q:Q").Insert
    SH1_OLAH1.Range("Q1") = "OPSI4"
    SH1_OLAH1.Range("Q2:Q" & LR1).FormulaR1C1 = _
        "=IFERROR(IF(RC[-1]<>"""",RC[-1],VLOOKUP(RC[-16],OLAH2!C1:C16,16,FALSE)),"""")"
        
    SH1_OLAH1.Range("R:R").Insert
    SH1_OLAH1.Range("R1") = "OPSI5"
    SH1_OLAH1.Range("R2:R" & LR1).FormulaR1C1 = _
        "=IFERROR(IF(RC[-1]<>"""",RC[-1],VLOOKUP(RC[-17],OLAH3!C1:C11,11,FALSE)),"""")"
    
    SH1_OLAH1.Range("S:S").Insert
    SH1_OLAH1.Range("S1") = "OPSI6"
    SH1_OLAH1.Range("S2:S" & LR1).FormulaR1C1 = _
        "=IFERROR(IF(RC[-1]<>"""",RC[-1],INDEX(CC_SALES_LOKAL!C2,MATCH(OLAH1!RC[-7],CC_SALES_LOKAL!C9,0))),"""")"

    With SH1_OLAH1.Range("Q2:S" & LR1)
        .Copy: .PasteSpecial xlPasteValuesAndNumberFormats: Application.CutCopyMode = False
    End With
    
    SH1_OLAH1.Range("M:R").Delete
    SH1_OLAH1.Range("M1") = "BUYER"
'    Stop
    
    SH1_OLAH1.AutoFilterMode = False
    SH1_OLAH1.Cells.AutoFilter 6, "OV"
    
    'ITEM
    SH1_OLAH1.Activate
    SH1_OLAH1.Range("A1:A" & LR1).Copy
    SH1_OLAH5.Activate
    SH1_OLAH5.Range("A1").PasteSpecial xlPasteValuesAndNumberFormats
    
    'DC.TY
    SH1_OLAH1.Activate
    SH1_OLAH1.Range("F1:F" & LR1).Copy
    SH1_OLAH5.Activate
    SH1_OLAH5.Range("B1").PasteSpecial xlPasteValuesAndNumberFormats
    
    'G/L DATE
    SH1_OLAH1.Activate
    SH1_OLAH1.Range("H1:H" & LR1).Copy
    SH1_OLAH5.Activate
    SH1_OLAH5.Range("C1").PasteSpecial xlPasteValuesAndNumberFormats
    
    Application.CutCopyMode = False
    SH1_OLAH1.AutoFilterMode = False
    
    SH1_OLAH5.Activate
    LR1 = SH1_OLAH5.Cells.Find("*", , xlFormulas, xlPart, xlByRows, xlPrevious).Row
    SH1_OLAH5.Sort.SortFields.Clear
    SH1_OLAH5.Sort.SortFields.Add2 Key:=Range("C2:C" & LR1), SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
    With SH1_OLAH5.Sort
        .SetRange Range("A1:C" & LR1)
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
    SH1_OLAH5.Cells.EntireColumn.AutoFit
    SH1_OLAH5.Cells(1, 1).Select
    
'    Stop
    SH1_OLAH1.Activate
    LR1 = SH1_OLAH1.Cells.Find("*", , xlFormulas, xlPart, xlByRows, xlPrevious).Row
    SH1_OLAH1.Range("V1").value = "MONTH"
    SH1_OLAH1.Range("V2:V" & LR1).FormulaR1C1 = "=TEXT(VLOOKUP(RC[-21],OLAH5!C1:C3,3,FALSE),""MMMM"")"
    
    SH1_OLAH1.Range("W1").value = "YEAR"
    SH1_OLAH1.Range("W2:W" & LR1).FormulaR1C1 = "=TEXT(VLOOKUP(RC[-22],OLAH5!C1:C3,3,FALSE),""YYYY"")"
    
    With SH1_OLAH1.Range("V2:W" & LR1)
        .Copy: .PasteSpecial xlPasteValuesAndNumberFormats: Application.CutCopyMode = False
    End With
    
    Range("SAM1").Copy
    Range("W2:W" & LR1).PasteSpecial xlPasteAll, xlPasteSpecialOperationAdd: Application.CutCopyMode = False
    
    SH1_OLAH1.Cells.EntireColumn.AutoFit
    SH1_OLAH1.Cells(1, 1).Select
    
    
''[HAPUS DATA DI OLAH 2]
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    SH1_OLAH2.Activate
    SH1_OLAH2.Cells.ClearContents
    SH1_OLAH2.Cells(1, 1).Select
    
''[BUAT PIVOT]
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Call Create_Pivot_1
    
''[OLAH DATA, HANYA ITEM YG ADA DI INV AGING]
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    SH1_OLAH3.Cells.ClearContents
    SH1_OLAH2.Activate
    Cells.Copy
    SH1_OLAH3.Activate
    SH1_OLAH3.Range("A1").PasteSpecial xlPasteValuesAndNumberFormats: Application.CutCopyMode = False
    SH1_OLAH3.Rows(1).Delete
    LR1 = SH1_OLAH3.Range("A" & Rows.Count).End(xlUp).Row
    
    '--->[HANYA ITEM YG ADA DI INV AGING]
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'    SH1_OLAH3.AutoFilterMode = False
'    SH1_OLAH3.Range("J1") = "FOUND"
'    With SH1_OLAH3.Range("J2:J" & LR1)
'        .FormulaR1C1 = _
'        "=IFERROR(VLOOKUP(RC[-9],INVENTORY_AGING!C1,1,FALSE),""NOT FOUND"")"
'        .Copy: .PasteSpecial xlPasteValuesAndNumberFormats: Application.CutCopyMode = False
'    End With
'    LC1 = SH1_OLAH3.Cells.Find("*", , xlFormulas, xlPart, xlByColumns, xlPrevious).Column
'    Set RNG = SH1_OLAH3.Range(SH1_OLAH3.Cells(1, 1), SH1_OLAH3.Cells(LR1, LC1))
'    RNG.AutoFilter 10, "<>NOT FOUND"
'    RNG.SpecialCells(xlCellTypeVisible).Copy
'    SH1_OLAH4.Activate
'    SH1_OLAH4.Range("A1").PasteSpecial xlPasteAll: Application.CutCopyMode = False
'    SH1_OLAH3.Activate
'    SH1_OLAH3.AutoFilterMode = False
'    SH1_OLAH3.Cells.ClearContents
'    SH1_OLAH4.Cells.Cut SH1_OLAH3.Range("A1")
'    Range("J:J").ClearContents

    '--->[HANYA YG MEMPUNYAI AMOUNT DI INV AGING DAN BUKAN BLANK HASIL PIVOT]
    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    SH1_OLAH3.AutoFilterMode = False
    SH1_OLAH3.Range("J1") = "AMOUNT"
    With SH1_OLAH3.Range("J2:J" & LR1)
        .FormulaR1C1 = _
        "=IFERROR(VLOOKUP(RC[-9],INVENTORY_AGING!C1:C19,19,FALSE),"""")"
        .Copy: .PasteSpecial xlPasteValuesAndNumberFormats: Application.CutCopyMode = False
    End With
    LC1 = SH1_OLAH3.Cells.Find("*", , xlFormulas, xlPart, xlByColumns, xlPrevious).Column
    Set RNG = SH1_OLAH3.Range(SH1_OLAH3.Cells(1, 1), SH1_OLAH3.Cells(LR1, LC1))
    RNG.Replace "(blank)", "", xlWhole
    
'    RNG.AutoFilter 4, "<>"
'    RNG.AutoFilter 5, "<>"
'    RNG.AutoFilter 7, "<>"
'    RNG.AutoFilter 10, "<>"
    
    '[## LAMUN AMOUNT LEDGER NA 0 / BLANK, TONG DIBAWA]
'    RNG.AutoFilter 9, "<>"
'
'    RNG.SpecialCells(xlCellTypeVisible).Copy
    
    RNG.Copy
    
    SH1_OLAH4.Activate
    SH1_OLAH4.Range("A1").PasteSpecial xlPasteAll: Application.CutCopyMode = False
    SH1_OLAH3.Activate
    SH1_OLAH3.AutoFilterMode = False
    SH1_OLAH3.Cells.ClearContents
    
    SH1_OLAH4.Cells.Cut SH1_OLAH3.Range("A1")
    
    SH1_OLAH3.Range("J:J").Insert
    
    LR1 = SH1_OLAH3.Range("A" & Rows.Count).End(xlUp).Row
    
    SH1_OLAH3.Range("J1") = "QTY"
    SH1_OLAH3.Range("J2:J" & LR1).FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-9],INVENTORY_AGING!C1:C9,9,FALSE),"""")"
    
'    SH1_OLAH3.Range("L1") = "YEAR"
'    SH1_OLAH3.Range("L2:L" & LR1).FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-11],INVENTORY_AGING!C1:C14,14,FALSE),"""")"
'
'    SH1_OLAH3.Range("M1") = "MONTH"
'    SH1_OLAH3.Range("M2:M" & LR1).FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-12],INVENTORY_AGING!C1:C15,15,FALSE),"""")"

    SH1_OLAH3.Range("L1") = "MONTH"
    SH1_OLAH3.Range("L2:L" & LR1).FormulaR1C1 = "=IFERROR(UPPER(TEXT(VLOOKUP(RC[-11],OLAH5!C1:C3,3,FALSE),""MMMM"")),"""")"

    SH1_OLAH3.Range("M1") = "YEAR"
    SH1_OLAH3.Range("M2:M" & LR1).FormulaR1C1 = "=IFERROR(UPPER(TEXT(VLOOKUP(RC[-12],OLAH5!C1:C3,3,FALSE),""YYYY"")),"""")"

    SH1_OLAH3.Range("N1").FormulaR1C1 = "C-IF"
    SH1_OLAH3.Range("N2:N" & LR1).FormulaR1C1 = "=COUNTIF(R2C1:RC[-13],RC[-13])"
    
    With SH1_OLAH3.Range("J1:N" & LR1)
        .Copy: .PasteSpecial xlPasteValuesAndNumberFormats: Application.CutCopyMode = False
    End With
    
    SH1_OLAH3.Range("SAM1").Copy
    SH1_OLAH3.Range("M2:N" & LR1).PasteSpecial xlPasteAll, xlPasteSpecialOperationAdd: Application.CutCopyMode = False
    
    
    For i = LR1 To 2 Step -1
        If SH1_OLAH3.Range("N" & i) <> "1" Then
            SH1_OLAH3.Range("J" & i & ":K" & i).ClearContents
        End If
    Next i
    SH1_OLAH3.Range("N:N").Delete
    
    Range("O1").FormulaR1C1 = "PK1"
    Range("O2:O" & LR1).FormulaR1C1 = "=RC[-8]&""-""&RC[-14]&""-""&RC[-13]&""-""&RC[-12]&""-""&RC[-9]"
    
    Range("P1").FormulaR1C1 = "SERI"
    Range("P2:P" & LR1).FormulaR1C1 = "=COUNTIF(R1C15:RC[-1],RC[-1])"
    
    Range("Q1").FormulaR1C1 = "[SUMIF] AMT BY PK1"
    Range("Q2:Q" & LR1).FormulaR1C1 = "=SUMIF(R2C15:R" & LR1 & "C15,RC[-2],R2C9:R" & LR1 & "C9)"
    
    Range("R1").FormulaR1C1 = "[SUMIF] QTY BY PK1"
    Range("R2:R" & LR1).FormulaR1C1 = "=SUMIF(R2C15:R" & LR1 & "C15,RC[-3],R2C10:R" & LR1 & "C10)"
    
    Rows("1:2").Insert
    LR1 = SH1_OLAH3.Range("A" & Rows.Count).End(xlUp).Row
    
    Range("H2").FormulaR1C1 = "=SUBTOTAL(9,R[2]C:R[" & LR1 - 2 & "]C)"
    Range("I2").FormulaR1C1 = "=SUBTOTAL(9,R[2]C:R[" & LR1 - 2 & "]C)"
    Range("J2").FormulaR1C1 = "=SUBTOTAL(9,R[2]C:R[" & LR1 - 2 & "]C)"
    Range("K2").FormulaR1C1 = "=SUBTOTAL(9,R[2]C:R[" & LR1 - 2 & "]C)"
    Range("H2:K2").NumberFormat = "_(* #,##0.00_);_(* (#,##0.00);_(* ""-""??_);_(@_)"
    
    Range("L2").value = "LAST RECEIVED(OV)"
    
    SH1_OLAH3.Cells.EntireColumn.AutoFit: SH1_OLAH3.Cells(1, 1).Select
    
    SH1_OLAH3.AutoFilterMode = False
    LR1 = SH1_OLAH3.Range("A" & Rows.Count).End(xlUp).Row
    
    SH1_OLAH3.Range("A3:R" & LR1).AutoFilter Field:=16, Criteria1:="1"
    SH1_OLAH3.Range("A3:R" & LR1).AutoFilter Field:=17, Criteria1:="<>0", Operator:=xlAnd
    
    SH1_OLAH3.Range("A1:R" & LR1).Copy
    
    SH1_OLAH4.Activate
    SH1_OLAH4.Range("A1").PasteSpecial xlPasteValuesAndNumberFormats: Application.CutCopyMode = False
    SH1_OLAH3.AutoFilterMode = False
    SH1_OLAH4.Rows("1:2").Delete
    SH1_OLAH4.Range("H:K").Delete
    SH1_OLAH4.Range("J:L").Delete
    SH1_OLAH4.Range("K:K").Delete
    SH1_OLAH4.Range("H1").value = "LAST RECEIVE (MONTH)"
    SH1_OLAH4.Range("I1").value = "LAST RECEIVE (YEAR)"
    SH1_OLAH4.Range("J1").value = "AMOUNT PER BIYER"
    SH1_OLAH4.Rows(1).Insert
    LR1 = SH1_OLAH4.Range("A" & Rows.Count).End(xlUp).Row
    
    SH1_OLAH4.Range("J1").FormulaR1C1 = "=SUBTOTAL(9,R[2]C:R[" & LR1 - 1 & "]C)"
    SH1_OLAH4.Range("J1").Style = "Comma"
    
    SH1_OLAH4.Cells.EntireColumn.AutoFit
    SH1_OLAH4.Cells(1, 1).Select
    
'    Range("H:I").Delete
'    LR1 = SH1_OLAH4.Range("A" & Rows.Count).End(xlUp).Row
'    SH1_OLAH4.Range("A3:K" & LR1).Borders.LineStyle = xlContinuous
'    With SH1_OLAH4.Range("A3:K3")
'        .Interior.Color = vbYellow
'        .Font.Bold = True
'    End With
'    Cells.EntireColumn.AutoFit

''[AKHIRT PROSES]
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Call DeleteSheetsExcept("Set_Ref", "HOME", "RPA1", "RPA2", "STATUS", "WA", "OLAH1", "OLAH2", "OLAH3", "OLAH4")
    SH1_OLAH1.Name = "SOURCE_PIVOT"
    SH1_OLAH2.Name = "PIVOT"
    SH1_OLAH3.Name = "RESUME LEDGER DAN AGING"
    SH1_OLAH4.Name = "RESUME AGING PERBUYER"
    SH1_Home.Activate
    SH1_Home.Cells(1, 1).Select
    
''[SAVE HASIL]
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    Sheets(Array("SOURCE_PIVOT", "PIVOT", "RESUME LEDGER DAN AGING", "RESUME AGING PERBUYER")).Copy
    Set WB2 = ActiveWorkbook
    Windows(WB2.Name).Activate
    Application.DisplayAlerts = False
    WB2.SaveAs SH1_Home.Range("E20"), xlOpenXMLWorkbook
    WB2.Close False
    
End Sub

Sub Data_Validation()
    Set WB1 = ThisWorkbook
    Set SH1_Home = WB1.Worksheets("HOME")
    
    If Dir(SH1_Home.Range("E14"), vbDirectory) = "" Then
        MsgBox "LOKASI FOLDER TARIKAN JDE TIDAK DITEMUKAN", vbExclamation, "PERINGATAN"
        End
    End If
    
    If Dir(SH1_Home.Range("E15")) = "" Then
        MsgBox "FILE ILDV FOR BRANCH " & SH1_Home.Range("E12") & " DOESN'T EXISTS", vbExclamation, "FILE NOT FOUND"
        End
    End If
    If Dir(SH1_Home.Range("E16")) = "" Then
        MsgBox "FILE SHIPMENT REPORT FOR BRANCH " & SH1_Home.Range("E12") & " DOESN'T EXISTS", vbExclamation, "FILE NOT FOUND"
        End
    End If
    If Dir(SH1_Home.Range("E17")) = "" Then
        MsgBox "FILE REPORT WO BUYER FOR BRANCH " & SH1_Home.Range("E12") & " DOESN'T EXISTS", vbExclamation, "FILE NOT FOUND"
        End
    End If
    If Dir(SH1_Home.Range("E18")) = "" Then
        MsgBox "FILE INVENTORY AGING FOR BRANCH " & SH1_Home.Range("E12") & " DOESN'T EXISTS", vbExclamation, "FILE NOT FOUND"
        End
    End If
    If Dir(SH1_Home.Range("E19")) = "" Then
        MsgBox "FILE SALES LOCAL FOR BRANCH " & SH1_Home.Range("E12") & " DOESN'T EXISTS", vbExclamation, "FILE NOT FOUND"
        End
    End If
    
    If Dir(SH1_Home.Range("E20"), vbDirectory) = "" Then
        MsgBox "LOKASI FOLDER HASIL MAKRO TIDAK DITEMUKAN", vbExclamation, "PERINGATAN"
        End
    End If
    
''[CEK SHEETS DI WB INVENTORY AGING]...........................................................................................
' -----------------------------------------------------------------------------------------------------------
    STR_NAMA_SHEET = SH1_Home.Range("E12") & " " & SH1_Home.Range("E13")
    Set WB2 = Workbooks.Open(SH1_Home.Range("E18"))
    Windows(WB2.Name).Activate
    If Not wsx(STR_NAMA_SHEET) Then
        WB2.Close False
        SH1_Home.Activate
        MsgBox "Sheet " & STR_NAMA_SHEET & " Tidak Ditemukan di File Inventory Aging", vbInformation, "SHEET NOT FOUND"
        End
    End If
    WB2.Close False
    
End Sub


