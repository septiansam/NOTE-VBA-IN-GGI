Sub PROSES()

Application.DisplayAlerts = False

Dim twb As Workbook, wbSRC As Workbook, srcNAME As String
Dim path As String
Dim i As Long, lr As Long, lc As Long, r As Long, rPaste As Long, x As Long
Dim rg As Range, cell As Range, rng As Range, ws As Worksheet, countUPDATE As Long
Dim pertama As Boolean, barisFilter As Long
Dim shTombol As Worksheet

Set twb = ThisWorkbook
Set shTombol = twb.Sheets("TOMBOL")

pertama = False
path = shTombol.Range("D5") & Application.PathSeparator & "*.xlsx"

If Dir(path) = vbNullString Then
    MsgBox "File AR tidak ditemukan", vbExclamation
    Exit Sub
End If

For i = twb.Sheets.Count To 2 Step -1
    Sheets(i).Delete
Next i

If Evaluate("isref('" & "EMAIL" & "'!A1)") Then Sheets("EMAIL").Delete
Sheets.Add(After:=Sheets("TOMBOL")).Name = "EMAIL"

For i = 1 To 5
    If Evaluate("isref('" & "TES" & i & "'!A1)") Then Sheets("TES" & i).Delete
    Sheets.Add(After:=Sheets(Sheets.Count)).Name = "TES" & i
Next i

Set wbSRC = Workbooks.Open(path)
srcNAME = wbSRC.Name
wbSRC.Activate
Sheets("EMAIL").Select
Cells.Copy
twb.Activate
Sheets("EMAIL").Select: Range("a1").PasteSpecial xlPasteAll: Application.CutCopyMode = False: Cells.EntireColumn.AutoFit: Cells(1, 1).Select

For Each cell In Range(Cells(2, 3), Cells(Cells(Rows.Count, 3).End(xlUp).Row, 4))
    If Left(cell.Value, 1) = "<" And Right(cell.Value, 1) = ">" Then
        cell.Value = Mid(cell.Value, 2, Len(cell.Value) - 2)
    End If
Next cell

For i = 1 To Range("A" & Rows.Count).End(xlUp).Row
    If Cells(i, 1) = "WEARTEX" Then
        Cells(i, 2) = "Fung2"
        Cells(i, 3) = "afungfung@gistexgroup.com"
        Exit For
    End If
Next i

Dim kolomStatus As Long
wbSRC.Activate
pertama = True

For Each ws In Worksheets
    ws.Activate
    If ws.AutoFilterMode = True Then Selection.AutoFilter
    countUPDATE = Application.WorksheetFunction.CountIf(ws.Range("Q:Q"), "UPDATE")
    
    If countUPDATE > 1 And pertama = True Then
        ws.Activate
        kolomStatus = Cells.Find("UPDATE", , , xlPart).Column
        Range("A1").End(xlDown).Select
        r = Selection.Row
        If Cells(r + 1, 1) = vbNullString Then
            barisFilter = r + 1
            Range("A" & r).EntireRow.Copy
            twb.Activate
            Sheets("TES1").Select
            Range("A1").PasteSpecial xlPasteValues: Application.CutCopyMode = False: Cells.EntireColumn.AutoFit: Cells(1, 1).Select
            Range("A1").EntireRow.Font.Bold = True
            rPaste = Cells(Rows.Count, 1).End(xlUp).Row + 1
        Else
            barisFilter = r
            twb.Activate
            Sheets("TES1").Select
            rPaste = Cells(Rows.Count, 1).End(xlUp).Row + 1
        End If
        pertama = False
    End If
    
    If countUPDATE > 1 And pertama = False Then
        wbSRC.Activate
        ws.Activate
        kolomStatus = Cells.Find("UPDATE", , , xlPart).Column
        ActiveSheet.Rows(barisFilter).AutoFilter Field:=kolomStatus, Criteria1:="UPDATE"
        Range(Cells(barisFilter + 1, 1), Cells(Range("a" & Rows.Count).End(xlUp).Row, 200)).SpecialCells(xlCellTypeVisible).Copy
        twb.Activate
        Sheets("TES1").Select
        Range("a" & rPaste).PasteSpecial xlPasteAll: Application.CutCopyMode = False: Cells.EntireColumn.AutoFit: Cells(1, 1).Select
        wbSRC.Activate
    End If
Next ws
wbSRC.Close False

twb.Activate
Dim cPembeli As Long, cInvoice As Long, cTanggal As Long, cAmount As Long
cPembeli = Cells.Find("PEMBELI", , , xlWhole).Column
cInvoice = Cells.Find("INVOICE", , , xlPart).Column
cTanggal = Cells.Find("TANGGAL PEB", , , xlPart).Column
cAmount = Cells.Find("AMOUNT", , , xlPart).Column

Cells(1, cPembeli).EntireColumn.Copy Sheets("TES2").Range("A1")
Cells(1, cInvoice).EntireColumn.Copy Sheets("TES2").Range("B1")
Cells(1, cTanggal).EntireColumn.Copy Sheets("TES2").Range("C1")
Cells(1, cAmount).EntireColumn.Copy Sheets("TES2").Range("D1")

Sheets("TES2").Select
lr = Range("a" & Rows.Count).End(xlUp).Row
lc = Cells(1, Columns.Count).End(xlToLeft).Column
Range(Cells(2, lc + 1), Cells(lr, lc + 1)).FormulaR1C1 = "=RC[-2]+30"
Range(Cells(2, lc + 2), Cells(lr, lc + 2)).FormulaR1C1 = "=IF(RC[-1]<NOW(),""tagih"",""-"")"
Range(Cells(2, lc + 1), Cells(lr, lc + 2)).Select
With Selection
    .Copy
    .PasteSpecial xlPasteValues: Application.CutCopyMode = False: Cells.EntireColumn.AutoFit: Cells(1, 1).Select
End With
Sheets("TES2").Select
If Sheets("TES2").AutoFilterMode = True Then Selection.AutoFilter
Range("A1").EntireRow.AutoFilter Field:=lc + 2, Criteria1:="tagih"
Cells.SpecialCells(xlCellTypeVisible).Copy

Sheets("TES3").Select
Range("A1").PasteSpecial xlPasteValuesAndNumberFormats: Application.CutCopyMode = False: Cells.EntireColumn.AutoFit: Cells(1, 1).Select
lr = Range("a" & Rows.Count).End(xlUp).Row
lc = Cells(1, Columns.Count).End(xlToLeft).Column
Range(Cells(1, lc + 1), Cells(lr, lc + 2)).EntireColumn.Clear
Range("C:C").Delete
lc = Cells(1, Columns.Count).End(xlToLeft).Column

Sheets("TES3").Select
If Sheets("TES3").AutoFilterMode = True Then Selection.AutoFilter
Range("A1").AutoFilter
ActiveSheet.AutoFilter.Sort.SortFields.Clear
ActiveSheet.AutoFilter.Sort.SortFields.Add2 Key:=Range( _
    "A1:A" & lr), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:= _
    xlSortTextAsNumbers
With ActiveSheet.AutoFilter.Sort
    .Header = xlYes
    .MatchCase = False
    .Orientation = xlTopToBottom
    .SortMethod = xlPinYin
    .Apply
End With

If Sheets("TES3").AutoFilterMode = True Then Sheets("TES3").AutoFilterMode = False

Sheets("EMAIL").Select

Dim newBuyer As String
For Each cell In Range("A2:A" & Range("A" & Rows.Count).End(xlUp).Row)
    For i = 1 To Len(cell.Value)
        newBuyer = Replace(cell.Value, "  ", " ")
    Next i
    cell.Value = newBuyer
Next cell

If Evaluate("isref('" & "BANTUAN" & "'!A1)") Then Sheets("BANTUAN").Delete
Sheets.Add(After:=Sheets(Sheets.Count)).Name = "BANTUAN"

Sheets("TES3").Select
Range("a:a").Copy Sheets("BANTUAN").Range("a1")

Sheets("BANTUAN").Select
Range("A:A").RemoveDuplicates 1, xlYes: Rows(1).Font.Bold = True
For Each cell In Range("A2:A" & Range("A" & Rows.Count).End(xlUp).Row)
    For i = 1 To Len(cell.Value)
        newBuyer = Replace(cell.Value, "  ", " ")
    Next i
    cell.Value = newBuyer
Next cell

Range("B1") = "EMAIL"
Range("C1") = "STATUS"
Range("D1") = "CC"
lr = Range("a" & Rows.Count).End(xlUp).Row
Range(Cells(2, 2), Cells(lr, 2)).FormulaR1C1 = _
        "=IFERROR(VLOOKUP(RC[-1],EMAIL!C[-1]:C[1],3,0),""Tidak Ditemukan"")"
Range(Cells(2, 4), Cells(lr, 4)).FormulaR1C1 = _
        "=IFERROR(VLOOKUP(RC[-3],EMAIL!C[-3]:C,4,0),""Tidak Ditemukan"")"
        
Range(Cells(2, 2), Cells(lr, 4)).Select
With Selection
    .Copy
    .PasteSpecial xlPasteValues: Application.CutCopyMode = False: Cells.EntireColumn.AutoFit: Cells(1, 1).Select
End With

Range("A1:A" & lr).Copy
Range("E1").PasteSpecial xlPasteAll: Application.CutCopyMode = False
Range("E1").Value = "ORIGIN PEMBELI"

Cells.EntireColumn.AutoFit: Cells(1, 1).Select

Dim invalidChars As Variant
invalidChars = Array(":", "\", "/", "?", "*", "[", "]")
For Each cell In Range("A2:A" & Cells(Rows.Count, "A").End(xlUp).Row)
    newBuyer = cell.Value
    For i = LBound(invalidChars) To UBound(invalidChars)
        newBuyer = Replace(newBuyer, invalidChars(i), "-")
    Next i
    cell.Value = newBuyer
Next cell

Dim sumPembeli As Long, j As Long, a As Variant, b As Variant, pesan As String, strPEMBELI As String
Dim sumInv As Long
sumPembeli = Application.WorksheetFunction.CountA(Range("a:a")) - 1

Dim originPembeli As String, newStrPembeli As String

For i = 1 To sumPembeli
    originPembeli = Sheets("BANTUAN").Cells(i + 1, 5)
    strPEMBELI = Sheets("BANTUAN").Cells(i + 1, 1)
    If Len(strPEMBELI) > 31 Then
        strPEMBELI = Left(strPEMBELI, 31)
        Sheets("BANTUAN").Cells(i + 1, 1) = strPEMBELI
    End If
    Sheets("TES4").Cells.Clear
    If Evaluate("isref('" & strPEMBELI & "'!A1)") Then Sheets(strPEMBELI).Delete
    
    If Sheets("BANTUAN").Cells(i + 1, 2).Value <> "Tidak Ditemukan" Or Sheets("BANTUAN").Cells(i + 1, 4).Value <> "Tidak Ditemukan" Then
    
        Sheets.Add(After:=Sheets(Sheets.Count)).Name = strPEMBELI
        Sheets(strPEMBELI).Select
    
        Cells(2, 1) = originPembeli
        Cells(2, 2) = Sheets("BANTUAN").Cells(i + 1, 2)
        Sheets("TES3").Select
        If Sheets("TES3").AutoFilterMode = True Then Selection.AutoFilter
        Range("A1").AutoFilter 1, Sheets("BANTUAN").Cells(i + 1, 1) & "*"
        Cells.SpecialCells(xlCellTypeVisible).Copy
        Sheets("TES4").Select
        Range("a1").PasteSpecial xlPasteAll: Application.CutCopyMode = False: Cells.EntireColumn.AutoFit: Cells(1, 1).Select
        lr = Range("a" & Rows.Count).End(xlUp).Row
        sumInv = Cells(Rows.Count, 2).End(xlUp).Row - 1
        For j = 1 To sumInv
            a = Cells(j + 1, 2)
            b = Cells(j + 1, 3)
            b = Format(b, "#,##0.0")
            pesan = "No Invoice " & a & " : " & b
            Cells(j + 1, 4) = pesan
        Next j
        Range("D2:D" & lr).Copy
        Sheets(strPEMBELI).Select
        Range("C2").PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:= _
            False, Transpose:=True: Cells.WrapText = False: Cells.EntireColumn.AutoFit: Range("A1").Select
    
    '    If ActiveSheet.Name = "KARYA SUKSES MANDIRI-KEMAYORAN" Then
    '    Stop
    '    End If
    
        Range("A1").Value = "Buyer"
        Range("B1").Value = "To"
        Range("C1").Value = "Ket 1"
        Range("D1").Value = "Ket 2"
        Range("E1").Value = "Ket 3"
        Range("F1").Value = "Ket 4"
        Range("G1").Value = "Ket 5"
        Range("H1").Value = "Ket 6"
        Range("I1").Value = "Ket 7"
        Range("J1").Value = "Ket 8"
        Range("K1").Value = "Ket 9"
        Range("L1").Value = "Ket 10"
        Range("M1").Value = "Ket 11"
        Range("N1").Value = "Ket 12"
        Range("O1").Value = "Ket 13"
        Range("P1").Value = "Ket 14"
        Range("Q1").Value = "Ket 15"
        Range("R1").Value = "Ket 16"
        Range("S1").Value = "Ket 17"
        Range("T1").Value = "Ket 18"
        Range("U1").Value = "Ket 19"
        Range("V1").Value = "Ket 20"
        
        Range("W1").Select
        Range("W1").Value = "Ket 21"
        Range("X1").Value = "Ket 22"
        Range("Y1").Value = "Ket 23"
        Range("Z1").Value = "Ket 24"
        Range("AA1").Value = "Ket 25"
        Range("AB1").Value = "Ket 26"
        Range("AC1").Value = "Ket 27"
        Range("AD1").Value = "Ket 28"
        Range("AE1").Value = "Ket 29"
        Range("AF1").Value = "Ket 30"
        Range("AG1").Value = "Ket 31"
        Range("AH1").Value = "Ket 32"
        Range("AI1").Value = "Ket 33"
        Range("AJ1").Value = "Ket 34"
        Range("AK1").Value = "Ket 35"
        Range("AL1").Value = "Ket 36"
        Range("AM1").Value = "Ket 37"
        Range("AN1").Value = "Ket 38"
        Range("AO1").Value = "Ket 39"
        Range("AP1").Value = "Ket 40"
        
        Range("AQ1").Value = "Ket 41"
        Range("AR1").Value = "Ket 42"
        Range("AS1").Value = "Ket 43"
        Range("AT1").Value = "Ket 44"
        Range("AU1").Value = "Ket 45"
        Range("AV1").Value = "Ket 46"
        Range("AW1").Value = "Ket 47"
        Range("AX1").Value = "Ket 48"
        Range("AY1").Value = "Ket 49"
        Range("AZ1").Value = "Ket 50"
        Range("BA1").Value = "Ket 51"
        Range("BB1").Value = "Ket 52"
        Range("BC1").Value = "Ket 53"
        Range("BD1").Value = "Ket 54"
        Range("BE1").Value = "Ket 55"
        Range("BF1").Value = "Ket 56"
        Range("BG1").Value = "Ket 57"
        Range("BH1").Value = "Ket 58"
        Range("BI1").Value = "Ket 59"
        Range("BJ1").Value = "Ket 60"
        Range("BK1").Value = "Ket 61"
        Range("BL1").Value = "Ket 62"
        Range("BM1").Value = "Ket 63"
        Range("BN1").Value = "Ket 64"
        Range("BO1").Value = "Ket 65"
        Range("BP1").Value = "Ket 66"
        Range("BQ1").Value = "Ket 67"
        Range("BR1").Value = "Ket 68"
        Range("BS1").Value = "Ket 69"
        Range("BT1").Value = "Ket 70"
        Range("BU1").Value = "Ket 71"
        Range("BV1").Value = "Ket 72"
        Range("BW1").Value = "Ket 73"
        Range("BX1").Value = "Ket 74"
        Range("BY1").Value = "Ket 75"
        Range("BZ1").Value = "Ket 76"
        Range("CA1").Value = "Ket 77"
        Range("CB1").Value = "Ket 78"
        Range("CC1").Value = "Ket 79"
        Range("CD1").Value = "Ket 80"
    '    Range("CE1").Value = "Ket 80"
    
    '    Range("CF1").Value = "STATUS"
        
        Range("CE1").Value = "STATUS"
        
        If Cells(2, Columns.Count).End(xlToLeft).Column > Cells(1, Columns.Count).End(xlToLeft).Column Then
            MsgBox "Oooops, ada invoice yg melebihi Ket 80!", vbExclamation, "PROSES DIHENTIKAN"
            End
        End If
        
        If Cells(2, Columns.Count).End(xlToLeft).Column >= 3 Then
            Cells(2, Cells(1, Columns.Count).End(xlToLeft).Column).Select
            Cells(2, Cells(1, Columns.Count).End(xlToLeft).Column).Value = 1
        Else
            Cells(2, Cells(1, Columns.Count).End(xlToLeft).Column).Value = 0
        End If
    
    
        Range("A:A").Insert
        Range("A1").Select
        Range("A1") = "CC"
        Range("A2").FormulaR1C1 = _
            "=INDEX(BANTUAN!C[3],MATCH('" & strPEMBELI & "'!RC[1],BANTUAN!C[4],0))"
        Range("A2").Select
        With Selection
            .Copy
            .PasteSpecial xlPasteValues: Application.CutCopyMode = False: Cells.EntireColumn.AutoFit: Cells(1, 1).Select
        End With
    
        Dim toEmail As String, ccAwal As String, ccTambahan As String
        toEmail = Range("c2").Value
        ccAwal = Range("a2").Value
        ccTambahan = "chalim@gistexgroup.com, nhalimah@gistexgroup.com, ssilalahi@gistexgroup.com"
        
        Range("c2").FormulaR1C1 = toEmail
        Range("a2").FormulaR1C1 = ccAwal
        If ccAwal = "Tidak Ditemukan" Then
            Range("a2").FormulaR1C1 = ccTambahan
        Else
            Range("a2").FormulaR1C1 = ccAwal & ", " & ccTambahan
        End If
        Cells.EntireColumn.AutoFit: Cells(1, 1).Select
    End If
Next i

Sheets("TES3").Select
If Sheets("TES3").AutoFilterMode = True Then Selection.AutoFilter
ActiveSheet.Name = "RPA"
Sheets("RPA").Select
Range("A1").AutoFilter

If Evaluate("isref('" & "EMAIL" & "'!A1)") Then Sheets("EMAIL").Delete
For i = 1 To 5
    If Evaluate("isref('" & "TES" & i & "'!A1)") Then Sheets("TES" & i).Delete
Next i

shTombol.Activate: Cells(1, 1).Select
twb.Save

'MsgBox "Running Success...", vbInformation, "Program Berhasil Dijalankan.."

Application.DisplayAlerts = True
End Sub

