'******************************************************
'
'[08 MARET 2024]
'
'******************************************************

Public TWB As Workbook
Public SH_TARIKAN As Worksheet
Public SH_CEK As Worksheet
Public SH_PIC As Worksheet
Public SH_PO As Worksheet, ws As Worksheet
Public TEMP1 As Worksheet, TEMP2 As Worksheet, TEMP3 As Worksheet, TEMP4 As Worksheet, TEMP5 As Worksheet
Public PATH_TARIKAN As String, WB_TARIKAN As Workbook
Public PATH_CEK As String, WB_CEK As Workbook
Public i As Long, COUNT_PIC As Long
Public j As Long, x As Long, k As Long
Public i_ReceiptDate As Long
Public ROW_PASTE As Long
Public ARR_KRITERIA() As Variant
Public DATA_RANGE As Range, CELL As Range
Public SH_SRC_CEK As Worksheet
Public STR_BU As String, STR_PO As String, STR_OR As String, STR_ITEM As String
Public LR_CEK As Long, LC_CEK As Long, LR_TARIKAN As Long, LC_TARIKAN As Long
Public LR_TEMP2 As Long, LR_TEMP3 As Long
Public COUNT_PO As Long
Public LASTROW As Long, LASTCOL As Long
Public LOOKUP As String
Public str_Month As String, str_Title As String
Public wb_Monthly As Workbook, path_Monthly As String
Public arr_Sheets() As String, ws_Names As String
Public namaBulan As String, monthNumber As Long
Public firstDate As Date, lastDate As Date, tahun As Long
Public shPeriode As Worksheet
Public data As Variant, dataLot As Variant, dataOR As Variant
Public dataFind As String, dataPO As String
Public found As Boolean
Public candidate As String

Public folderPath As String, folderYear As String
Const PathSRC As String = "\\10.8.0.35\Bersama\IT\RPA Purchasing\WO Purchasing\Performance\.backup\FIELDS RESUME"

Sub MAIN()
Application.DisplayAlerts = False

Set TWB = ThisWorkbook

For i = TWB.Sheets.Count To 3 Step -1
    Sheets(i).Delete
Next i

'PATH_TARIKAN = HOME.Range("C8").Value & _
'                Application.PathSeparator & _
'                HOME.Range("D8").Value & _
'                HOME.Range("E8").Value
PATH_CEK = HOME.Range("C9").Value & _
                Application.PathSeparator & _
                HOME.Range("D9").Value & _
                HOME.Range("E9").Value
                
'If Dir(PATH_TARIKAN) = "" Then
'    MsgBox "File " & HOME.Range("D8") & " Doesn't Exists", vbCritical, "File Not Found"
'    Exit Sub
'End If

If Dir(PATH_CEK) = "" Then
    MsgBox "File " & HOME.Range("D9") & " Doesn't Exists", vbCritical, "File Not Found"
    Exit Sub
End If

'If WorksheetExists("TARIKAN") Then Sheets("TARIKAN").Delete
If WorksheetExists("CEK") Then Sheets("CEK").Delete
If WorksheetExists("S_PIC") Then Sheets("S_PIC").Delete
If WorksheetExists("S_PO") Then Sheets("S_PO").Delete

For i = 1 To 2
    If WorksheetExists("TEMP" & i) Then Sheets("TEMP" & i).Delete
Next i

Set SH_TARIKAN = TWB.Sheets("TARIKAN")

Set SH_CEK = Sheets.Add(After:=Sheets(Sheets.Count)): ActiveSheet.Name = "CEK"
Set SH_PIC = Sheets.Add(After:=Sheets(Sheets.Count)): ActiveSheet.Name = "S_PIC"
Set SH_PO = Sheets.Add(After:=Sheets(Sheets.Count)): ActiveSheet.Name = "S_PO"

Set TEMP1 = Sheets.Add(After:=Sheets(Sheets.Count)): ActiveSheet.Name = "TEMP1"
Set TEMP2 = Sheets.Add(After:=Sheets(Sheets.Count)): ActiveSheet.Name = "TEMP2"


'[DAPATKAN NAMA SETIAP PIC]
'...................................................................................
SH_TARIKAN.Activate: Cells.EntireColumn.AutoFit: Cells(1, 1).Select

Range("V:V").Copy SH_PIC.Range("A1")
SH_PIC.Activate
Range("A:A").RemoveDuplicates 1, xlYes
Range("A1").Delete xlUp
For i = LR To 1 Step -1
    If Cells(i, 1) = "" Or _
        Cells(i, 1) = "NSITI" Then
        
        Cells(i, 1).Delete xlUp
        
    End If
Next i
COUNT_PIC = LR(SH_PIC)

'[DAPATKAN NAMA SETIAP PO]
'...................................................................................
SH_TARIKAN.Activate

If SH_TARIKAN.AutoFilterMode = True Then SH_TARIKAN.AutoFilterMode = False
Range("A1:AM" & LR).AutoFilter 39, "1"
Range("AL1:AL" & LR).SpecialCells(xlCellTypeVisible).Copy SH_PO.Range("A1")
If SH_TARIKAN.AutoFilterMode = True Then SH_TARIKAN.AutoFilterMode = False

'[AMBIL DATA CEK]
'[OLAH DATA DAN AMBIL YANG HANYA DIPERLUKAN]
'.........................................................................................

SH_PO.Activate

Set DATA_RANGE = Range("A2:A" & Cells(Rows.Count, "A").End(xlUp).Row)
ReDim ARR_KRITERIA(1 To DATA_RANGE.Rows.Count)
i = 1
DATA_RANGE.Activate
For Each CELL In DATA_RANGE
    ARR_KRITERIA(i) = CELL.Value
    i = i + 1
Next CELL

Set WB_CEK = Workbooks.Open(PATH_CEK)
WB_CEK.Activate
Set SH_SRC_CEK = WB_CEK.Sheets(1)
SH_SRC_CEK.Select

'LASTROW = LR
'
'If SH_SRC_CEK.AutoFilterMode = True Then SH_SRC_CEK.AutoFilterMode = False
'
'Range("A:A").Insert: Range("A1") = "BU"
'With Range("A2:A" & LR)
'    .FormulaR1C1 = "=C3*1"
'    .Copy
'    .PasteSpecial xlPasteValuesAndNumberFormats: Application.CutCopyMode = False: Cells(1, 1).Select
'End With
'
'Range("A:A").Insert: Range("A1") = "OR"
'With Range("A2:A" & LR)
'    .FormulaR1C1 = _
'        "=IF(LEN(RC[8])=8,RC[8],IF(AND(RC[1]=1205,LEN(RC[8])>18),LEFT(RC[8],8),IF(OR(RC[1]=1205,LEN(RC[8])=18,LEFT(RC[8],2)=""40"",LEFT(RC[8],3)=""/23""),RIGHT(RC[8],8),IF(AND(RC[13]=""ININ"",LEN(RC[8])>28,RC[9]<>""23001272""),MID(RC[8],LEN(RC[8])-16,8),IF(RC[13]=""INFA"",MID(RC[8],LEN(RC[8])-16,8),RIGHT(RC[8],8))))))"
'    .Copy
'    .PasteSpecial xlPasteAll: Application.CutCopyMode = False: Cells(1, 1).Select
'End With


'For i = 2 To LASTROW
'    dataFind = SH_SRC_CEK.Cells(i, 9).Value ' Kolom A
'    dataFind = Trim$(dataFind)
'    dataPO = SH_SRC_CEK.Cells(i, 10).Value ' Kolom B
'    found = False
'
'    For j = Len(dataFind) - 7 To 1 Step -1
'        If Left(dataFind, 2) Like String(2, "#") Or _
'            Left(dataFind, 1) Like String(1, "/") Then
'
'            candidate = Mid(dataFind, j, 8)
'            If candidate Like String(8, "#") _
'                And candidate <> dataPO _
'                And Right(candidate, 6) <> Right(dataPO, 6) _
'                And Left(candidate, 1) = Left(dataPO, 1) Then
'
'                SH_SRC_CEK.Cells(i, 1).Value = candidate
'                found = True
'                Exit For ' Keluar dari loop jika sudah menemukan
'            End If
'        End If
'    Next j
'
'    If Not found Then
'        SH_SRC_CEK.Cells(i, 1).Value = "Tidak ada"
'    End If
'Next i
'
'Range("A:A").Insert: Range("A1") = "PO-OR.No-Item.Short-BU"
'With Range("A2:A" & LR)
'    .FormulaR1C1 = "=RC[10]&""-""&RC[1]&""-""&RC[5]&""-""&RC[2]"
'    .Copy
'    .PasteSpecial xlPasteValuesAndNumberFormats: Application.CutCopyMode = False: Cells(1, 1).Select
'End With
'
'SH_CEK.Range(SH_CEK.Cells(1, 1), SH_CEK.Cells(LR(SH_CEK), LC(SH_CEK))).Clear

Call PreprocessingDataCek

LR_CEK = LR(WB_CEK.Sheets(1))
LC_CEK = LC(WB_CEK.Sheets(1))

If SH_SRC_CEK.AutoFilterMode = True Then SH_SRC_CEK.AutoFilterMode = False

'FILTER
SH_SRC_CEK.Range("$A$1:$SAM$" & LR_CEK).AutoFilter Field:=1, Criteria1:=ARR_KRITERIA(), Operator:=xlFilterValues
If Range("C" & Rows.Count).End(xlUp).Value <> "PO" Then
    Range(Cells(1, 1), Cells(LR_CEK, LC_CEK)).SpecialCells(xlCellTypeVisible).Copy

    SH_CEK.Range("A1").PasteSpecial xlPasteValuesAndNumberFormats

    Application.CutCopyMode = False
Else
    Stop
End If

'NOT FILTER
'Range(Cells(1, 1), Cells(LR_CEK, LC_CEK)).Copy
'SH_CEK.Range("A1").PasteSpecial xlPasteValuesAndNumberFormats
'Application.CutCopyMode = False

WB_CEK.Close False

'[PENGOLAHAN]
'.............................................
SH_CEK.Activate
LASTROW = LR(SH_CEK)
LASTCOL = LC(SH_CEK)

'[COUNT (PO - OR - ITEM)]
Range("Y1") = "COUNT (PO - OR - ITEM)"
With Range("Y2:Y" & LASTROW)
    .FormulaR1C1 = "=COUNTIF(R1C1:RC[-24],RC[-24])"
    .Copy
    .PasteSpecial xlPasteValuesAndNumberFormats: Application.CutCopyMode = False
End With

'[PO - OR - ITEM - COUNT]
Range("Z1") = "PO-OR-ITEM-COUNT"
With Range("Z2:Z" & LASTROW)
    .FormulaR1C1 = "=CONCATENATE(RC[-25],""-"",RC[-1])"
    .Copy
    .PasteSpecial xlPasteValuesAndNumberFormats: Application.CutCopyMode = False
End With

Range("A:A").Insert: Range("A1").Value = "Start.Produksi.PIC"
With Range("A2:A" & LASTROW)
    .FormulaR1C1 = "=IFERROR(INDEX(TARIKAN!C[21],MATCH(CEK!RC[1],TARIKAN!C[37],0)),"""")"
    .Copy
    .PasteSpecial xlPasteValuesAndNumberFormats: Application.CutCopyMode = False: Cells(1, 1).Select
End With

If SH_CEK.AutoFilterMode = True Then SH_CEK.AutoFilterMode = False
Cells.AutoFilter 1, "="

Range(Cells(1, 1), Cells(LASTROW, LASTCOL)).Offset(1).Delete Shift:=xlUp
SH_CEK.AutoFilterMode = False
Cells.EntireColumn.AutoFit: Cells(1, 1).Select

'LASTROW = LR(SH_CEK)

'Range("A:A").Insert: Range("A1").Value = "COUNT"
'With Range("A2:A" & LASTROW)
'    .FormulaR1C1 = "=COUNTIF(R1C3:RC[2],RC[2])"
'    .Copy
'    .PasteSpecial xlPasteValuesAndNumberFormats: Application.CutCopyMode = False: Cells(1, 1).Select
'End With
'Range("A:A").Insert: Range("A1").Value = "LOOKUP - COUNT"
'With Range("A2:A" & LASTROW)
'    .FormulaR1C1 = "=RC[3]&""-""&RC[1]"
'    .Copy
'    .PasteSpecial xlPasteValuesAndNumberFormats: Application.CutCopyMode = False: Cells(1, 1).Select
'End With
'Range("B:B").Delete

SH_TARIKAN.Activate
'LASTROW = LR(SH_TARIKAN)

'Stop
'BARU 14 - MARET - 2024
'Range("AM1") = "PO-OR.No-Item.Short-BU-WO"
'Range("AN1") = "COUNT"
'Range("AO1") = "LOOKUP - COUNT"
'Range("AP1") = "PO-OR.No-Item"
'Range("AM2:AM" & LASTROW).FormulaR1C1 = "=RC[-1]&""-""&RC[-28]"
'Range("AN2:AN" & LASTROW).FormulaR1C1 = "=COUNTIF(R1C39:RC[-1],RC[-1])"
'Range("AO2:AO" & LASTROW).FormulaR1C1 = "=RC[-3]&""-""&COUNTIF(R1C38:RC[-3],RC[-3])"
'Range("AP2:AP" & LASTROW).FormulaR1C1 = "=CONCATENATE(RC[-39],""-"",RC[-33],""-"",RC[-35])"
'
'With Range("AM2:AP" & LASTROW)
'    .Copy: .PasteSpecial xlPasteValuesAndNumberFormats: Application.CutCopyMode = False
'End With

Dim i_PIC As Integer, STR_PIC As String, SH_STR_PIC As Worksheet
Dim LASTROW_TEMP1 As Long, LASTCOL_TEMP1 As Long
'TWB.Save
LR_TARIKAN = LR(SH_TARIKAN)
LC_TARIKAN = LC(SH_TARIKAN)
LR_CEK = LR(SH_CEK)
'Stop


'Stop
Dim KEY As String
For i_PIC = 1 To COUNT_PIC
    'If i_PIC = 6 Then Stop
    '[BERSIHKAN DULU]
    '''''''''''''''''
    SH_PO.Cells.Clear
    TEMP1.Cells.Clear
    TEMP2.Cells.Clear
    TEMP2.Cells.Borders.LineStyle = xlNone
    TEMP2.Cells.Interior.Color = xlNone
    
    '''''''''''''''''
    '[SELESAI]

    SH_PIC.Activate
    STR_PIC = Cells(i_PIC, 1).Value

    If WorksheetExists(STR_PIC) Then Sheets(STR_PIC).Delete
    Set SH_STR_PIC = Sheets.Add(After:=Sheets(Sheets.Count)): ActiveSheet.Name = STR_PIC
    
    SH_TARIKAN.Activate
    'TWB.Save
    If SH_TARIKAN.AutoFilterMode = True Then Selection.AutoFilter

    SH_TARIKAN.Range("A1:SAM" & LR_TARIKAN).AutoFilter 22, STR_PIC
    'Stop
    'SH_TARIKAN.Range("A1:AZ" & LR_TARIKAN).AutoFilter 40, "1"
    'Stop
    ' BARU 15-03-2024
    Range(Cells(1, 1), Cells(LR_TARIKAN, LC_TARIKAN)).SpecialCells(xlCellTypeVisible).Copy TEMP1.Range("A1")
    Range(Cells(1, 38), Cells(LR_TARIKAN, 38)).Copy SH_PO.Range("A1")
    SH_TARIKAN.AutoFilterMode = False
    
    SH_PO.Activate
    Range("A:A").RemoveDuplicates 1, xlYes
    LASTROW = Range("A" & Rows.Count).End(xlUp).Row
    Range("B1") = "Count Lookup Tarikan"
    Range("C1") = "Count Lookup Cek"
    Range("D1") = "Max"
    Range("B2:B" & LASTROW).FormulaR1C1 = "=COUNTIF(TEMP1!C[36],S_PO!RC[-1])"
    Range("C2:C" & LASTROW).FormulaR1C1 = "=COUNTIF(CEK!C[-1],S_PO!RC[-2])"
    Range("D2:D" & LASTROW).FormulaR1C1 = "=MAX(RC[-2],RC[-1])"
    With Range("A2:D" & LASTROW)
        .Copy
        .PasteSpecial xlPasteValuesAndNumberFormats: Application.CutCopyMode = False: Cells.EntireColumn.AutoFit
    End With
    Range("B:C").Delete
    
    Range("D1") = "PO-OR-ITEM"
'    With Range("D2:D" & LASTROW)
'        .FormulaR1C1 = "=LEFT(RC[-3],LEN(RC[-3])-5)"
'        .Copy
'        .PasteSpecial xlPasteValuesAndNumberFormats: Application.CutCopyMode = False: Cells.EntireColumn.AutoFit
'    End With
    Range("A2:A" & LASTROW).Copy Range("D2")
    
    Range("D:D").RemoveDuplicates 1, xlYes
    Range("E1") = "SORT"
    Range("E2") = "1"
    Range("E2").DataSeries xlColumns, xlLinear, , 1, Range("D" & Rows.Count).End(xlUp).Row - 1
    
    Cells.EntireColumn.AutoFit
    Cells(1, 1).Select
    'Stop
    ROW_PASTE = 1
    
    For i = 2 To LASTROW
        LOOKUP = SH_PO.Range("A" & i).Value
        For j = 1 To CInt(SH_PO.Range("B" & i).Value)
            TEMP2.Range("A" & ROW_PASTE + j).Value = LOOKUP
        Next j
        ROW_PASTE = TEMP2.Range("A" & Rows.Count).End(xlUp).Row
    Next i
    
    TEMP2.Activate
'    LASTROW = TEMP2.Range("A" & Rows.Count).End(xlUp).Row
    LASTROW = LR(TEMP2)
    
    '[X]... [15-APRIL-2024]
    'GET ITEM KARENA SUSAH JIKA DI BAWAH
    With Range("F2:F" & LASTROW)
        .FormulaR1C1 = "=RIGHT(RC[-5],LEN(RC[-5])-18)*1" '"=MID(RC[-5],19,LEN(RC[-5])-18-5)*1"
        .Copy
        .PasteSpecial xlPasteValuesAndNumberFormats: Application.CutCopyMode = False: Cells.EntireColumn.AutoFit
    End With
    
    'GET SORT KARENA SUSAH JIKA DI BAWAH
    With Range("Q2:Q" & LASTROW)
        .FormulaR1C1 = "=VLOOKUP(RC[-16],S_PO!C[-13]:C[-12],2,0)"
        .Copy
        .PasteSpecial xlPasteValuesAndNumberFormats: Application.CutCopyMode = False: Cells.EntireColumn.AutoFit
    End With

    Range("B1") = "LOOKUP - COUNT"
    With Range("B2:B" & LASTROW)
        .FormulaR1C1 = "=RC[-1]&""-""&COUNTIF(R1C1:RC[-1],RC[-1])"
        .Copy
        .PasteSpecial xlPasteValuesAndNumberFormats: Application.CutCopyMode = False: Cells.EntireColumn.AutoFit
    End With
    Range("A:A").Delete
    Range("G:H").NumberFormat = "m/d/yyyy"
    Range("L:M").NumberFormat = "[=0]0%;[<1]0.000%;0%;_(@_)"
    
    Cells(1, 2) = "UK(PO-OR-ITEM)"
    Cells(1, 3) = "PO.No": Cells(1, 4) = "OR.No"
    Cells(1, 5) = "ITEM": Cells(1, 6) = "Qty Order"
    Cells(1, 7) = "Receipt.Due.Date": Cells(1, 8) = "Tgl OV"
    Cells(1, 9) = "Qty Received": Cells(1, 10) = "Day Diff"
    Cells(1, 11) = "SUM Qty Received": Cells(1, 12) = "% Order Kedatangan"
    Cells(1, 13) = "Ontime": Cells(1, 14) = "CC": Cells(1, 15) = "USED"
    Cells(1, 16) = "SORT"
    
    Range("B2:B" & LASTROW).FormulaR1C1 = "=RC[1]&""-""&RC[2]&""-""&RC[3]"
    Range("N2:N" & LASTROW).FormulaR1C1 = "=COUNTIF(R1C2:RC[-12],RC[-12])"
    Range("O2:O" & LASTROW).FormulaR1C1 = "=IF(RC[-1]=1,RC[-4],"""")"
    Range("C2:C" & LASTROW).FormulaR1C1 = "=LEFT(RC[-2],8)*1"
    Range("D2:D" & LASTROW).FormulaR1C1 = "=MID(RC[-3],10,8)*1"

'    Range("E2:E" & LASTROW).FormulaR1C1 = _
'        "=IFERROR(INDEX(TARIKAN!C[2],MATCH(TEMP2!RC[-4],TARIKAN!C[36],0)),"""")"
'    Range("E2:E" & LASTROW).FormulaR1C1 = "=MID(RC[-4],19,LEN(RC[-4])-18-7)*1"

    Range("F2:F" & LASTROW).FormulaR1C1 = _
        "=IFERROR(INDEX(TEMP1!C[13],MATCH(TEMP2!RC[-5],TEMP1!C[34],0)),"""")"
    Range("G2:G" & LASTROW).FormulaR1C1 = _
        "=IFERROR(INDEX(TEMP1!C[10],MATCH(TEMP2!RC[-6],TEMP1!C[33],0)),"""")"
    
    Range("H2:H" & LASTROW).FormulaR1C1 = _
        "=IFERROR(INDEX(CEK!C[-2],MATCH(TEMP2!RC[-7],CEK!C[19],0)),"""")"
    Range("I2:I" & LASTROW).FormulaR1C1 = _
        "=IFERROR(INDEX(CEK!C[9],MATCH(TEMP2!RC[-8],CEK!C[18],0)),"""")"
'    Range("J2:J" & LASTROW).FormulaR1C1 = _
'        "=IF(AND(RC[-3]<>"""",RC[-2]<>""""),RC[-3]-RC[-2],"""")"
    
    '-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_
    '[X].. ISI SEL KOSONG DI RECEIPT DUE DATE DENGAN DATA DI ATASNYA
    'NEW
    With Range("G2:G" & LASTROW)
        .Copy
        .PasteSpecial xlPasteValuesAndNumberFormats: Application.CutCopyMode = False
    End With
    On Error Resume Next
    TEMP2.AutoFilterMode = False
'    Cells.AutoFilter 7, "="
'    Range(TEMP2.AutoFilter.Range.Offset(1).SpecialCells(xlCellTypeVisible).Cells(1, 7), _
'            TEMP2.AutoFilter.Range.SpecialCells(xlCellTypeVisible).Cells(Rows.Count, 7).End(xlUp)).ClearContents
'
'    TEMP2.AutoFilterMode = False
'    With Range("G2:G" & LASTROW)
'        .SpecialCells(xlCellTypeBlanks).FormulaR1C1 = "=IF(RC[2]<>"""",R[-1]C,"""")"
'        .Copy
'        .PasteSpecial xlPasteValuesAndNumberFormats: Application.CutCopyMode = False
'    End With
    
'    For i_ReceiptDate = 2 To LASTROW
'        If Cells(i_ReceiptDate, 7) = vbNullString _
'            And Cells(i_ReceiptDate - 1, 7) <> vbNullString Then
'
'            Cells(i_ReceiptDate, 7) = Cells(i_ReceiptDate - 1, 7)
'
'        End If
'    Next i_ReceiptDate
    
    ' Mendapatkan data dari kolom L ke dalam array
    data = Range("G2:G" & LASTROW).Value
    
    ' Mengisi sel kosong di kolom L dengan data di atasnya
    For i = 1 To UBound(data, 1)
        If IsEmpty(data(i, 1)) Or data(i, 1) = vbNullString Then
            data(i, 1) = data(i - 1, 1)
        End If
    Next i
    
    ' Memasukkan kembali data yang telah diperbarui ke dalam kolom L
    Range("G2:G" & LASTROW).Value = data
    
    On Error GoTo 0
    '-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_
    
    '[15-04-2024]..................
    '[X].. PINDAHKAN PERHITUNGAN KOLOM J DARI KODINGAN DI ATAS KE BAWAH INI
    Range("J2:J" & LASTROW).FormulaR1C1 = _
        "=IF(AND(RC[-3]<>"""",RC[-2]<>""""),RC[-3]-RC[-2],"""")"

    '[18-03-2024]..................
    'Terjadi Perubahan Perhitungan'
    Range("K2:K" & LASTROW).FormulaR1C1 = "=SUMIF(C[-9],RC[-9],C[-2])"
    'Range("L2:L" & LASTROW).FormulaR1C1 = "=IFERROR(IF(RC[-2]>0,RC[-3]/RC[-1],0),0)"
    
    '22-03-2024
    Range("L2:L" & LASTROW).FormulaR1C1 = _
        "=IFERROR(IF(AND(RC[-5]<>"""",RC[-4]<>"""",RC[-3]<>"""",RC[-2]<>"""",RC[-2]>0),IF(SUMIF(TEMP1!C[26],CONCATENATE(TEMP2!RC[-9],""-"",TEMP2!RC[-8],""-"",TEMP2!RC[-7]),TEMP1!C[7])>RC[-1],RC[-3]/SUMIF(TEMP1!C[26],CONCATENATE(TEMP2!RC[-9],""-"",TEMP2!RC[-8],""-"",TEMP2!RC[-7]),TEMP1!C[7]),RC[-3]/RC[-1]),0),0)"
        
        '"=IFERROR(IF(AND(RC[-5]<>"""",RC[-4]<>"""",RC[-3]<>"""",RC[-2]<>"""",RC[-2]>0),IF(SUMIF(TEMP1!C[30],CONCATENATE(TEMP2!RC[-9],""-"",TEMP2!RC[-8],""-"",TEMP2!RC[-7]),TEMP1!C[7])>RC[-1],RC[-3]/SUMIF(TEMP1!C[30],CONCATENATE(TEMP2!RC[-9],""-"",TEMP2!RC[-8],""-"",TEMP2!RC[-7]),TEMP1!C[7]),RC[-3]/RC[-1]),0),0)"
        '"=IFERROR(IF(AND(RC[-5]<>"""",RC[-4]<>"""",RC[-3]<>"""",RC[-2]<>"""",RC[-2]>0),IF(SUMIF(TEMP1!C[26],LEFT(TEMP2!RC[-11],LEN(TEMP2!RC[-11])-2),TEMP1!C[7])>RC[-1],RC[-3]/SUMIF(TEMP1!C[26],LEFT(TEMP2!RC[-11],LEN(TEMP2!RC[-11])-2),TEMP1!C[7]),RC[-3]/RC[-1]),0),0)"
        '"=IFERROR(IF(AND(RC[-5]<>"""",RC[-4]<>"""",RC[-3]<>"""",RC[-2]<>"""",RC[-2]>0),IF(RC[-6]>RC[-1],RC[-3]/RC[-6],RC[-3]/RC[-1]),0),0)"
        '"=IFERROR(IF(AND(RC[-5]<>"""",RC[-4]<>"""",RC[-3]<>"""",RC[-2]<>"""",RC[-2]>0),RC[-3]/RC[-1],0),0)"
        '"=IFERROR(IF(AND(RC[-5]<>"""",RC[-4]<>"""",RC[-3]<>"""",RC[-2]<>"""",RC[-2]>0),RC[-3]/SUMIF(C[-10],RC[-10],C[-6]),0),0)"
        '"=IFERROR(IF(RC[-2]>0,RC[-3]/SUMIF(C[-10],RC[-10],C[-6]),0),0)"
    'Done
    'Range("M2:M" & LASTROW).FormulaR1C1 = "=IF(RC[1]=1,SUMIF(C[-11],RC[-11],C[-1]),"""")"
    
    '[28-03-2024]...
    'Jika OnTime 100%, Tapi Ternyata SUM QTY Received < SUM QTY ORDER, Maka
    'SUM QTY ORDER dibagi SUM QTY Received
    Range("M2:M" & LASTROW).FormulaR1C1 = _
        "=IF(RC[1]=1,IF(AND(SUMIF(C[-11],RC[-11],C[-1])=1,RC[2]<SUMIF(TEMP1!C[25],CONCATENATE(TEMP2!RC[-10],""-"",TEMP2!RC[-9],""-"",TEMP2!RC[-8]),TEMP1!C[6])),RC[2]/SUMIF(TEMP1!C[25],CONCATENATE(TEMP2!RC[-10],""-"",TEMP2!RC[-9],""-"",TEMP2!RC[-8]),TEMP1!C[6]),SUMIF(C[-11],RC[-11],C[-1])),"""")"
        
        '"=IF(RC[1]=1,IF(AND(SUMIF(C[-11],RC[-11],C[-1])=1,RC[2]<SUMIF(TEMP1!C[29],CONCATENATE(TEMP2!RC[-10],""-"",TEMP2!RC[-9],""-"",TEMP2!RC[-8]),TEMP1!C[6])),RC[2]/SUMIF(TEMP1!C[29],CONCATENATE(TEMP2!RC[-10],""-"",TEMP2!RC[-9],""-"",TEMP2!RC[-8]),TEMP1!C[6]),SUMIF(C[-11],RC[-11],C[-1])),"""")"
        '"=IF(RC[1]=1,IF(AND(SUMIF(C[-11],RC[-11],C[-1])=1,RC[2]<SUMIF(TEMP1!C[25],LEFT(TEMP2!RC[-12],LEN(TEMP2!RC[-12])-2),TEMP1!C[6])),RC[2]/SUMIF(TEMP1!C[25],LEFT(TEMP2!RC[-12],LEN(TEMP2!RC[-12])-2),TEMP1!C[6]),SUMIF(C[-11],RC[-11],C[-1])),"""")"
        '"=IF(RC[1]=1,IF(AND(SUMIF(C[-11],RC[-11],C[-1])=1,RC[2]<SUMIF(C[-11],RC[-11],C[-7])),RC[2]/SUMIF(C[-11],RC[-11],C[-7]),SUMIF(C[-11],RC[-11],C[-1])),"""")"
    
    
    'Range("P2:P" & LASTROW).FormulaR1C1 = "=IFERROR(VLOOKUP(RC[-14],S_PO!C[-12]:C[-11],2,0),"""")"
    
    'NEW
'    Range("K2").FormulaR1C1 = "=SUM(RC[-2]:R[1718]C[-2])"
'    Range("L2:L" & LASTROW).FormulaR1C1 = _
'        "=IFERROR(IF(AND(RC[-2]>0,RC[-3]<>"""",RC[-2]<>""""),RC[-3]/R2C11,""""),"""")"
'    Range("M2").FormulaR1C1 = "=SUM(RC[-1]:R[1718]C[-1])"
    
    With Range("B2:P" & LASTROW)
        .Copy
        .PasteSpecial xlPasteValuesAndNumberFormats: Application.CutCopyMode = False
    End With
    
    Range("O2:O" & LASTROW).Copy Range("K2")
    TEMP2.Sort.SortFields.Clear
    TEMP2.Sort.SortFields.Add2 KEY:=Range("P2:P" & LASTROW _
        ), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortNormal
    With TEMP2.Sort
        .SetRange Range("A1:P" & LASTROW)
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    Range("A:B,N:P").Delete xlLeft
    Range("K:K").NumberFormat = "0%"
    
    'LASTROW = TEMP2.Range("A" & Rows.Count).End(xlUp).Row
    LASTROW = LR(TEMP2)
    
    For i = LASTROW To 1 Step -1
        If Cells(i + 1, 1) = Cells(i, 1) And _
            Cells(i + 1, 2) = Cells(i, 2) And Cells(i + 1, 3) = Cells(i, 3) Then

            Range(Cells(i + 1, 1), Cells(i + 1, 3)).ClearContents
        End If
        If Cells(i + 1, 10) <> "" Or Cells(i + 1, 10) = 1 Or Cells(i + 1, 10) = 0 Then
                Cells(i + 1, 10).NumberFormat = "0%"
        End If
    Next i
    LASTCOL = TEMP2.Range("SAM1").End(xlToLeft).Column
    
    Range(Cells(1, 1), Cells(LASTROW, LASTCOL)).Copy
    SH_STR_PIC.Activate
    Range("A1").PasteSpecial xlPasteAll
    Application.CutCopyMode = False
    'LASTROW = Range("A" & Rows.Count).End(xlUp).Row
    LASTROW = LR(SH_STR_PIC)
    LASTCOL = Range("SAM1").End(xlToLeft).Column
    
    Rows(1).Font.Size = 14
    
    Rows(1).RowHeight = 25
    Cells.HorizontalAlignment = xlCenter
    Cells.VerticalAlignment = xlCenter

    With Range(Cells(1, 1), Cells(1, LASTCOL))
        .Font.Name = "Century Gothic"
        .Font.Color = vbWhite
        .Font.Bold = True
        .Interior.Color = RGB(52, 98, 101)
        .Interior.Pattern = xlSolid
        .Interior.PatternColorIndex = xlAutomatic
    End With
    
    Rows("2:" & LASTROW).Font.Name = "Verdana"
    Rows("2:" & LASTROW).Font.Size = 11.5
        
    Range(Cells(2, 1), Cells(LASTROW, LASTCOL)).Borders.LineStyle = xlNone
    Range(Cells(2, 1), Cells(LASTROW, LASTCOL)).Interior.Color = xlNone
    If ActiveWindow.FreezePanes = True Then ActiveWindow.FreezePanes = False
    
'    LASTCOL = LASTCOL - 1
    LASTROW = LR(SH_STR_PIC)
    
    x = 2
'    k = 0
    Dim var1 As String
    Dim var2 As String
    
    For i = 2 To LASTROW
        
'        If i = LASTROW Then Stop
        If Cells(i, 1).Value <> "" And i > 2 Or i = LASTROW Then
            k = k + 1
            j = i - 1
            If i = LASTROW Then
                var1 = Cells(i - 1, 1) & "-" & Cells(i - 1, 2) & "-" & Cells(i - 1, 3)
                var2 = Cells(i, 1) & "-" & Cells(i, 2) & "-" & Cells(i, 3)
                If var1 <> var2 Then
                    x = LASTROW
                    j = LASTROW
                    k = k + 1
                End If
            End If
            If k Mod 2 = 0 Then
                With Range(Cells(x, 1), Cells(j, LASTCOL))
                    .Interior.Pattern = xlSolid
                    .Interior.PatternColor = xlAutomatic
                    .Interior.Color = RGB(255, 255, 255)
                End With
            ElseIf k Mod 2 <> 0 Then
                With Range(Cells(x, 1), Cells(j, LASTCOL))
                    .Interior.Pattern = xlSolid
                    .Interior.PatternColor = xlAutomatic
                    .Interior.Color = RGB(228, 240, 241)
                End With
            End If
            With Range(Cells(x, 1), Cells(j, LASTCOL))
                .Borders(xlEdgeTop).LineStyle = xlContinuous
                .Borders(xlEdgeTop).Color = RGB(52, 98, 101)
                .Borders(xlEdgeTop).TintAndShade = 0
                .Borders(xlEdgeTop).Weight = xlThin
                .Borders(xlEdgeBottom).LineStyle = xlContinuous
                .Borders(xlEdgeBottom).Color = RGB(52, 98, 101)
                .Borders(xlEdgeBottom).TintAndShade = 0
                .Borders(xlEdgeBottom).TintAndShade = 0
            End With
            x = j + 1
        End If
    Next i
    With Range("A:L")
        .Borders(xlEdgeLeft).LineStyle = xlContinuous
        .Borders(xlEdgeLeft).ThemeColor = 1
        .Borders(xlEdgeLeft).TintAndShade = 0
        .Borders(xlEdgeLeft).Weight = xlThin
        .Borders(xlInsideVertical).LineStyle = xlContinuous
        .Borders(xlInsideVertical).ThemeColor = 1
        .Borders(xlInsideVertical).TintAndShade = 0
        .Borders(xlInsideVertical).Weight = xlThin
        .Borders(xlEdgeRight).LineStyle = xlContinuous
        .Borders(xlEdgeRight).ThemeColor = 1
        .Borders(xlEdgeRight).TintAndShade = 0
        .Borders(xlEdgeRight).Weight = xlThin
    End With
    
    Range("L1").FormulaR1C1 = "Performance PIC"
    Range("L2").FormulaR1C1 = "=AVERAGE(C[-1])"
    Range("A1").Select
    With Range("L1")
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Bold = True
        .Font.Name = "Century Gothic"
        .Font.Color = vbWhite
        .Interior.Pattern = xlSolid
        .Interior.PatternColor = xlAutomatic
        .Interior.Color = RGB(52, 98, 101)
    End With
    With Range("L2")
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Bold = True
        .Font.Name = "Century Gothic"
        .Interior.Pattern = xlSolid
        .Interior.PatternColor = xlAutomatic
        .Interior.Color = RGB(255, 255, 255)
        .NumberFormat = "[=0]0%;[<1]0.0%;0%;_(@_)"
    End With

    Rows(2).Insert
    Range(Cells(2, 1), Cells(2, LASTCOL + 1)).Interior.Pattern = xlSolid
    Range(Cells(2, 1), Cells(2, LASTCOL + 1)).Interior.PatternColor = xlAutomatic
    Range(Cells(2, 1), Cells(2, LASTCOL + 1)).Interior.Color = RGB(0, 50, 72)
    Rows(2).RowHeight = 5
    ActiveWindow.Zoom = 85
    Rows(3).Select
    ActiveWindow.FreezePanes = True
    Cells.EntireColumn.AutoFit
    Range(Cells(1, 1), Cells(LASTROW + 1, LASTCOL + 1)).AutoFilter
    Cells(1, 1).Select
    Cells.HorizontalAlignment = xlCenter
    Cells.VerticalAlignment = xlCenter
    Cells.EntireColumn.AutoFit
Next i_PIC
'Stop
For i = 1 To 5
    If WorksheetExists("TEMP" & i) Then Sheets("TEMP" & i).Delete
Next i
If WorksheetExists("S_PIC") Then Sheets("S_PIC").Delete
If WorksheetExists("S_PO") Then Sheets("S_PO").Delete

If SH_TARIKAN.AutoFilterMode = True Then SH_TARIKAN.AutoFilterMode = False
If SH_CEK.AutoFilterMode = True Then SH_CEK.AutoFilterMode = False

Dim SH_RESUME As Worksheet
If WorksheetExists("Resume PIC") Then Sheets("Resume PIC").Delete
Set SH_RESUME = Sheets.Add(After:=Sheets(Sheets.Count)): ActiveSheet.Name = "Resume PIC"
With SH_RESUME.Tab
    .Color = 65535
    .TintAndShade = 0
End With
SH_RESUME.Activate

str_Month = WorksheetFunction.Text(SH_TARIKAN.Range("P2").Value, "[$-id-ID]mmmm")
str_Title = "WO PURCHASING - PERFORMANCE PER PIC " & "(" & UCase(str_Month) & ")"
Range("A1") = "NO": Range("B1") = "PIC": Range("C1") = str_Month
For i = 4 To Sheets.Count - 1
    SH_RESUME.Range("A" & i - 2) = i - 3
    SH_RESUME.Range("B" & i - 2) = Sheets(i).Name
    Sheets(i).Range("L3").Copy
    SH_RESUME.Range("C" & i - 2).PasteSpecial xlPasteValuesAndNumberFormats: Application.CutCopyMode = False
Next i

Call CreateDataSource

LASTCOL = Range("SAM1").End(xlToLeft).Column
LASTROW = Range("A1000").End(xlUp).Row
Cells.HorizontalAlignment = xlCenter
Cells.VerticalAlignment = xlCenter
Cells.EntireColumn.AutoFit

Set DATA_RANGE = Range(Cells(1, 1), Cells(LASTROW, LASTCOL))
With DATA_RANGE
    .Borders.LineStyle = xlContinuous
    .Rows(1).Font.Bold = True
    .Rows(1).Interior.Color = RGB(217, 217, 217)
    .HorizontalAlignment = xlCenter
    .VerticalAlignment = xlCenter
    .EntireColumn.AutoFit
    .Columns(2).Offset(1).HorizontalAlignment = xlLeft
End With
Rows("1:3").Insert
Range("A:B").Insert
LASTCOL = LASTCOL + 2
With Range(Cells(2, 2), Cells(2, LASTCOL + 1))
    .Merge
    .Value = str_Title
    .Font.Bold = True
    .Font.Size = 12
    .HorizontalAlignment = xlCenter
    .VerticalAlignment = xlCenter
    .Interior.Color = RGB(217, 217, 217)
    .Interior.Pattern = xlSolid
    .Interior.PatternColorIndex = xlAutomatic
    .RowHeight = 20
End With
Range("A1").Select
Range("A:A").ColumnWidth = 5
Range("B:B,F:F").ColumnWidth = 13

Cells(1, 1).Select

'[X].....
'[SAVE DATA RESUME PER BULAN]...
path_Monthly = HOME.Range("C15") & Application.PathSeparator & str_Month & ".xlsx"
ReDim arr_Sheets(1 To TWB.Sheets.Count - 3)
For i = 4 To TWB.Sheets.Count
    arr_Sheets(i - 3) = TWB.Sheets(i).Name
Next i

Sheets(arr_Sheets()).Copy
Set wb_Monthly = ActiveWorkbook
wb_Monthly.Activate
Sheets("Resume PIC").Move Before:=Sheets(1)
Sheets("Resume PIC").Select: Cells(1, 1).Select
wb_Monthly.SaveAs Filename:=path_Monthly, FileFormat:=xlOpenXMLWorkbook
wb_Monthly.Close (True)

SH_TARIKAN.Activate: Cells(1, 1).Select
SH_CEK.Activate: Cells(1, 1).Select


HOME.Activate
Cells(1, 1).Select

Application.DisplayAlerts = True

End Sub

Sub GeneratePeriodeGCC()

Application.DisplayAlerts = False

Set TWB = ThisWorkbook

For i = TWB.Sheets.Count To 2 Step -1
    Sheets(i).Delete
Next i

firstDate = DateSerial(Year(Date), Month(Date), 1)
lastDate = DateSerial(Year(Date), Month(Date) + 1, 1 - 1)

If WorksheetExists("PERIODE") Then Sheets("PERIODE").Delete
Set shPeriode = Sheets.Add(After:=Sheets(Sheets.Count)): ActiveSheet.Name = "PERIODE"
shPeriode.Activate
Range("A1").Value = "PERIODE AWAL"
Range("A2").Value = firstDate

Range("B1").Value = "PERIODE AKHIR"
Range("B2").Value = lastDate
Cells.EntireColumn.AutoFit: Cells(1, 1).Select

HOME.Activate
Cells(1, 1).Select

Application.DisplayAlerts = True
End Sub

Sub GetFileAndPO()

Application.DisplayAlerts = False

Set TWB = ThisWorkbook

For i = TWB.Sheets.Count To 2 Step -1
    Sheets(i).Delete
Next i

PATH_TARIKAN = HOME.Range("C8").Value & _
                Application.PathSeparator & _
                HOME.Range("D8").Value & _
                HOME.Range("E8").Value
                
If Dir(PATH_TARIKAN) = "" Then
    MsgBox "File " & HOME.Range("D8") & " Doesn't Exists", vbCritical, "File Not Found"
    Exit Sub
End If

If WorksheetExists("TARIKAN") Then Sheets("TARIKAN").Delete
If WorksheetExists("PO") Then Sheets("PO").Delete

Set SH_TARIKAN = Sheets.Add(After:=Sheets(Sheets.Count)): ActiveSheet.Name = "TARIKAN"
Set SH_PO = Sheets.Add(After:=Sheets(Sheets.Count)): ActiveSheet.Name = "S_PO"

'[AMBIL DATA TARIKAN GCC]
'.........................................................................................

Set WB_TARIKAN = Workbooks.Open(PATH_TARIKAN)
WB_TARIKAN.Activate: Sheets(1).Select

If ActiveSheet.AutoFilterMode = True Then ActiveSheet.AutoFilterMode = False
With Columns("S")
    .Replace what:="@CLN", Replacement:=""
    .Replace what:="SULI.R", Replacement:="SULI"
    .Replace what:="GITA", Replacement:="GITA WIDIANTI"
    .Replace what:="RPA10", Replacement:="SUHARTATI"
    .Replace what:="RPA_PC1", Replacement:="SILVIA"
    .Replace what:="RPA8", Replacement:="SILVIA"
End With
Range("A:A").Insert
Range("A1") = "PO"
With Range("A2:A" & LR)
    .FormulaR1C1 = "=LEFT(RC[2],8)"
    .Copy
    .PasteSpecial xlPasteValuesAndNumberFormats: Application.CutCopyMode = False: Cells(1, 1).Select
End With
Range("A:A").Insert
Range("A1") = "Do Ty"
With Range("A2:A" & LR)
    .FormulaR1C1 = "=MID(RC[3],10,2)"
    .Copy
    .PasteSpecial xlPasteValuesAndNumberFormats: Application.CutCopyMode = False: Cells(1, 1).Select
End With
Range("A:A").Insert
Range("A1") = "BU"
With Range("A2:A" & LR)
    .FormulaR1C1 = "=RIGHT(RC[4],4)"
    .Copy
    .PasteSpecial xlPasteValuesAndNumberFormats: Application.CutCopyMode = False: Cells(1, 1).Select
End With
Range("AJ1") = "PO-OR.No-Item.Short-BU-Receipt.Due.Date"
With Range("AJ2:AJ" & LR)
    .FormulaR1C1 = "=RC[-33]&""-""&RC[-27]&""-""&RC[-29]&""-""&RC[-35]&""-""&RC[-19]"
    .Copy
    .PasteSpecial xlPasteValuesAndNumberFormats: Application.CutCopyMode = False: Cells(1, 1).Select
End With

Range("AK1") = "Used1"
With Range("AK2:AK" & LR)
    .FormulaR1C1 = "=COUNTIF(R1C36:RC[-1],RC[-1])"
    .Copy
    .PasteSpecial xlPasteValuesAndNumberFormats: Application.CutCopyMode = False: Cells(1, 1).Select
End With

Range("AL1") = "LOOKUP"
With Range("AL2:AL" & LR)
    .FormulaR1C1 = "=MID(RC[-2],1,LEN(RC[-2])-11)"
    .Copy
    .PasteSpecial xlPasteValuesAndNumberFormats: Application.CutCopyMode = False: Cells(1, 1).Select
End With

Range("AM1") = "Used2"
With Range("AM2:AM" & LR)
    .FormulaR1C1 = "=COUNTIF(R1C38:RC[-1],RC[-1])"
    .Copy
    .PasteSpecial xlPasteValuesAndNumberFormats: Application.CutCopyMode = False: Cells(1, 1).Select
End With

Range("AN1") = "PO-OR-ITEM-COUNT"
With Range("AN2:AN" & LR)
    .FormulaR1C1 = "=CONCATENATE(RC[-2],""-"",RC[-1])"
    .Copy
    .PasteSpecial xlPasteValuesAndNumberFormats: Application.CutCopyMode = False: Cells(1, 1).Select
End With

If ActiveSheet.AutoFilterMode = True Then ActiveSheet.AutoFilterMode = False
Range("SAM1").Copy
Columns("C:C").PasteSpecial Paste:=xlPasteAll, Operation:=xlAdd, SkipBlanks:= _
    False, Transpose:=False
Range("SAM1").Copy
Columns("A:A").PasteSpecial Paste:=xlPasteAll, Operation:=xlAdd, SkipBlanks:= _
    False, Transpose:=False
Application.CutCopyMode = False

Range("A1:AL" & LR).AutoFilter 22, "<>NSITI", xlAnd, "<>"
Range(Cells(1, 1), Cells(LR, LC)).SpecialCells(xlCellTypeVisible).Copy SH_TARIKAN.Range("A1")
If ActiveSheet.AutoFilterMode = True Then ActiveSheet.AutoFilterMode = False

WB_TARIKAN.Close False

SH_TARIKAN.Activate: Cells.EntireColumn.AutoFit: Cells(1, 1).Select
Range("C1:C" & Range("C" & Rows.Count).End(xlUp).Row).Copy SH_PO.Range("A1")

'[Preprocessing PO]
'--------------------------------------------------------------------------------
SH_PO.Activate
Range("A:A").RemoveDuplicates 1, xlYes
Range("SAM1").Copy
Columns("A:A").PasteSpecial Paste:=xlPasteAll, Operation:=xlAdd, SkipBlanks:= _
    False, Transpose:=False
Application.CutCopyMode = False

Range("B1").Value = "MIN"
Range("C1").Value = "MAX"
Range("D1").Value = "Periode Awal"
Range("E1").Value = "Periode Akhir"

'namaBulan = HOME.Range("D8").Value
namaBulan = Format(Date, "MMMM")
monthNumber = NomorBulan(namaBulan)
tahun = Year(Application.WorksheetFunction.Max(SH_TARIKAN.Range("Q:Q")))
'firstDate = DateSerial(tahun - 1, monthNumber, 1) '1 TAHUN KEBELAKANG
firstDate = DateAdd("m", -6, Date)
firstDate = DateSerial(Year(firstDate), Month(firstDate), 1) '6 BULAN KEBELAKANG
lastDate = TanggalTerakhirBulan(monthNumber, tahun)

'[Dapatkan PO MIN ~ MAX]
'---------------------------------------------------------------------------------
Range("B2") = Application.WorksheetFunction.Min(Range("A:A")) 'MIN
Range("C2") = Application.WorksheetFunction.Max(Range("A:A")) 'MAX

'[Dapatkan Periode]
'---------------------------------------------------------------------------------
Range("D2").Value = firstDate '"Periode Awal"
Range("E2").Value = lastDate  '"Periode Akhir"

Range("A:A").Delete
Cells.EntireColumn.AutoFit: Cells(1, 1).Select
HOME.Activate: Cells(1, 1).Select

Application.DisplayAlerts = True

End Sub

Sub CreateDataSource()

    Application.DisplayAlerts = False
    
    Dim sh_Src As Worksheet
    Dim wb_Src As Workbook, src_Name As String
    Dim src_Path As String
    
    Set TWB = ThisWorkbook
    Set SH_TARIKAN = TWB.Sheets("TARIKAN")
    Set sh_Src = TWB.Sheets("Resume PIC")
    
    folderYear = Format(Date, "yyyy")
    src_Name = WorksheetFunction.Text(SH_TARIKAN.Range("P2").Value, "[$-id-ID]mmmm")
    
    '[*]..PERIKSA APAKAH FOLDER PER TAHUN SUDAH ADA
    '.....JIKA BELUM MAKA BUATKAN TERLEBIH DAHULU
    folderPath = PathSRC & Application.PathSeparator & folderYear
    
    If Dir(folderPath, vbDirectory) = "" Then
        MkDir folderPath
    End If
    
    src_Path = folderPath & Application.PathSeparator & src_Name & ".xlsx"
    
    sh_Src.Copy
    Set wb_Src = ActiveWorkbook
    Cells.EntireColumn.AutoFit
    wb_Src.SaveAs Filename:=src_Path, FileFormat:=xlOpenXMLWorkbook
    wb_Src.Close (True)
    
    sh_Src.Activate: Cells(1, 1).Select
    
End Sub



