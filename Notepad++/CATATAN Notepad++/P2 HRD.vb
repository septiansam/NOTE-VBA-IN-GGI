'
'
'
'
Public Function WorksheetExists(shtName As String) As Boolean
    On Error Resume Next
    WorksheetExists = Not Sheets(shtName) Is Nothing
    On Error GoTo 0
End Function

'
' UPDATE PROSES1 SEPTIAN

Sub NEW_PROSES1()

Dim TWB As Workbook
Dim SH_DB As Worksheet, SH_InputUser As Worksheet, SH_InputRpa As Worksheet
Dim SH_STATUS As Worksheet, SH_BANTU As Worksheet
Dim LR As Long, LC As Long, WS As Worksheet, i As Long
Dim SUM_SH_RUN As Long
Dim AN As String
Dim SUMSUPLY As Long
Dim TANGGAL As Date
Dim SH_RUN As Worksheet, STR_SH_RUN As String, KEY As Variant

Application.DisplayAlerts = False

Set TWB = ThisWorkbook
Set SH_DB = TWB.Sheets("Database Supplier")
Set SH_InputUser = TWB.Sheets("Input User")

If WorksheetExists("Status") Then Sheets("Status").Delete
Set SH_STATUS = Sheets.Add(After:=Sheets(Sheets.Count)): ActiveSheet.Name = "Status"
SH_STATUS.Range("A1").Value = "RPA Code"
SH_STATUS.Range("A2").Value = 1

For Each WS In TWB.Worksheets
    If WS.Name <> "Database Supplier" And _
        WS.Name <> "Input User" And _
        WS.Name <> "Status" Then
        
        WS.Delete
        
    End If
Next WS

Sheets("Input User").Select
Rows(1).Font.Bold = True
Range("N:N").NumberFormat = "@"
If Range("N" & Rows.Count).End(xlUp).Value <> "Account Number" Then
    For i = 2 To Range("N" & Rows.Count).End(xlUp).Row
        If Left(Cells(i, 14), 1) = "'" Then
            AN = Right(Cells(i, 14), Len(Cells(i, 14)) - 1)
            Cells(i, 14).Value = AN
        End If
    Next i
End If

Sheets("Input User").Select
Range("O" & Rows.Count).End(xlUp).Select
SUMSUPLY = ActiveCell.Row
For i = SUMSUPLY To 2 Step -1
    If Cells(i, 1) = Cells(i - 1, 1) And _
        Cells(i, 2) = Cells(i - 1, 2) And _
        Cells(i, 3) = Cells(i - 1, 3) And _
        Cells(i, 4) = Cells(i - 1, 4) And _
        Cells(i, 5) = Cells(i - 1, 5) And _
        Cells(i, 6) = Cells(i - 1, 6) And _
        Cells(i, 27) = Cells(i - 1, 27) Then
        
        Range(Cells(i, 1), Cells(i, 6)).ClearContents
        
    End If
Next i

For i = 2 To SUMSUPLY
    If Cells(i, 15) = "N" Then Range(Cells(i, 16), Cells(i, 17)) = ""
Next i

For i = 1 To SUMSUPLY
If Cells(1 + i, 5) <> "" Then
    TANGGAL = Cells(1 + i, 5)
    Cells(1 + i, 5) = TANGGAL
End If
If Cells(1 + i, 6) <> "" Then
    TANGGAL = Cells(1 + i, 6)
    Cells(1 + i, 6) = TANGGAL
End If
Next i

If WorksheetExists("Input RPA") Then Sheets("Input RPA").Delete
If WorksheetExists("Bantuan") Then Sheets("Bantuan").Delete

Set SH_InputRpa = Sheets.Add(After:=Sheets(Sheets.Count)): ActiveSheet.Name = "Input RPA"
Set SH_BANTU = Sheets.Add(After:=Sheets(Sheets.Count)): ActiveSheet.Name = "Bantuan"

SH_BANTU.Activate: Rows(1).Font.Bold = True
Range("A1") = "NO"
Range("B1") = "Nama Sheet"
Range("C1") = "OW Number"
Range("D1") = "Address Number"
Range("E1") = "Ship To"
Range("F1") = "Currency"
Range("G1") = "Requested"
Range("H1") = "Promised Delivery"
Range("I1") = "PO Number"
Range("J1") = "Status"
Range("K1") = "Transaction ID"

SH_InputUser.Activate
LR = SH_InputUser.Cells.Find(What:="*" _
    , LookAt:=xlPart _
    , LookIn:=xlFormulas _
    , SearchOrder:=xlByRows _
    , SearchDirection:=xlPrevious).Row
    
Range("B1:T" & LR).Copy SH_InputRpa.Range("B1")
Range("AA1:AA" & LR).Copy SH_InputRpa.Range("U1")

SH_InputRpa.Activate: Range("A1") = "OW": Range("A1").Select
Rows(1).Font.Bold = True

SH_InputUser.Activate
Range("AA1:AA" & LR).Copy Range("AZ1"): Range("AZ1").Delete xlUp: Range("AZ1").Select
Range("AZ:AZ").RemoveDuplicates 1, xlNo
SUM_SH_RUN = Range("AZ" & Rows.Count).End(xlUp).Row


For i = 1 To SUM_SH_RUN
    SH_BANTU.Activate
    STR_SH_RUN = "SheetRun" & i
    Cells(i + 1, 1) = i
    Cells(i + 1, 2) = STR_SH_RUN
    If WorksheetExists(STR_SH_RUN) Then Sheets(STR_SH_RUN).Delete
    Set SH_RUN = Sheets.Add(After:=Sheets(Sheets.Count)): ActiveSheet.Name = STR_SH_RUN
    SH_InputUser.Activate
    KEY = SH_InputUser.Cells(i, 52)
    If SH_InputUser.AutoFilterMode = True Then SH_InputUser.AutoFilterMode = False
    Range("A:AO").AutoFilter 27, KEY
    Range("B1:Q" & LR).SpecialCells(xlCellTypeVisible).Copy
    SH_RUN.Activate
    Range("A1").PasteSpecial xlPasteValuesAndNumberFormats
    Application.CutCopyMode = False: Cells.EntireColumn.AutoFit
    Range("A2:E2").Copy
    SH_BANTU.Activate
    Range("D" & Rows.Count).End(xlUp).Offset(1).PasteSpecial xlPasteValuesAndNumberFormats
    Range("K" & Rows.Count).End(xlUp).Offset(1).Value = KEY
    Range("A1").Select
    Application.CutCopyMode = False: Cells.EntireColumn.AutoFit
    SH_RUN.Activate
    Range("A:E").Delete: Range("A1").Select
    Rows(1).Font.Bold = True
Next i

If WorksheetExists("Status") Then Sheets("Status").Delete
Set SH_STATUS = Sheets.Add(After:=Sheets(Sheets.Count)): ActiveSheet.Name = "Status"
SH_STATUS.Activate
Range("A1") = "RPA Code"
Range("A2") = "1"

SH_InputRpa.Activate
If SH_InputRpa.AutoFilterMode = True Then SH_InputRpa.AutoFilterMode = False
Cells.EntireColumn.AutoFit: Cells(1, 1).Select
SH_InputUser.Activate
Cells(1, 1).Select
If SH_InputUser.AutoFilterMode = True Then SH_InputUser.AutoFilterMode = False
Cells.EntireColumn.AutoFit: Cells(1, 1).Select
Range("AZ:AZ").Delete

SH_DB.Activate: Cells(1, 1).Select

End Sub

'[NEW PROSES2]
'............................................................

Sub NEW_PROSES2()

Dim TWB As Workbook
Dim SH_DB As Worksheet, SH_InputUser As Worksheet, SH_InputRpa As Worksheet
Dim SH_STATUS As Worksheet, SH_BANTU As Worksheet
Dim LR As Long, LC As Long, WS As Worksheet, i As Long
Dim SUM_SH_RUN As Long
Dim AN As String
Dim SUMSUPLY As Long
Dim TANGGAL As Date
Dim SH_RUN As Worksheet, STR_SH_RUN As String, KEY As Variant

Application.DisplayAlerts = False

Set TWB = ThisWorkbook
Set SH_DB = TWB.Sheets("Database Supplier")
Set SH_InputUser = TWB.Sheets("Input User")
Set SH_BANTU = TWB.Sheets("Bantuan")
Set SH_InputRpa = TWB.Sheets("Input RPA")

For i = TWB.Sheets.Count To 5 Step -1
    Sheets(i).Delete
Next i

SH_InputRpa.Activate
LR = SH_InputRpa.Cells.Find(What:="*" _
    , LookAt:=xlPart _
    , LookIn:=xlFormulas _
    , SearchOrder:=xlByRows _
    , SearchDirection:=xlPrevious).Row

With Range("A2:A" & LR)
    .FormulaR1C1 = _
        "=IF(RC[1]<>"""",IF(INDEX(Bantuan!C[2],MATCH('Input RPA'!RC[20],Bantuan!C[10],0))=0,"""",INDEX(Bantuan!C[2],MATCH('Input RPA'!RC[20],Bantuan!C[10],0))),"""")"
    .Copy
    .PasteSpecial xlPasteValuesAndNumberFormats: Application.CutCopyMode = False: Cells.EntireColumn.AutoFit
End With
Cells(1, 1).Select

SH_InputUser.Activate
LR = SH_InputUser.Cells.Find(What:="*" _
    , LookAt:=xlPart _
    , LookIn:=xlFormulas _
    , SearchOrder:=xlByRows _
    , SearchDirection:=xlPrevious).Row
    
With Range("V2:V" & LR)
    .FormulaR1C1 = _
        "=IFERROR(IF(RC[-21]<>"""",IF(INDEX(Bantuan!C[-19],MATCH('Input User'!RC[5],Bantuan!C[-11],0))=0,"""",INDEX(Bantuan!C[-19],MATCH('Input User'!RC[5],Bantuan!C[-11],0))),""""),"""")"
    .Copy
    .PasteSpecial xlPasteValuesAndNumberFormats: Application.CutCopyMode = False: Cells.EntireColumn.AutoFit
End With

With Range("W2:W" & LR)
    .FormulaR1C1 = _
        "=IFERROR(IF(RC[-22]<>"""",IF(INDEX(Bantuan!C[-14],MATCH('Input User'!RC[4],Bantuan!C[-12],0))=0,"""",INDEX(Bantuan!C[-14],MATCH('Input User'!RC[4],Bantuan!C[-12],0))),""""),"""")"
    .Copy
    .PasteSpecial xlPasteValuesAndNumberFormats: Application.CutCopyMode = False: Cells.EntireColumn.AutoFit
End With
Cells(1, 1).Select

Dim NAME_RESUME As String, PATH_RESUME As String, WB_RESUME As Workbook

SH_DB.Activate
NAME_RESUME = Range("L9").Value
PATH_RESUME = Range("L10") & Application.PathSeparator & NAME_RESUME & ".xlsx"

SH_InputUser.Copy
Application.DisplayAlerts = False
Set WB_RESUME = ActiveWorkbook
WB_RESUME.Activate
Sheets(1).Select
ActiveSheet.Name = "RESUME"

'----------------------------------------------------
Range("A:A").Insert
Range("A1").Value = "No"

LR = ActiveSheet.Cells.Find(What:="*", LookAt:=xlPart, LookIn:=xlFormulas, SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row

Dim No As Long
No = 1
For i = 2 To LR
    If Cells(i, 2) <> "" Then
        Cells(i, 1) = No
        No = No + 1
    End If
Next i

LC = ActiveSheet.Cells.Find(What:="*", LookAt:=xlPart, LookIn:=xlFormulas, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column

With Range(Cells(1, 1), Cells(1, LC))
    .Font.Name = "Century Gothic"
    .Font.Color = vbWhite
    .Font.Bold = True
    .Font.Size = 12
    .Interior.Color = RGB(52, 98, 101)
    .Interior.Pattern = xlSolid
    .Interior.PatternColorIndex = xlAutomatic
    .RowHeight = .RowHeight + 4
End With
Rows(2).Insert
With Range(Cells(2, 1), Cells(2, LC))
    .Interior.Color = 4074248
    .Interior.Pattern = xlSolid
    .Interior.PatternColorIndex = xlAutomatic
    .RowHeight = 3
End With

LR = ActiveSheet.Cells.Find(What:="*", LookAt:=xlPart, LookIn:=xlFormulas, SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
LC = ActiveSheet.Cells.Find(What:="*", LookAt:=xlPart, LookIn:=xlFormulas, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column

For i = 3 To LR
    If Cells(i - 1, 1) <> "" And Cells(i, 1) = "" Then
        Cells(i, 1) = Cells(i - 1, 1)
    End If
Next i

For i = 3 To LR
    If Cells(i, 1).Value Mod 2 = 0 Then
        With Range(Cells(i, 1), Cells(i, LC))
            .Interior.Pattern = xlSolid
            .Interior.PatternColor = xlAutomatic
            .Interior.Color = RGB(255, 255, 255)
            .RowHeight = .RowHeight + 2
        End With
    ElseIf Cells(i, 1).Value Mod 2 <> 0 Then
        With Range(Cells(i, 1), Cells(i, LC))
            .Interior.Pattern = xlSolid
            .Interior.PatternColor = xlAutomatic
            .Interior.Color = RGB(228, 240, 241)
            .RowHeight = .RowHeight + 2
        End With
    End If
Next i

For i = LR To 3 Step -1
    If Cells(i + 1, 1) = Cells(i, 1) Then
        Cells(i + 1, 1) = ""
    End If
Next i

Set RNG_RESUME = Range(Cells(1, 1), Cells(LR, LC))
For Each COL In RNG_RESUME.Columns
    COL.EntireColumn.AutoFit
    COL.ColumnWidth = COL.ColumnWidth + 1
Next COL

Cells.HorizontalAlignment = xlCenter
Cells.VerticalAlignment = xlCenter
Range("G:G").HorizontalAlignment = xlLeft
Rows("3:3").Select
ActiveWindow.FreezePanes = True
Range("A:A").Delete
'----------------------------------------------------

Cells(1, 1).Select

WB_RESUME.SaveAs PATH_RESUME, xlOpenXMLWorkbook
WB_RESUME.Close (True)
Application.DisplayAlerts = True
ThisWorkbook.Activate

SH_DB.Activate: Cells(1, 1).Select

Application.DisplayAlerts = True

End Sub


