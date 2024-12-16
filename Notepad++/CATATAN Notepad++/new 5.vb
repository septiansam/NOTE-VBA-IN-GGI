Sub PROSES1()

Application.DisplayAlerts = False

'''inisialisasi(validasi)'''
Set Twb = ThisWorkbook
Set Home = Twb.Sheets("home")
Set Wf = Application.WorksheetFunction

PathFile = Home.Range("D6")
Ext = Home.Range("F6")

SumFile = Wf.CountA(Home.Range("H:H")) - 1
For I = 1 To SumFile
    StrFile = Home.Cells(I + 5, 8)
    If Dir(PathFile & "\" & StrFile & Ext) = vbNullString Then
        MsgBox "File " & StrFile & " Doesn't Exists", vbInformation, "File Not Found"
        Home.Activate: Cells(1, 1).Select
        Exit Sub
    End If
Next I

'If Home.Range("AB6") = "DONE" Then
'    MsgBox "PROCESS 1 HAS BEEN COMPLETED", vbCritical, "PROCESS 1 DONE RUNNING"
'    Exit Sub
'End If

For I = Sheets.Count To 2 Step -1
    Sheets(I).Delete
Next I

'Home.Range("AC6") = ""
DateProcessed = DateAdd("m", -1, Date)
StrMonth = Wf.Text(DateProcessed, "[$-id-ID]mmmm")
StrYear = Format(DateProcessed, "yyyy")

StrHisResume = "HISTORY RESUME " & StrYear

'---Create Folder Resume---
'````````````````````````````
StrResume = Home.Range("E7")
FolderResume = Twb.Path & "\" & "RESUME" & "\"

If Dir(FolderResume & StrYear, vbDirectory) = "" Then
    MkDir FolderResume & StrYear
End If
Ext = Range("F7")

Home.Hyperlinks.Add _
    Anchor:=Home.Range("D7"), _
    Address:=FolderResume & StrYear, _
    TextToDisplay:=FolderResume & StrYear
    
PathResume = FolderResume & StrYear & "\" & StrResume & Ext

'___________________________
'---Create Folder History---
'````````````````````````````

'...[History Resume]
FolderHistoryResume = Home.Range("Z6") & "\"
If Dir(FolderHistoryResume & StrYear, vbDirectory) = "" Then
    MkDir FolderHistoryResume & StrYear
End If
Ext = Range("E14")
PathHistoryResume = FolderHistoryResume & StrYear & "\" & UCase(StrHisResume) & Ext

''...[History Results]
'FolderHistoryResults = Home.Range("D15") & "\"
'If Dir(FolderHistoryResume & StrYear, vbDirectory) = "" Then
'    MkDir FolderHistoryResume & StrYear
'End If
'Ext = Range("F14")
'PathHistoryResume = FolderHistoryResume & StrYear & "\" & StrHisResume & Ext

Set Data = Sheets.Add(AFTER:=Sheets(Sheets.Count)): Data.Name = "DATA"

'''import file'''
For I = 1 To SumFile
    If I = 1 Then
        FR = 1
        RowPaste = 1
    Else
        FR = 2
        RowPaste = Data.Cells.Find("*", , xlFormulas, xlPart, xlByRows, xlPrevious).Row + 1
    End If
    StrFile = Home.Cells(I + 5, 8)
    Ext = Home.Range("F6")
    Set WbFile = Workbooks.Open(PathFile & "\" & StrFile & Ext): WbFile.Activate
    ActiveSheet.AutoFilterMode = False
    LR = Range("A1000000").End(xlUp).Row
    LC = Cells.Find("*", , xlFormulas, xlPart, xlByColumns, xlPrevious).Column
    Set Rng = Range(Cells(FR, 1), Cells(LR, LC))
    Rng.Copy
    Data.Activate: Range("A" & RowPaste).Select: Data.Paste
    WbFile.Close (False)
    
    Set Rng = Nothing
    Set WbFile = Nothing
Next I

Data.Activate
LR = Cells.Find(What:="*", LookAt:=xlPart, LookIn:=xlFormulas, SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
Range("A:A").Insert: Range("A1").Value = "Months"
With Range("A2:A" & LR)
    .FormulaR1C1 = _
        "=IFERROR(Proper(LEFT(TEXT(RC[5], ""[$-id-ID]mmmm""),3)), ""Month Not Found"")"
    .Copy
    .PasteSpecial xlPasteValuesAndNumberFormats
End With
Range("A:A").Insert: Range("A1") = "Reason"
Range("A:A").Insert: Range("A1") = "MD"
Range("A:A").Insert: Range("A1") = "Factory"
Application.CutCopyMode = False

LC = Cells.Find(What:="*", LookAt:=xlPart, LookIn:=xlFormulas, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column
Set Rng = Range(Cells(1, 1), Cells(LR, LC))

Set PV = Sheets.Add(AFTER:=Sheets(Sheets.Count)): PV.Name = "PIVOT": PV.Activate
Set PC = Twb.PivotCaches.Create(SourceType:=xlDatabase, _
                    SourceData:=Rng)
Set PT = PC.CreatePivotTable _
                    (TableDestination:=PV.Range("A1"), _
                    TableName:="PivotTable")

With PT
    
'''ROW'''
    With .PivotFields("Business Unit")
        .Orientation = xlRowField
        .Position = 1
        .Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False, False)
    End With

    With .PivotFields("Explanation -Remark-")
        .Orientation = xlRowField
        .Position = 2
        .Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False, False)
    End With

    With .PivotFields("Months")
        .Orientation = xlRowField
        .Position = 3
        .Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False, False)
    End With
    
    With .PivotFields("Factory")
        .Orientation = xlRowField
        .Position = 4
        .Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False, False)
    End With
    
    With .PivotFields("MD")
        .Orientation = xlRowField
        .Position = 5
        .Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False, False)
    End With
    
    With .PivotFields("Reason")
        .Orientation = xlRowField
        .Position = 6
        .Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False, False)
    End With

'''COLUMN'''

'''DATA FIELD'''
    With .PivotFields("LT 1 Debit")
        .Orientation = xlDataField
        .Position = 1
        .Function = xlSum
        .Caption = "Total"
    End With
    
    .RowAxisLayout xlTabularRow
    .ShowValuesRow = False
    .DisplayErrorString = False
    .DisplayNullString = True
    .NullString = ""
    .RepeatAllLabels xlRepeatLabels
End With
LR = Cells.Find(What:="*", LookAt:=xlPart, LookIn:=xlFormulas, SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
LC = Cells.Find(What:="*", LookAt:=xlPart, LookIn:=xlFormulas, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column
Set Rng = Range(Cells(1, 1), Cells(LR, LC))
With Rng
    .Copy
    .PasteSpecial xlPasteValuesAndNumberFormats
    .Replace "(blank)", "", xlPart
End With
Range(Cells(2, LC), Cells(LR, LC)).NumberFormat = _
        "_($* #,##0.00_);[Red]_($* #,##0.00);_($* ""-""??_);_(@_)"

Range("SAM1").Copy
Range(Cells(2, 1), Cells(LR - 1, 1)).PasteSpecial Paste:=xlPasteAll, Operation:=xlAdd
Application.CutCopyMode = False

Data.Activate
Range("A:D").Delete: Cells.EntireColumn.AutoFit: Cells(1, 1).Select
PV.Activate

Cells.EntireColumn.AutoFit
Cells.EntireRow.AutoFit
Cells.HorizontalAlignment = xlCenter
Cells.VerticalAlignment = xlCenter
Range("B2:B" & LR).HorizontalAlignment = xlLeft
Rng.Borders.LineStyle = xlContinuous
For Each Cell In Rng.Columns
    Cell.ColumnWidth = Cell.ColumnWidth + 2
Next Cell
Rows(1).Insert
Range("A1").Value = "BY AIR"
With Range(Cells(1, 1), Cells(1, LC))
    .Merge
    .Font.Bold = True
    .Cells.HorizontalAlignment = xlCenter
End With

With Range(Cells(2, 1), Cells(2, LC))
    .Font.Bold = True
    .Interior.Color = RGB(166, 166, 166)
    .Interior.Pattern = xlSolid
    .Interior.PatternColorIndex = xlAutomatic
    .RowHeight = .RowHeight + 3
End With
Cells(1, 1).Select

PV.Name = "RESUME"
Home.Range("AA6") = UCase(StrMonth)

'CREATE HYPERLINK AND SAVE HISTORY RESUME
Home.Activate
Home.Hyperlinks.Add _
    Anchor:=Home.Range("D14"), _
    Address:=PathHistoryResume, _
    TextToDisplay:=PathHistoryResume
    
If Dir(PathHistoryResume) = "" Then
    PV.Copy
    Set WbHisResume = ActiveWorkbook
    WbHisResume.SaveAs PathHistoryResume, xlOpenXMLWorkbook
    WbHisResume.Close (False)
    
    PV.Copy
    Set WbResume = ActiveWorkbook
    Cells(1, 1).Select
    WbResume.SaveAs PathResume, xlOpenXMLWorkbook
    WbResume.Close (False)
Else
    Set Temp = Sheets.Add(AFTER:=Sheets(Sheets.Count)): Temp.Name = "TEMP"
    Set WbHisResume = Workbooks.Open(PathHistoryResume): WbHisResume.Activate: Sheets(1).Activate
    Cells.Copy
    Temp.Activate: Temp.Paste: Application.CutCopyMode = False
    WbHisResume.Close (False)
    Temp.Activate
    Cells.Borders.LineStyle = xlNone
    LR = Cells.Find(What:="*", LookAt:=xlPart, LookIn:=xlFormulas, SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
    Rows(LR).Delete
    PV.Activate
    Cells.Borders.LineStyle = xlNone
    LR = Cells.Find(What:="*", LookAt:=xlPart, LookIn:=xlFormulas, SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
    LC = Cells.Find(What:="*", LookAt:=xlPart, LookIn:=xlFormulas, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column
    Set Rng = Range(Cells(3, 1), Cells(LR - 1, LC))
    Rng.Copy: Temp.Activate
    Range("A" & Rows.Count).End(xlUp).Offset(1).PasteSpecial xlPasteAll
    Application.CutCopyMode = False
    LR = Cells.Find(What:="*", LookAt:=xlPart, LookIn:=xlFormulas, SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
    LC = Cells.Find(What:="*", LookAt:=xlPart, LookIn:=xlFormulas, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column
    Temp.Range(Cells(2, 1), Cells(LR, LC)).RemoveDuplicates Columns:=Array(1, 2, 3, 7) _
        , Header:=xlYes
    Cells(1, 1).Select
    LR = Cells.Find(What:="*", LookAt:=xlPart, LookIn:=xlFormulas, SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
    LR = LR + 1
    Range("A" & LR) = "Grand Total"
    With Cells(LR, LC)
        .FormulaR1C1 = "=SUM(R[-" & (LR - 3) & "]C:R[-1]C)"
        .Copy
        .PasteSpecial xlPasteValuesAndNumberFormats: Application.CutCopyMode = False
    End With
    Range(Cells(3, LC), Cells(LR, LC)).NumberFormat = _
        "_($* #,##0.00_);[Red]_($* #,##0.00);_($* ""-""??_);_(@_)"
    Cells.EntireColumn.AutoFit
    Set Rng = Range(Cells(2, 1), Cells(LR, LC))
    Rng.Borders.LineStyle = xlContinuous
    For Each Cell In Rng.Columns
        Cell.ColumnWidth = Cell.ColumnWidth + 2
    Next Cell
    Cells(1, 1).Select
    PV.Delete
    Temp.Name = "RESUME"
    
    Temp.Copy
    Set WbHisResume = ActiveWorkbook
    WbHisResume.SaveAs PathHistoryResume, xlOpenXMLWorkbook
    WbHisResume.Close (False)
    
    Temp.Copy
    Set WbResume = ActiveWorkbook
    Cells(1, 1).Select
    WbResume.SaveAs PathResume, xlOpenXMLWorkbook
    WbResume.Close (False)
End If

Home.Activate
Home.Hyperlinks.Add _
    Anchor:=Home.Range("AB6"), _
    Address:=PathResume, _
    TextToDisplay:=PathResume
Cells(1, 1).Select

Application.DisplayAlerts = True

End Sub