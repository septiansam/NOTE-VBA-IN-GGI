'Option Explicit
Dim name_Sheet As String, BULAN_STR As String, TES7 As String, TES8 As String, DATATANGGAL As Integer

Sub pertanggal_Lur()

Application.ScreenUpdating = False
Dim TES9 As String, FILTER_TGL As Date
Dim LAST_TES2 As Integer, LAST_OLAH1 As Integer, j As Integer, standar As String
Dim SMail As String, DATATESS1 As Integer, judul As String, n As Integer, i As Integer, RESUMEEEE As String, LASTOLAHAN As Integer, twb As Workbook, ws1 As Worksheet, TES1 As String, TES2 As String, TES3 As String, LastRow As Integer, FILEPATHWO As String, FILEWO As String, DirFile As String, LASTROW1 As Integer, CM0 As String, rng As Range, RESUMEE As String, RPAA As String, GGG As Integer, USDD As String, SETENGAH As String

Set twb = ThisWorkbook: Set ws1 = twb.Sheets("BANTUAN")
TES7 = "TES7": TES8 = "TES8": TES9 = "TES9"

Application.DisplayAlerts = False
If Evaluate("isref('" & TES7 & "'!A1)") Then
    Sheets(TES7).Delete
End If
If Evaluate("isref('" & TES8 & "'!A1)") Then
    Sheets(TES8).Delete
End If
If Evaluate("isref('" & TES9 & "'!A1)") Then
    Sheets(TES9).Delete
End If
Sheets.Add(after:=Sheets(Sheets.Count)).Name = TES7: Sheets.Add(after:=Sheets(Sheets.Count)).Name = TES8
Sheets.Add(after:=Sheets(Sheets.Count)).Name = TES9

Sheets("OLAHAN38").Select
If ActiveSheet.AutoFilterMode = True Then Selection.AutoFilter
BULAN_STR = Format(Sheets("OLAHAN38").Cells(2, 1), "MMMM")


Cells.AutoFilter
ActiveSheet.Range("$A$1:$A$100000").AutoFilter field:=2, Criteria1:="MJ2"
Range("A1:AA100000").SpecialCells(xlCellTypeVisible).Copy Destination:=Sheets(TES9).Cells(1, 1)
 

'Sheets("OLAHAN38").Select 'FOR ALL
'If ActiveSheet.AutoFilterMode = True Then Selection.AutoFilter
'Cells.AutoFilter
'Columns("A").Copy Destination:=Sheets(TES7).Cells(1, 1)

Sheets(TES9).Columns("A").Copy Destination:=Sheets(TES7).Cells(1, 1)

Sheets(TES7).Select
Columns("A").RemoveDuplicates Columns:=1, Header:=xlYes

DATATANGGAL = Sheets(TES7).Cells(Rows.Count, 1).End(xlUp).Row

For i = 2 To DATATANGGAL
    FILTER_TGL = Sheets(TES7).Cells(i, 1)
'    Sheets("OLAHAN38").Select 'FOR ALL
    Sheets(TES9).Select
    ActiveSheet.Range("$A$1:$A$100000").AutoFilter field:=1, Criteria1:=FILTER_TGL
    Range("A1:AA100000").SpecialCells(xlCellTypeVisible).Copy Destination:=Sheets(TES8).Cells(1, 1)
    
    name_Sheet = Format(FILTER_TGL, "dd mmm yy")
    If Evaluate("isref('" & name_Sheet & "'!A1)") Then
        Application.DisplayAlerts = False: Sheets(name_Sheet).Delete: Application.DisplayAlerts = True
    End If
    
    Sheets.Add(after:=Sheets(Sheets.Count)).Name = name_Sheet
    
    Call PIVOT_TANGGAL
    
    Sheets(TES8).Cells.ClearContents
    
Next i

Call BUAT_newfile

Sheets("OLAHAN38").Select
If ActiveSheet.AutoFilterMode = True Then Selection.AutoFilter

Application.DisplayAlerts = False
If Evaluate("isref('" & TES7 & "'!A1)") Then
    Sheets(TES7).Delete
End If
If Evaluate("isref('" & TES8 & "'!A1)") Then
    Sheets(TES8).Delete
End If
If Evaluate("isref('" & TES9 & "'!A1)") Then
    Sheets(TES9).Delete
End If
Application.DisplayAlerts = False
Application.ScreenUpdating = True



End Sub

Sub PIVOT_TANGGAL()

Dim TOLAST As Integer
Dim PTable As PivotTable, twb As Workbook, LASTT As Integer, i As Integer, PCache As PivotCache, PSheet As Worksheet, PRange As Range, rng As Range, DATAHAPUS As Integer

Set twb = ThisWorkbook: Set PSheet = twb.Worksheets(name_Sheet)

Set PRange = Sheets("TES8").Range("C3").CurrentRegion
Set PCache = ActiveWorkbook.PivotCaches.Create(SourceType:=xlDatabase, SourceData:=PRange) 'UNTUK INSERT PIVOT DARI SHEET REKAP
Set PTable = PCache.CreatePivotTable(TableDestination:=PSheet.Cells(1, 1), TableName:="PivotTable") 'MEMBUAT PIVOT DI SHEET TES

Sheets(name_Sheet).Select
With PTable.PivotFields("Factory")
    .Caption = "Factory": .Orientation = xlRowField: .Position = 1
End With
With PTable.PivotFields("WO")
    .Caption = "WO": .Orientation = xlRowField: .Position = 2
End With
With PTable.PivotFields("BUYER2")
    .Caption = "Buyer2": .Orientation = xlRowField: .Position = 3
End With
With PTable.PivotFields("Line")
    .Caption = "Line": .Orientation = xlRowField
    .Position = 4
End With
With PTable.PivotFields("Item")
    .Caption = "Item": .Orientation = xlRowField
    .Position = 5
End With
With PTable.PivotFields("Style")
    .Caption = "Style": .Orientation = xlRowField
    .Position = 6
End With '
With PTable.PivotFields("FOB")
    .Caption = "FOB": .Orientation = xlRowField
    .Position = 7
End With
With PTable.PivotFields("CMT")
    .Caption = "CMT": .Orientation = xlRowField
    .Position = 8
End With
With PTable.PivotFields("CM(USD)")
    .Caption = "CM(USD)": .Orientation = xlRowField
    .Position = 9
End With

'With PTable.PivotFields("CM(USD)")
'    .Orientation = xlDataField
'    .Position = 1: .Function = xlAverage
'End With
With PTable.PivotFields("Output")
    .Orientation = xlDataField
    .Position = 1: .Function = xlSum
End With
With PTable.PivotFields("Amount.CM(USD)")
    .Orientation = xlDataField
    .Position = 2: .Function = xlSum
End With
With PTable.PivotFields("Cost.Proporsional(US)")
    .Orientation = xlDataField
    .Position = 3: .Function = xlSum
End With
With PTable.PivotFields("Profit.Lost(USD)")
    .Orientation = xlDataField
    .Position = 4: .Function = xlSum
End With

Sheets(name_Sheet).Select
ActiveSheet.PivotTables("PivotTable").RowAxisLayout xlTabularRow
ActiveSheet.PivotTables("PivotTable").PivotFields("Item").Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False, False)
ActiveSheet.PivotTables("PivotTable").PivotSelect "Line[All]", xlLabelOnly, True
ActiveSheet.PivotTables("PivotTable").PivotFields("Line").Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False, False)
ActiveSheet.PivotTables("PivotTable").PivotFields("FOB").Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False, False)
ActiveSheet.PivotTables("PivotTable").PivotFields("CMT").Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False, False)
ActiveSheet.PivotTables("PivotTable").PivotFields("Buyer2").Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False, False)
ActiveSheet.PivotTables("PivotTable").PivotFields("Style").Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False, False)
ActiveSheet.PivotTables("PivotTable").PivotFields("WO").Subtotals = Array(False, False, False, False, False, False, False, False, False, False, False, False)

Cells.Copy: Cells(1, 1).PasteSpecial xlPasteValues
Application.CutCopyMode = False

Cells.Replace what:="(blank)", Replacement:=""
Columns("G:H").Cut: Range("N1").Insert Shift:=xlToRight
    
Cells(2, 14) = "Sum of FOB": Cells(2, 15) = "Sum of CMT"

TOLAST = Range("A" & Rows.Count).End(xlUp).Row - 1
Range("N3:O" & TOLAST).Formula = "=$G3*L3"
Cells(2, 16) = "Total Sales"
Range("P3:P" & TOLAST).Formula = "=N3+O3"

Set rng = Range("A2").CurrentRegion
With rng.Borders 'BUAT BORDER
    .LineStyle = xlContinuous: .Weight = xlThin
End With

If ActiveSheet.AutoFilterMode = True Then Selection.AutoFilter
rng.AutoFilter
'ActiveSheet.Range("$A$1:$A$10000").AutoFilter Field:=2, Criteria1:="*Total*"
'RNG.Font.Bold = True
'
'With RNG.Interior
'    .Pattern = xlSolid: .PatternColorIndex = xlAutomatic
'    .ThemeColor = xlThemeColorAccent6: .TintAndShade = 0.599993896298105: .PatternTintAndShade = 0
'End With
ActiveSheet.Range("$A$1:$A$10000").AutoFilter field:=1, Criteria1:="*Total*"
rng.Font.Bold = True
Columns("N:P").ClearContents

With rng.Interior
    .Color = 15773696
End With
If ActiveSheet.AutoFilterMode = True Then Selection.AutoFilter

Rows(2).Font.Bold = True
Rows(1).Delete
Range("A1:P1").Interior.ColorIndex = 33
With Rows("1:1")
    .Font.Size = 13
    .RowHeight = 30: .VerticalAlignment = xlCenter
    .HorizontalAlignment = xlCenter
End With
Cells(1, 3) = "Buyer"

Columns("I:K").NumberFormat = "#,##0.00_);[Red](#,##0.00)"
Columns("H").NumberFormat = "#,##0"

ActiveWindow.Zoom = 75: Cells.EntireColumn.AutoFit: Cells(1, 1).Select

End Sub


Sub BUAT_newfile()
Dim ARAN_TGL As Date, ARAN_SHEET As String
Dim twb As Workbook, FILEBARU As Workbook, TGL_STR As Variant, FILECNJ As Workbook, SAVEFILE As String, ws1 As Worksheet, saveee As Variant, i As Integer, n As Integer

Set twb = ThisWorkbook
Set ws1 = twb.Sheets("BantuAN")
Set FILECNJ = ActiveWorkbook

saveee = "\\10.8.0.35\Bersama\IT\RPA\Perhari" & "\"

n = Sheets.Count
Set FILEBARU = Workbooks.Add
Application.DisplayAlerts = False
    ActiveWorkbook.SaveAs Filename:=saveee & "Daily " & BULAN_STR, FileFormat:=xlOpenXMLWorkbook, CreateBackup:=False
Application.DisplayAlerts = True

'On Error Resume Next
For i = DATATANGGAL To 2 Step -1

    ARAN_TGL = twb.Sheets("TES7").Cells(i, 1)
    ARAN_SHEET = Format(ARAN_TGL, "dd mmm yy")
    
    twb.Activate
    If Evaluate("isref('" & ARAN_SHEET & "'!A1)") Then
          FILECNJ.Worksheets(ARAN_SHEET).Copy Before:=FILEBARU.Sheets(1)
    End If
    
    twb.Activate
    If Evaluate("isref('" & ARAN_SHEET & "'!A1)") Then
        Application.DisplayAlerts = False
            Sheets(ARAN_SHEET).Delete
        Application.DisplayAlerts = True
    End If
Next i

Workbooks("Daily " & BULAN_STR & ".XLSX").Activate
Application.DisplayAlerts = False
    Sheets(Sheets.Count).Delete
Application.DisplayAlerts = True

Workbooks("Daily " & BULAN_STR & ".XLSX").Close savechanges:=True


End Sub

