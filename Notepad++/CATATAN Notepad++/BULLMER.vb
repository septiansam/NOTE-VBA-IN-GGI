
Sub Main()
Application.DisplayAlerts = False

Dim TWB As Workbook, WB_TARIKAN As Workbook, SH_TARIKAN As Worksheet, SH_RESULTS As Worksheet
Dim LR As Long, LC As Long, RNG As Range, Baris As Range, RNG_PDF As Range, cell As Range
Dim i As Long, j As Long
Dim path_Tarikan As String, nama_Tarikan As String, ext As String
Dim path_Hasil As String, JUDUL As String

Set TWB = ThisWorkbook

For i = TWB.Sheets.Count To 2 Step -1
    Sheets(i).Delete
Next i

nama_Tarikan = HOME.Range("C8").Value
ext = HOME.Range("D8").Value
path_Tarikan = HOME.Range("B8").Value & Application.PathSeparator & nama_Tarikan & ext
path_Hasil = HOME.Range("F8").Value & _
                Application.PathSeparator & _
                HOME.Range("G8").Value & _
                HOME.Range("H8").Value
JUDUL = "Report List Bullmer Periode " _
        & Application.WorksheetFunction.Text(Date - 8, "[$-id-ID]dd Mmmm yyyy") & " - " _
        & Application.WorksheetFunction.Text(Date - 1, "[$-id-ID]dd Mmmm yyyy")
'Debug.Print JUDUL
'Stop

If Dir(path_Tarikan) = "" Then
    MsgBox "File " & nama_Tarikan & " Does't Exists", vbInformation, "File " & nama_Tarikan & " Not Found!"
    Exit Sub
End If

Set SH_TARIKAN = Sheets.Add(After:=Sheets(Sheets.Count)): ActiveSheet.Name = "TARIKAN"

Set WB_TARIKAN = Workbooks.Open(path_Tarikan)
With WB_TARIKAN
    .Sheets(1).Activate
    .Sheets(1).UsedRange.Copy
End With
SH_TARIKAN.Range("A1").PasteSpecial xlPasteAll: Application.CutCopyMode = False
WB_TARIKAN.Close False

SH_TARIKAN.Activate
ActiveWindow.Zoom = 85
Rows(1).RowHeight = 30
Rows(1).UnMerge
Rows(1).Clear
Range("B:B").Delete

Range("A1").Value = JUDUL
LR = Range("A" & Rows.Count).End(xlUp).Row
LC = Range("XFA2").End(xlToLeft).Column

Range(Cells(1, 1), Cells(1, LC)).Merge

Cells.Font.Name = "Verdana"
Cells.EntireColumn.AutoFit
Cells.HorizontalAlignment = xlCenter
Cells.VerticalAlignment = xlCenter

Rows(1).Font.Name = "Century Gothic"
Rows(1).Font.Size = 16
Rows(1).Font.Bold = True
Set RNG = Range(Cells(2, 1), Cells(LR, LC))
With RNG
    .Rows(1).Font.Bold = True
    .Rows(1).Font.Size = 14
    .Rows(1).Font.Name = "Century Gothic"
    .Rows(1).Font.Color = vbWhite
    .Rows(1).Interior.Pattern = xlSolid
    .Rows(1).Interior.PatternColorIndex = xlAutomatic
    .Rows(1).Interior.Color = RGB(52, 98, 101)
    .Rows(1).RowHeight = .Rows(1).RowHeight + 8
    For Each Baris In .Rows
    
        If Baris.Row > 2 And Baris.Row Mod 2 = 0 Then
            Baris.Font.Name = "Verdana"
            With Baris
                .Font.Size = 12
                .Interior.Pattern = xlSolid
                .Interior.PatternColorIndex = xlAutomatic
                .Interior.Color = RGB(255, 255, 255)
                .RowHeight = .RowHeight + 5
            End With
        ElseIf Baris.Row > 2 And Baris.Row Mod 2 <> 0 Then
            Baris.Font.Name = "Verdana"
            With Baris
                .Font.Size = 12
                .Interior.Pattern = xlSolid
                .Interior.PatternColorIndex = xlAutomatic
                .Interior.Color = RGB(228, 240, 241)
                .RowHeight = .RowHeight + 5
            End With
        End If
    
    Next Baris
End With
With Range(Cells(LR, 1), Cells(LR, LC))
    .Font.Bold = True
    .Font.Color = vbWhite
    .Interior.Pattern = xlSolid
    .Interior.PatternColorIndex = xlAutomatic
    .Interior.Color = RGB(0, 72, 65)
    .RowHeight = .RowHeight + 3
End With

LR = Range("A" & Rows.Count).End(xlUp).Row
LC = Range("XFA2").End(xlToLeft).Column

Stop
Dim colEfficiency As Long
colEfficiency = Application.WorksheetFunction.Match("Efficiency", Rows(2), 0)

Columns(colEfficiency).NumberFormat = "_(#,##0.0_);[Red]_((#,##0.0);_("" - ""?_);_(@_)"

'Set RNG = Range("F3:F" & Range("F" & Rows.Count).End(xlUp).Row - 1)
Set RNG = Range(Cells(3, colEfficiency), Cells(Cells(Rows.Count, colEfficiency).End(xlUp).Row - 1, colEfficiency))
RNG.Select

For Each cell In RNG
    If cell.Value < 60 Then
        With cell
            .Interior.Pattern = xlSolid
            .Interior.PatternColorIndex = xlAutomatic
            .Interior.Color = vbYellow
        End With
    End If
Next cell
Range("F" & LR).NumberFormat = "_(#,##0.00_);[Red]_((#,##0.00);_("" - ""?_);_(@_)"

Cells.EntireColumn.AutoFit
Set RNG = Range(Cells(1, 1), Cells(LR, LC))
For Each cell In RNG.Columns
    cell.EntireColumn.AutoFit
    cell.EntireColumn.ColumnWidth = cell.EntireColumn.ColumnWidth + 2
Next cell

Rows("1:1").Insert
Range("A:A").Insert
Range("A:A").ColumnWidth = 3
Rows("4:4").Insert
Rows("4:4").RowHeight = 5

LC = Range("XFA3").End(xlToLeft).Column
With Range(Cells(4, 2), Cells(4, LC))
    .Interior.Pattern = xlSolid
    .Interior.PatternColorIndex = xlAutomatic
    .Interior.Color = RGB(0, 36, 52)
End With
Rows("5:5").Select
ActiveWindow.FreezePanes = True
Rows("3:3").Insert
Rows("3:3").RowHeight = 15
Range("A2").EntireRow.Borders.LineStyle = xlNone

LR = Range("B" & Rows.Count).End(xlUp).Row
Set RNG_PDF = Range(Cells(2, 2), Cells(LR, LC))

Application.PrintCommunication = False
With SH_TARIKAN.PageSetup
    .PrintArea = RNG_PDF.Address
    .CenterHorizontally = True
    .LeftMargin = Application.InchesToPoints(0.7)
    .RightMargin = Application.InchesToPoints(0.7)
    .TopMargin = Application.InchesToPoints(0.5)
    .BottomMargin = Application.InchesToPoints(0.1)
    .HeaderMargin = Application.InchesToPoints(0.1)
    .FooterMargin = Application.InchesToPoints(0.1)
    .Orientation = xlPortrait ' Ganti dengan xlLandscape jika ingin Landscape
    .FitToPagesTall = 1
    .FitToPagesWide = 1
End With
Application.PrintCommunication = True
Cells(1, 1).Select

SH_TARIKAN.ExportAsFixedFormat Type:=xlTypePDF, Filename:=path_Hasil
SH_TARIKAN.Name = "RESULTS"
With SH_TARIKAN.Tab
    .Color = 15773696
    .TintAndShade = 0
End With

HOME.Activate
Cells(1, 1).Select

Application.DisplayAlerts = True
End Sub

Public Function WorksheetExists(shtName As String) As Boolean
    On Error Resume Next
    WorksheetExists = Not Sheets(shtName) Is Nothing
    On Error GoTo 0
End Function
