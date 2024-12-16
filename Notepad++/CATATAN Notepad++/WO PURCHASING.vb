Sheets("TES3").Select

save = alamat & ARAN
Cells.EntireColumn.AutoFit
'Cells.EntireRow.AutoFit
Columns("F").EntireColumn.AutoFit
Columns("E").NumberFormat = "MM/D/YYYY"
Columns("AF").NumberFormat = "MM/D/YYYY"
Columns("AC").NumberFormat = "MM/D/YYYY"
Columns("M:N").NumberFormat = "MM/D/YYYY"

Application.DisplayAlerts = False
ARAN = "WO PURCHASING_" & NAMA_ORANG & ".pdf"
namafile = ARAN
With ActiveSheet.PageSetup
    .Orientation = xlLandscape
    .PrintArea = "A1:AB" & LASTROW
    .FitToPagesTall = False
    .FitToPagesWide = 1
    .Zoom = False
    .PaperSize = xlPaperA4
End With
Application.DisplayAlerts = True

Rows(2).AutoFilter
    ActiveWorkbook.Worksheets("TES3").AutoFilter.Sort.SortFields.Add2 Key:=Range("AA2"), SortOn:=xlSortOnValues, Order:=xlAscending, DataOption:=xlSortTextAsNumbers
    With ActiveWorkbook.Worksheets("TES3").AutoFilter.Sort
        .Header = xlYes
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    
For i = 3 To LASTROW
    Cells(i, 1) = i - 2
Next i

Cells.AutoFilter

With RNG
    .Rows(1).Font.Bold = True
    .Rows(1).Font.Name = "Verdana"
    .Rows(1).Font.Color = vbWhite
    .Rows(1).Interior.Pattern = xlSolid
    .Rows(1).Interior.PatternColorIndex = xlAutomatic
    .Rows(1).Interior.Color = RGB(52, 98, 101)
    .Rows(1).RowHeight = .RowHeight + 2
    .Rows(1).VerticalAlignment = xlCenter
    .Rows(1).RowHeight = .Rows(1).RowHeight + 3
    For Each Baris In .Rows
    
        If Baris.Row > 2 And Baris.Row Mod 2 = 0 Then
            Baris.Font.Name = "Verdana"
            With Baris
                .Interior.Pattern = xlSolid
                .Interior.PatternColorIndex = xlAutomatic
                .Interior.Color = RGB(255, 255, 255)
                .VerticalAlignment = xlCenter
                .RowHeight = .RowHeight + 3
            End With
        ElseIf Baris.Row > 2 And Baris.Row Mod 2 <> 0 Then
            Baris.Font.Name = "Verdana"
            With Baris
                .Interior.Pattern = xlSolid
                .Interior.PatternColorIndex = xlAutomatic
                .Interior.Color = RGB(228, 240, 241)
                .VerticalAlignment = xlCenter
                .RowHeight = .RowHeight + 3
            End With
        End If
    
    Next Baris
End With

With RNG
    .Borders.LineStyle = xlContinuous: .Borders.Weight = xlThin
End With

Cells.EntireColumn.AutoFit