    LR = Cells.Find("*", , xlFormulas, xlPart, xlByRows, xlPrevious).Row
    LC = Cells.Find("*", , xlFormulas, xlPart, xlByColumns, xlPrevious).Column
    
    Set rng = Range(Cells(1, 1), Cells(LR, LC))
    With rng
        .Copy: .PasteSpecial xlPasteValuesAndNumberFormats: Application.CutCopyMode = False: Cells.EntireColumn.AutoFit: Cells(1, 1).Select
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Borders.LineStyle = xlContinuous
    End With
    
    Range(Cells(2, 1), Cells(LR, 2)).HorizontalAlignment = xlLeft
    
    Cells.Font.Name = "Trebuchet MS"
    
    With Range(Cells(1, 1), Cells(1, LC))
        .Font.Bold = True
        .Font.Name = "Century Gothic"
        .Font.Size = 12
        .Font.Color = vbWhite
        .Interior.Pattern = xlSolid
        .Interior.PatternColor = xlAutomatic
        .Interior.Color = RGB(31, 76, 81)
    End With
    
    Cells.EntireColumn.AutoFit
    Cells.EntireRow.AutoFit
    
    Rows(1).RowHeight = 39
    Range("D1:G1").WrapText = True
    Range("D:D").ColumnWidth = 8
    Range("E:E").ColumnWidth = 11
    Range("F:F").ColumnWidth = 10
    Range("G:G").ColumnWidth = 11
    Cells.EntireColumn.AutoFit
    
    For i = 2 To LR
        If i Mod 2 = 0 And i <> LR Then
            With Range(Cells(i, 1), Cells(i, LC))
                .Interior.Pattern = xlSolid
                .Interior.PatternColor = xlAutomatic
                .Interior.Color = RGB(191, 227, 231)
            End With
        ElseIf i Mod 2 <> 0 And i <> LR Then
            With Range(Cells(i, 1), Cells(i, LC))
                .Interior.Pattern = xlSolid
                .Interior.PatternColor = xlAutomatic
                .Interior.Color = RGB(255, 255, 255)
            End With
        ElseIf i = LR Then
            With Range(Cells(i, 1), Cells(i, LC))
                .Interior.Pattern = xlSolid
                .Interior.PatternColor = xlAutomatic
                .Interior.Color = RGB(56, 145, 154)
                .RowHeight = .RowHeight + 1
            End With
        End If
    Next i
    
    LR = Cells.Find("*", , xlFormulas, xlPart, xlByRows, xlPrevious).Row
    LC = Cells.Find("*", , xlFormulas, xlPart, xlByColumns, xlPrevious).Column
    
    Set rng_Data = Range(Cells(1, 1), Cells(LR, LC))

    For Each rng In rng_Data.Columns
        rng.ColumnWidth = rng.ColumnWidth + 2
    Next rng
    
    Rows("1:3").Insert
    With Range(Cells(2, 1), Cells(2, LC))
        .Merge
        .Value = "RESUME AUDIT MESIN (WEEKLY)"
        .Font.Bold = True
        .Font.Name = "Trebuchet MS"
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .RowHeight = 25
        .Interior.Pattern = xlSolid
        .Interior.PatternColor = xlAutomatic
        .Interior.Color = RGB(151, 211, 217)
    End With
        
    With Range(Cells(3, 1), Cells(3, LC))
        .Interior.Pattern = xlSolid
        .Interior.PatternColor = xlAutomatic
        .Interior.Color = RGB(233, 246, 247)
    End With
    
    Range("A:A").Insert
    Range("A:A").ColumnWidth = 3
    Rows(5).Select
    ActiveWindow.FreezePanes = True
    
    LR = Cells.Find("*", , xlFormulas, xlPart, xlByRows, xlPrevious).Row
    LC = Cells.Find("*", , xlFormulas, xlPart, xlByColumns, xlPrevious).Column

    Set rng_Data = Range(Cells(2, 2), Cells(LR, LC))
    
    Range("A1").Select
    TEMP3.Name = "RESUME"
    
    With TEMP3.PageSetup
        .PrintArea = rng_Data.Address
        .Orientation = xlPortrait
        .CenterHorizontally = True
        .Zoom = False
        .FitToPagesTall = 1
        .FitToPagesWide = 1
    End With
    
    TEMP3.ExportAsFixedFormat Type:=xlTypePDF, Filename:=str_PathPercent
    
    On Error Resume Next
        Randomize
        randomColor = RGB(Int(Rnd() * 256), Int(Rnd() * 256), Int(Rnd() * 256))
        With TEMP3.Tab
        .Color = randomColor
        .TintAndShade = 0
        End With
        
        TEMP3.PageSetup.PrintArea = ""
        TEMP3.DisplayPageBreaks = False
    On Error GoTo 0
