Sub SaveSheetAsPDF2()
    Dim ws As Worksheet, rgSAVE As Range, lr As Long, lc As Long
    Set ws = ThisWorkbook.Sheets("HASIL2") ' Ganti "Sheet1" dengan nama lembar kerja yang diinginkan
    
    'ws.Activate
    'ws.Rows(1).RowHeight = 20
'    ws.Rows(1).VerticalAlignment = xlBottom
'    Range("A:A").ColumnWidth = 10
'    Range("D:D").ColumnWidth = 10
    'Stop
    
    Range("AJ:AL").Delete
    'Range("A:A,D:D,K:K,N:N,U:U,X:X,AI:AI,AL:AL,AR:AR,AU:AU,AZ:AZ,BC:BC").ColumnWidth = 8
    'Range("A:A,D:D,K:K,N:N,U:U,X:X,AK:AK,AN:AN,AT:AT,AW:AW,BC:BC,BE:BE").ColumnWidth = 8
    'Range("BB:BB").ColumnWidth = 8
    
    Dim LuasLebar As Long
    'Range("BC:BC").EntireColumn.AutoFit
    'LuasLebar = Range("BC:BC").ColumnWidth
    'Range("BC:BC").ColumnWidth = LuasLebar + 2
    
    'Range("A:A,D:D,K:K,N:N,U:U,X:X,AN:AN,AQ:AQ,AW:AW,AZ:AZ,BE:BE,BH:BH").ColumnWidth = 8
    'Stop
    Dim RNG_COL As Range, LEBAR_AWAL As Double
    lr = ws.Cells.Find(what:="*" _
            , lookat:=xlPart _
            , LookIn:=xlFormulas _
            , Searchorder:=xlByRows _
            , searchdirection:=xlPrevious).Row
    
    lc = ws.Cells.Find(what:="*" _
            , lookat:=xlPart _
            , LookIn:=xlFormulas _
            , Searchorder:=xlByColumns _
            , searchdirection:=xlPrevious).Column
    
    Set rgSAVE = Range(Cells(1, 1), Cells(lr, lc))
    
    For Each RNG_COL In rgSAVE.Columns
        If RNG_COL.Count <> 0 Then
            RNG_COL.EntireColumn.AutoFit
            LEBAR_AWAL = RNG_COL.ColumnWidth
            RNG_COL.ColumnWidth = LEBAR_AWAL + 1
        End If
    Next RNG_COL

    With Range("a1")
        .Font.Bold = True
        .Font.name = "Century Gothic"
        .Font.Size = 30
        .Font.Color = RGB(79, 146, 151)
        .Value = "WO PRODUCTION CUTTING"
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlBottom
        .Select
    End With
    Rows(1).Insert
    Rows("3").Insert
    Rows("3").RowHeight = 15
    
    lr = ws.Cells.Find(what:="*" _
            , lookat:=xlPart _
            , LookIn:=xlFormulas _
            , Searchorder:=xlByRows _
            , searchdirection:=xlPrevious).Row
    
    lc = ws.Cells.Find(what:="*" _
            , lookat:=xlPart _
            , LookIn:=xlFormulas _
            , Searchorder:=xlByColumns _
            , searchdirection:=xlPrevious).Column
            
    With Range(Cells(2, 1), Cells(2, "AC"))
        .Merge
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
    End With
    
    Set rgSAVE = Range(Cells(1, 1), Cells(lr, lc))
            
    ws.PageSetup.PrintArea = rgSAVE.Address
    
    Application.PrintCommunication = False
    With ws.PageSetup
        .CenterHorizontally = True
        .LeftMargin = Application.InchesToPoints(0.7)
        .RightMargin = Application.InchesToPoints(0.7)
        .TopMargin = Application.InchesToPoints(0.1)
        .BottomMargin = Application.InchesToPoints(0.1)
        .HeaderMargin = Application.InchesToPoints(0.1)
        .FooterMargin = Application.InchesToPoints(0.1)
        .Orientation = xlLandscape ' Ganti dengan xlPortrait jika ingin potret
        .FitToPagesTall = 1
        .FitToPagesWide = 2
    End With
    
    Application.PrintCommunication = True
    
    ActiveWindow.View = xlPageBreakPreview
    ActiveSheet.ResetAllPageBreaks
    Set ActiveSheet.VPageBreaks(1).Location = Range("AD1")
    'ActiveSheet.VPageBreaks(2).DragOff Direction:=xlToRight, RegionIndex:=1
    Columns("AE:AJ").Delete Shift:=xlToLeft
    ActiveWindow.View = xlNormalView

'    ActiveWindow.View = xlPageBreakPreview
'    ActiveSheet.ResetAllPageBreaks
'    Set ActiveSheet.VPageBreaks(1).Location = Range("AK1")
'    ActiveWindow.View = xlNormalView
    
    ' Menyimpan satu lembar kerja sebagai satu file PDF
    ws.ExportAsFixedFormat Type:=xlTypePDF, Filename:=ThisWorkbook.path & Application.PathSeparator & "WO Produksi - Cutting.pdf_2"
    ' Mengembalikan orientasi ke potret jika diperlukan
    With ws.PageSetup
        .Orientation = xlPortrait
    End With
End Sub