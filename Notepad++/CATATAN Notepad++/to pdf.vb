Sub Export_PDF_and_End_Program()
    TMP3.Activate
    lr = Cells.Find("*", , xlFormulas, xlPart, xlByRows, xlPrevious).Row
    lc = Cells.Find("*", , xlFormulas, xlPart, xlByColumns, xlPrevious).Column
    Set Rng = Range(Cells(1, 1), Cells(lr, lc))
    With TMP3.PageSetup
        .PrintArea = Rng.Address
        .Orientation = xlLandscape
        .CenterHorizontally = True
        .CenterVertically = True
        .Zoom = False
        .FitToPagesTall = 3
        .FitToPagesWide = 1
    End With
    TMP3.ExportAsFixedFormat Type:=xlTypePDF, Filename:=path_resume
    
    TMP1.Name = "SOURCE"
    TMP2.Name = "PIVOT"
    TMP3.Name = "RESUME"
    SH_PARENT.Activate: Cells(1, 1).Select
    HOME.Activate: Cells(1, 1).Select
    
End Sub

Sub Export_File_To_Pdf()

    Dim TWB As Workbook
    Dim SH_HOME As Worksheet
    Dim SH_RESUME As Worksheet
    Dim SH_INFA_ININ As Worksheet
    Dim SH_INPA_INUM_INAC As Worksheet
    
    Dim PATH_PDF_INFA_ININ As String
    Dim PATH_INPA_INUM_INAC As String
    
    Dim LR As Long, LC As Long, RNG As Range
    
    Set TWB = ThisWorkbook
    Set SH_HOME = TWB.Sheets("HOME")
    Set SH_RESUME = TWB.Sheets("RESUME")
    Set SH_INFA_ININ = TWB.Sheets("RESUME INFA_ININ")
    Set SH_INPA_INUM_INAC = TWB.Sheets("RESUME INPA_INUM_INAC")
    
    PATH_PDF_INFA_ININ = SH_HOME.Range("F32") & Application.PathSeparator & "INFA_ININ.pdf"
    PATH_INPA_INUM_INAC = SH_HOME.Range("F33") & Application.PathSeparator & "INPA_INUM_INAC.pdf"
    
    SH_INFA_ININ.Activate
    LR = Range("A1000").End(xlUp).Row
    If LR = 2 Then
        Cells(2, 1) = "WO FINANCE EXCESES FABRIC"
        With Range(Cells(2, 1), Cells(3, 11))
            .Merge
            .Font.Bold = True
            .Font.Size = 24
            .Font.Name = "Calibri"
            .Font.Underline = True
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With
        Range("A5") = "NO DATA EXCESS"
        Cells.EntireColumn.AutoFit
        Cells(1, 1).Select
        Set RNG = Range(Cells(1, 1), Cells(20, 11))
    Else
        LR = Cells.Find("*", , xlFormulas, xlPart, xlByRows, xlPrevious).Row
        LC = Cells.Find("*", , xlFormulas, xlPart, xlByColumns, xlPrevious).Column
        Set RNG = Range(Cells(1, 1), Cells(LR, LC))
    End If
    
    With SH_INFA_ININ.PageSetup
        .TopMargin = Application.InchesToPoints(0.5)
        .PrintArea = RNG.Address
        .Orientation = xlPortrait
        .CenterHorizontally = True
        .CenterVertically = False
        .Zoom = False
        .FitToPagesTall = 1
        .FitToPagesWide = 1
    End With
    SH_INFA_ININ.ExportAsFixedFormat Type:=xlTypePDF, Filename:=PATH_PDF_INFA_ININ
    
    
    SH_INPA_INUM_INAC.Activate
    LR = Range("A1000").End(xlUp).Row
    If LR = 2 Then
        Cells(2, 1) = "WO FINANCE EXCESES ACCESSORIES"
        With Range(Cells(2, 1), Cells(3, 11))
            .Merge
            .Font.Bold = True
            .Font.Size = 24
            .Font.Name = "Calibri"
            .Font.Underline = True
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlCenter
        End With
        Range("A5") = "NO DATA EXCESS"
        Cells.EntireColumn.AutoFit
        Cells(1, 1).Select
        Set RNG = Range(Cells(1, 1), Cells(20, 11))
    Else
        LR = Cells.Find("*", , xlFormulas, xlPart, xlByRows, xlPrevious).Row
        LC = Cells.Find("*", , xlFormulas, xlPart, xlByColumns, xlPrevious).Column
        Set RNG = Range(Cells(1, 1), Cells(LR, LC))
    End If

    With SH_INPA_INUM_INAC.PageSetup
        .PrintArea = RNG.Address
        .Orientation = xlPortrait
        .TopMargin = Application.InchesToPoints(0.5)
        .CenterHorizontally = True
        .CenterVertically = False
        .Zoom = False
        .FitToPagesTall = 1
        .FitToPagesWide = 1
    End With
    SH_INPA_INUM_INAC.ExportAsFixedFormat Type:=xlTypePDF, Filename:=PATH_INPA_INUM_INAC
    
End Sub