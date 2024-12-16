Sub TAMBAHAN_BARU()

Application.DisplayAlerts = False

Dim TWB As Workbook, SH_KOTRET As Worksheet, SH_HASIL2 As Worksheet
Dim Rng_Src As Range
Dim Rng_Fill As Range
Dim RNG_ROW As Range

Set TWB = ThisWorkbook
Set SH_HASIL2 = TWB.Sheets("HASIL2")

If Evaluate("isref('" & "KOTRET" & "'!A1)") Then Sheets("KOTRET").Delete
Sheets.Add(After:=Sheets(Sheets.Count)).name = "KOTRET"
Set SH_KOTRET = TWB.Sheets("KOTRET")

SH_KOTRET.Cells.Clear
SH_HASIL2.Activate
Set Rng_Src = Range("A1").CurrentRegion
Rng_Src.Copy
SH_KOTRET.Activate: Range("A1").PasteSpecial xlPasteAll: Application.CutCopyMode = False: Cells.EntireColumn.AutoFit: Cells.EntireRow.AutoFit: Cells(1, 1).Select

'[*]...21 - FEB - 2024
Cells.Borders.LineStyle = xlNone
With Range("a1:i1")
    .Font.name = "Verdana"
    .Font.Color = vbWhite
    .Interior.Pattern = xlSolid
    .Interior.PatternColorIndex = xlAutomatic
    .Interior.Color = RGB(79, 146, 151)
End With

With Range("a2:i2")
    .Font.name = "Verdana"
    .Font.Color = vbWhite
    .Interior.Pattern = xlSolid
    .Interior.PatternColorIndex = xlAutomatic
    .Interior.Color = RGB(52, 98, 101)
    .AutoFilter
End With
Rows(2).Insert
With Range("a2", Cells(2, Range("XFA3").End(xlToLeft).Column))
    .RowHeight = 5
    .Interior.Pattern = xlSolid
    .Interior.PatternColorIndex = xlAutomatic
    .Interior.Color = RGB(228, 240, 241)
End With

If Range("A" & Rows.Count).End(xlUp) <> "GOOD JOB !!" _
    And Range("A" & Rows.Count).End(xlUp) <> "WO" Then
    
    Set Rng_Fill = Range("A4", Cells(Range("A" & Rows.Count).End(xlUp).Row, Range("XFA3").End(xlToLeft).Column))
    For Each RNG_ROW In Rng_Fill.Rows
        If RNG_ROW.Row Mod 2 = 0 Then
            RNG_ROW.Font.name = "Verdana"
            With RNG_ROW.Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .Color = RGB(255, 255, 255)
            End With
        Else
            RNG_ROW.Font.name = "Verdana"
            With RNG_ROW.Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .Color = RGB(228, 240, 241)
            End With
        End If
    Next RNG_ROW
End If
Cells.EntireColumn.AutoFit

Call Yellow_Highlight
Set Rng_Fill = ActiveSheet.UsedRange
Rng_Fill.Copy
SH_HASIL2.Activate: Rng_Src.PasteSpecial xlPasteAll: Application.CutCopyMode = False: Rng_Src.EntireColumn.AutoFit: Cells(1, 1).Select

SH_KOTRET.Cells.Clear
SH_HASIL2.Activate
Set Rng_Src = Range("K1").CurrentRegion
Rng_Src.Copy
SH_KOTRET.Activate: Range("A1").PasteSpecial xlPasteAll: Application.CutCopyMode = False: Cells.EntireColumn.AutoFit: Cells.EntireRow.AutoFit: Cells(1, 1).Select

'[*]...21 - FEB - 2024
Cells.Borders.LineStyle = xlNone
With Range("a1:i1")
    .Font.name = "Verdana"
    .Font.Color = vbWhite
    .Interior.Pattern = xlSolid
    .Interior.PatternColorIndex = xlAutomatic
    .Interior.Color = RGB(79, 146, 151)
End With

With Range("a2:i2")
    .Font.name = "Verdana"
    .Font.Color = vbWhite
    .Interior.Pattern = xlSolid
    .Interior.PatternColorIndex = xlAutomatic
    .Interior.Color = RGB(52, 98, 101)
    .AutoFilter
End With
Rows(2).Insert
With Range("a2", Cells(2, Range("XFA3").End(xlToLeft).Column))
    .RowHeight = 5
    .Interior.Pattern = xlSolid
    .Interior.PatternColorIndex = xlAutomatic
    .Interior.Color = RGB(228, 240, 241)
End With

If Range("A" & Rows.Count).End(xlUp) <> "GOOD JOB !!" _
    And Range("A" & Rows.Count).End(xlUp) <> "WO" Then
    
    Set Rng_Fill = Range("A4", Cells(Range("A" & Rows.Count).End(xlUp).Row, Range("XFA3").End(xlToLeft).Column))
    For Each RNG_ROW In Rng_Fill.Rows
        If RNG_ROW.Row Mod 2 = 0 Then
            RNG_ROW.Font.name = "Verdana"
            With RNG_ROW.Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .Color = RGB(255, 255, 255)
            End With
        Else
            RNG_ROW.Font.name = "Verdana"
            With RNG_ROW.Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .Color = RGB(228, 240, 241)
            End With
        End If
    Next RNG_ROW
End If
Cells.EntireColumn.AutoFit

Call Yellow_Highlight
Set Rng_Fill = ActiveSheet.UsedRange
Rng_Fill.Copy
SH_HASIL2.Activate: Rng_Src.PasteSpecial xlPasteAll: Application.CutCopyMode = False: Rng_Src.EntireColumn.AutoFit: Cells(1, 1).Select

SH_KOTRET.Cells.Clear
SH_HASIL2.Activate
Set Rng_Src = Range("U1").CurrentRegion
Rng_Src.Copy
SH_KOTRET.Activate: Range("A1").PasteSpecial xlPasteAll: Application.CutCopyMode = False: Cells.EntireColumn.AutoFit: Cells.EntireRow.AutoFit: Cells(1, 1).Select
'[*]...21 - FEB - 2024
Cells.Borders.LineStyle = xlNone
With Range("a1:i1")
    .Font.name = "Verdana"
    .Font.Color = vbWhite
    .Interior.Pattern = xlSolid
    .Interior.PatternColorIndex = xlAutomatic
    .Interior.Color = RGB(79, 146, 151)
End With

With Range("a2:i2")
    .Font.name = "Verdana"
    .Font.Color = vbWhite
    .Interior.Pattern = xlSolid
    .Interior.PatternColorIndex = xlAutomatic
    .Interior.Color = RGB(52, 98, 101)
    .AutoFilter
End With
Rows(2).Insert
With Range("a2", Cells(2, Range("XFA3").End(xlToLeft).Column))
    .RowHeight = 5
    .Interior.Pattern = xlSolid
    .Interior.PatternColorIndex = xlAutomatic
    .Interior.Color = RGB(228, 240, 241)
End With

If Range("A" & Rows.Count).End(xlUp) <> "GOOD JOB !!" _
    And Range("A" & Rows.Count).End(xlUp) <> "WO" Then
    
    Set Rng_Fill = Range("A4", Cells(Range("A" & Rows.Count).End(xlUp).Row, Range("XFA3").End(xlToLeft).Column))
    For Each RNG_ROW In Rng_Fill.Rows
        If RNG_ROW.Row Mod 2 = 0 Then
            RNG_ROW.Font.name = "Verdana"
            With RNG_ROW.Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .Color = RGB(255, 255, 255)
            End With
        Else
            RNG_ROW.Font.name = "Verdana"
            With RNG_ROW.Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .Color = RGB(228, 240, 241)
            End With
        End If
    Next RNG_ROW
End If
Cells.EntireColumn.AutoFit

Call Yellow_Highlight
Set Rng_Fill = ActiveSheet.UsedRange
Rng_Fill.Copy
SH_HASIL2.Activate: Rng_Src.PasteSpecial xlPasteAll: Application.CutCopyMode = False: Rng_Src.EntireColumn.AutoFit: Cells(1, 1).Select
SH_KOTRET.Cells.Clear


SH_HASIL2.Activate
Set Rng_Src = Range("AN1").CurrentRegion
Rng_Src.Copy
SH_KOTRET.Activate: Range("A1").PasteSpecial xlPasteAll: Application.CutCopyMode = False: Cells.EntireColumn.AutoFit: Cells.EntireRow.AutoFit: Cells(1, 1).Select

'[*]...21 - FEB - 2024
Cells.Borders.LineStyle = xlNone
With Range("A1:H1")
    .Font.name = "Verdana"
    .Font.Color = vbWhite
    .Interior.Pattern = xlSolid
    .Interior.PatternColorIndex = xlAutomatic
    .Interior.Color = RGB(79, 146, 151)
End With

With Range("a2:h2")
    .Font.name = "Verdana"
    .Font.Color = vbWhite
    .Interior.Pattern = xlSolid
    .Interior.PatternColorIndex = xlAutomatic
    .Interior.Color = RGB(52, 98, 101)
    .AutoFilter
End With
Rows(2).Insert
With Range("a2", Cells(2, Range("XFA3").End(xlToLeft).Column))
    .RowHeight = 5
    .Interior.Pattern = xlSolid
    .Interior.PatternColorIndex = xlAutomatic
    .Interior.Color = RGB(228, 240, 241)
End With

If Range("A" & Rows.Count).End(xlUp) <> "GOOD JOB !!" _
    And Range("A" & Rows.Count).End(xlUp) <> "WO" Then
    
    Set Rng_Fill = Range("A4", Cells(Range("A" & Rows.Count).End(xlUp).Row, Range("XFA3").End(xlToLeft).Column))
    For Each RNG_ROW In Rng_Fill.Rows
        If RNG_ROW.Row Mod 2 = 0 Then
            RNG_ROW.Font.name = "Verdana"
            With RNG_ROW.Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .Color = RGB(255, 255, 255)
            End With
        Else
            RNG_ROW.Font.name = "Verdana"
            With RNG_ROW.Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .Color = RGB(228, 240, 241)
            End With
        End If
    Next RNG_ROW
End If
Cells.EntireColumn.AutoFit

'Call Yellow_Highlight
Set Rng_Fill = ActiveSheet.UsedRange
Rng_Fill.Copy
SH_HASIL2.Activate: Rng_Src.PasteSpecial xlPasteAll: Application.CutCopyMode = False: Rng_Src.EntireColumn.AutoFit: Cells(1, 1).Select
SH_KOTRET.Cells.Clear
SH_HASIL2.Activate
Set Rng_Src = Range("AW1").CurrentRegion
Rng_Src.Copy
SH_KOTRET.Activate: Range("A1").PasteSpecial xlPasteAll: Application.CutCopyMode = False: Cells.EntireColumn.AutoFit: Cells.EntireRow.AutoFit: Cells(1, 1).Select

'[*]...21 - FEB - 2024
Cells.Borders.LineStyle = xlNone
With Range("a1:g1")
    .Font.name = "Verdana"
    .Font.Color = vbWhite
    .Interior.Pattern = xlSolid
    .Interior.PatternColorIndex = xlAutomatic
    .Interior.Color = RGB(79, 146, 151)
End With

With Range("a2:g2")
    .Font.name = "Verdana"
    .Font.Color = vbWhite
    .Interior.Pattern = xlSolid
    .Interior.PatternColorIndex = xlAutomatic
    .Interior.Color = RGB(52, 98, 101)
    .AutoFilter
End With
Rows(2).Insert
With Range("a2", Cells(2, Range("XFA3").End(xlToLeft).Column))
    .RowHeight = 5
    .Interior.Pattern = xlSolid
    .Interior.PatternColorIndex = xlAutomatic
    .Interior.Color = RGB(228, 240, 241)
End With

If Range("A" & Rows.Count).End(xlUp) <> "GOOD JOB !!" _
    And Range("A" & Rows.Count).End(xlUp) <> "WO" Then
    
    Set Rng_Fill = Range("A4", Cells(Range("A" & Rows.Count).End(xlUp).Row, Range("XFA3").End(xlToLeft).Column))
    For Each RNG_ROW In Rng_Fill.Rows
        If RNG_ROW.Row Mod 2 = 0 Then
            RNG_ROW.Font.name = "Verdana"
            With RNG_ROW.Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .Color = RGB(255, 255, 255)
            End With
        Else
            RNG_ROW.Font.name = "Verdana"
            With RNG_ROW.Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .Color = RGB(228, 240, 241)
            End With
        End If
    Next RNG_ROW
End If
Cells.EntireColumn.AutoFit

Call Yellow_Highlight
Set Rng_Fill = ActiveSheet.UsedRange
Rng_Fill.Copy
SH_HASIL2.Activate: Rng_Src.PasteSpecial xlPasteAll: Application.CutCopyMode = False: Rng_Src.EntireColumn.AutoFit: Cells(1, 1).Select
SH_KOTRET.Cells.Clear
SH_HASIL2.Activate
Set Rng_Src = Range("BE1").CurrentRegion
Rng_Src.Copy
SH_KOTRET.Activate: Range("A1").PasteSpecial xlPasteAll: Application.CutCopyMode = False: Cells.EntireColumn.AutoFit: Cells.EntireRow.AutoFit: Cells(1, 1).Select

'[*]...21 - FEB - 2024
Cells.Borders.LineStyle = xlNone
With Range("a1:g1")
    .Font.name = "Verdana"
    .Font.Color = vbWhite
    .Interior.Pattern = xlSolid
    .Interior.PatternColorIndex = xlAutomatic
    .Interior.Color = RGB(79, 146, 151)
End With

With Range("a2:g2")
    .Font.name = "Verdana"
    .Font.Color = vbWhite
    .Interior.Pattern = xlSolid
    .Interior.PatternColorIndex = xlAutomatic
    .Interior.Color = RGB(52, 98, 101)
    .AutoFilter
End With
Rows(2).Insert
With Range("a2", Cells(2, Range("XFA3").End(xlToLeft).Column))
    .RowHeight = 5
    .Interior.Pattern = xlSolid
    .Interior.PatternColorIndex = xlAutomatic
    .Interior.Color = RGB(228, 240, 241)
End With

If Range("A" & Rows.Count).End(xlUp) <> "GOOD JOB !!" _
    And Range("A" & Rows.Count).End(xlUp) <> "WO" Then
    
    Set Rng_Fill = Range("A4", Cells(Range("A" & Rows.Count).End(xlUp).Row, Range("XFA3").End(xlToLeft).Column))
    For Each RNG_ROW In Rng_Fill.Rows
        If RNG_ROW.Row Mod 2 = 0 Then
            RNG_ROW.Font.name = "Verdana"
            With RNG_ROW.Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .Color = RGB(255, 255, 255)
            End With
        Else
            RNG_ROW.Font.name = "Verdana"
            With RNG_ROW.Interior
                .Pattern = xlSolid
                .PatternColorIndex = xlAutomatic
                .Color = RGB(228, 240, 241)
            End With
        End If
    Next RNG_ROW
End If
Cells.EntireColumn.AutoFit

'Call Yellow_Highlight
Set Rng_Fill = ActiveSheet.UsedRange
Rng_Fill.Copy
SH_HASIL2.Activate: Rng_Src.PasteSpecial xlPasteAll: Application.CutCopyMode = False: Rng_Src.EntireColumn.AutoFit: Cells(1, 1).Select
If Evaluate("isref('" & "KOTRET" & "'!A1)") Then Sheets("KOTRET").Delete

SH_HASIL2.Activate
ActiveWindow.Zoom = 85
Rows(1).RowHeight = 20
Rows(2).RowHeight = 5
Rows(3).RowHeight = 25

Rows("1:2").Insert
Cells(1, 1).Select

End Sub
