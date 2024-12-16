Sub menu05()
    Dim n_fty As Integer
    Dim nn As Integer
    Dim ntgl As Integer
    Dim nweek As Integer
    Dim nnn As Integer
    Dim nnweek As Integer
    Dim eachweek As Integer
    Dim sumamount As Long
    Dim sheetname As String
    Dim weekname As String
    
    Sheets("All_Fty").Select
    Range("A1048576").End(xlUp).Offset(1, 0).Select
    Range(Selection, "A1").Select
    Selection.EntireRow.Delete
    Range("A1").Value = "SUMMARY OVERTIME ALL FACTORY"
    Range("A2").Select
    Selection.Formula = "=RecapCLN!R2C1"
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues
    
    Range("A4").Value = "BULAN"
    Range("B4").Value = "TANGGAL"
    Sheets("Tombol").Select
    Range("O1").End(xlToLeft).Select
    Range(Selection, "D1").Select
    Selection.Copy
    Sheets("All_Fty").Select
    Range("C4").Select
    Selection.PasteSpecial Paste:=xlPasteValues
    n_fty = Selection.Columns.Count
    Range("O4").End(xlToLeft).Offset(0, 1).Select
    Selection.Value = "TOTAL"

    Sheets("RecapCLN").Select
    Range("A1048576").End(xlUp).Offset(0, 1).Select
    Range(Selection, "A7").Select
    ntgl = Selection.Rows.Count
    Selection.Copy
    Sheets("All_Fty").Select
    Range("A5").Select
    Selection.PasteSpecial Paste:=xlPasteValues
    
    nn = 1
    For nn = 1 To n_fty
        Sheets("Tombol").Select
        Range("D2").Offset(0, nn - 1).Select
        sheetname = Selection.Value
        Sheets("Recap" & sheetname).Select
        Range("A1048576").End(xlUp).Offset(0, 2).Select
        Range(Selection, "C7").Select
        Selection.Copy
        Sheets("All_Fty").Select
        Range("C5").Offset(0, nn - 1).Select
        Selection.PasteSpecial Paste:=xlPasteValues
    Next nn

    Range("O5").End(xlToLeft).Offset(0, 1).Select
    Selection.Formula = "=SUM(RC[-" & n_fty & "]:RC[-1])"
    Selection.Copy
    Range("A1048576").End(xlUp).Offset(0, n_fty + 2).Select
    Range(Selection, Selection.End(xlUp)).Select
    ActiveSheet.Paste
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues
    
    Range("A1048576").End(xlUp).Offset(0, n_fty + 2).Select
    Range(Selection, "C5").Select
    Selection.NumberFormat = "#,##0.00_);[Red](#,##0.00)"

    Range("A1048576").End(xlUp).Offset(1, 0).Select
    Selection.Value = "TOTAL OVERTIME"
    Selection.Offset(0, 2).Select
    Selection.Formula = "=SUM(R[-" & ntgl & "]C:R[-1]C)"
    Selection.Copy
    Selection.Resize(Selection.Rows.Count, Selection.Columns.Count + n_fty).Select
    ActiveSheet.Paste
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues
    
    Range("A1048576").End(xlUp).Offset(0, n_fty + 2).Select
    sumamount = Selection.Value
    
    Range("A1048576").End(xlUp).Offset(1, 2).Select
    Selection.Formula = "=if(" & sumamount & "=0,0,R[-1]C/" & sumamount & ")"
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .ReadingOrder = xlContext
    End With
    Selection.NumberFormat = "0.0%"
    Selection.Copy
    Selection.Resize(Selection.Rows.Count, Selection.Columns.Count + n_fty - 1).Select
    ActiveSheet.Paste
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues
    
    Range("A1048576").End(xlUp).Offset(0, n_fty + 2).Select
    Range(Selection, "A4").Select
    With Selection.Borders(xlEdgeLeft)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeTop)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeBottom)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    With Selection.Borders(xlEdgeRight)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideVertical)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    With Selection.Borders(xlInsideHorizontal)
        .LineStyle = xlContinuous
        .Weight = xlThin
    End With
    
    Range("A1048576").End(xlUp).Select
    Selection.Resize(Selection.Rows.Count, Selection.Columns.Count + 2 + n_fty).Select
    With Selection.Font
        .Color = -16776961
        .TintAndShade = 0
    End With
    Selection.Font.Bold = True
    Range("A1048576").End(xlUp).Select
    Selection.Resize(Selection.Rows.Count, Selection.Columns.Count + 1).Select
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .ReadingOrder = xlContext
        .MergeCells = True
    End With
    
    Range("O4").End(xlToLeft).Select
    Range(Selection, "A4").Select
    Selection.Font.Bold = True
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .ReadingOrder = xlContext
    End With
    
    Range("O4").End(xlToLeft).Offset(-3, 0).Select
    Range(Selection, "A1").Select
    Selection.Font.Bold = True
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .ReadingOrder = xlContext
        .MergeCells = True
    End With
    
    Range("O4").End(xlToLeft).Offset(-2, 0).Select
    Range(Selection, "A2").Select
    Selection.Font.Bold = True
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .ReadingOrder = xlContext
        .MergeCells = True
    End With
    
    Range("O5").End(xlToLeft).Offset(0, 1).Select
    Selection.Formula = "=CLN!R2C19"
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues
    Selection.Offset(1, 0).Select
    Selection.Formula = "=R[-1]C+1"
    Selection.Copy
    Selection.Resize(Selection.Rows.Count + ntgl - 1, Selection.Columns.Count).Select
    ActiveSheet.Paste
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues
    
    Range("O5").End(xlToLeft).Offset(0, 1).Select
    Selection.Formula = "=WEEKDAY(RC[-1])"
    Selection.Copy
    Selection.Resize(Selection.Rows.Count + ntgl, Selection.Columns.Count).Select
    ActiveSheet.Paste
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues
    Range("O5").End(xlToLeft).Select
    Selection.End(xlDown).Select
    Selection.Value = "2"
    Selection.End(xlUp).Offset(-1, 0).Select
    Selection.Formula = "=IF(R5C13=2,COUNTIF(R5C" & n_fty + 5 & ":R" & ntgl + 5 & "C" & n_fty + 5 & ",2)-1,COUNTIF(R5C" & n_fty + 5 & ":R" & ntgl + 5 & "C" & n_fty + 5 & ",2))"
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues
    nnweek = Selection.Value
    
    Range("O5").End(xlToLeft).Offset(1, 1).Select
    Selection.Formula = "=IF(RC[-1]=2,IF(R5C13=2,COUNTIF(R5C" & n_fty + 5 & ":RC[-1],2)-1,COUNTIF(R5C" & n_fty + 5 & ":RC[-1],2)),0)"
    Selection.Copy
    Selection.Resize(Selection.Rows.Count + ntgl - 1, Selection.Columns.Count).Select
    ActiveSheet.Paste
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues

    Range("Z5").End(xlToLeft).Offset(1, 2).Select
    Selection.Formula = "=IF(RC[-1]>0,CONCATENATE(""WEEK"",RC[-1]),0)"
    Selection.Copy
    Selection.Resize(Selection.Rows.Count + ntgl - 1, Selection.Columns.Count).Select
    ActiveSheet.Paste
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues
    
    Range("Z5").End(xlToLeft).Offset(0, 3).Select
    Selection.Formula = "=IF(RC[-3]=1,WEEKNUM(RC[-4])-1,WEEKNUM(RC[-4]))"
    Selection.Copy
    Selection.Resize(Selection.Rows.Count + ntgl, Selection.Columns.Count).Select
    ActiveSheet.Paste
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues
    Range("Z5").End(xlToLeft).Select
    Selection.End(xlDown).Select
    Selection.Value = "NEXT"
    
    Range("Z5").End(xlToLeft).Offset(1, 1).Select
    Selection.Formula = _
        "=IF(R[-1]C[-1]<>RC[-1],COUNTIF(R5C" & n_fty + 8 & ":RC[-1],R[-1]C[-1]),0)"
    Selection.Copy
    Selection.Resize(Selection.Rows.Count + ntgl - 1, Selection.Columns.Count).Select
    ActiveSheet.Paste
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues
    
    Range("Z5").End(xlToLeft).Offset(1, 2).Select
    Selection.Formula = "=IF(RC[-1]>0,RC[-6]-RC[-1],0)"
    Selection.Copy
    Selection.Resize(Selection.Rows.Count + ntgl - 1, Selection.Columns.Count).Select
    ActiveSheet.Paste
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues
    
    Range("Z5").End(xlToLeft).Offset(1, 3).Select
    Selection.Formula = _
        "=IF(RC[-2]>0,CONCATENATE(RC[-4],"" : "",TEXT(RC[-1],""MM""),""/"",TEXT(RC[-1],""DD""),""/"",TEXT(RC[-1],""YYYY""),"" - "",TEXT(R[-1]C[-7],""MM""),""/"",TEXT(R[-1]C[-7],""DD""),""/"",TEXT(R[-1]C[-7],""YYYY"")),0)"
    Selection.Copy
    Selection.Resize(Selection.Rows.Count + ntgl - 1, Selection.Columns.Count).Select
    ActiveSheet.Paste
    Selection.Copy
    Selection.PasteSpecial Paste:=xlPasteValues

    nnn = 1
    Range("A3").Select
    For nnn = 1 To nnweek
        Cells.Find(What:="WEEK" & nnn, After:=ActiveCell, LookIn:=xlFormulas, LookAt _
            :=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlNext, MatchCase:= _
            False, SearchFormat:=False).Activate
        Selection.EntireRow.Insert , CopyOrigin:=xlFormatFromLeftOrAbove
        Selection.Offset(1, 2).Select
        eachweek = Selection.Value
        If ActiveCell.Value = 1 Then
            If nnn = nnweek Then
                Selection.Offset(0, 2).Select
                ActiveCell.Formula = "=CONCATENATE(RC[-4],"" : "",TEXT(RC[-1],""MM""),""/"",TEXT(RC[-1],""DD""),""/"",TEXT(RC[-1],""YYYY""))"
                Selection.Offset(0, -2).Select
            Else
                Selection.Offset(0, 2).Select
                ActiveCell.Formula = "=CONCATENATE(RC[-4],"" : "",TEXT(R[-2]C[-1],""MM""),""/"",TEXT(R[-2]C[-1],""DD""),""/"",TEXT(R[-2]C[-1],""YYYY""),"" - "",TEXT(R[-2]C[-7],""MM""),""/"",TEXT(R[-2]C[-7],""DD""),""/"",TEXT(R[-2]C[-7],""YYYY""))"
                eachweek = ActiveCell.Offset(-2, -2).Value + 1
                Selection.Offset(0, -2).Select
            End If
        End If
        
        Selection.Offset(0, 2).Select
        weekname = Selection.Value
        Selection.End(xlToLeft).Select
        Selection.End(xlToLeft).Offset(-1, 0).Select
        Selection.Value = weekname
        Selection.Resize(Selection.Rows.Count, Selection.Columns.Count + 1).Select
        With Selection
            .VerticalAlignment = xlBottom
            .ReadingOrder = xlContext
            .MergeCells = True
        End With
        Selection.Offset(0, 1).Select
        Selection.Formula = "=SUM(R[-" & eachweek & "]C:R[-1]C)"
        Selection.Copy
        Selection.Resize(Selection.Rows.Count, Selection.Columns.Count + n_fty).Select
        ActiveSheet.Paste
        Selection.Copy
        Selection.PasteSpecial Paste:=xlPasteValues
    Next nnn
    
    Columns("A:Z").EntireColumn.AutoFit
    Columns("A:B").ColumnWidth = 17
    
    Range("A4").End(xlToRight).Offset(0, 1).Select
    Selection.Resize(Selection.Rows.Count, Selection.Columns.Count + 7).Select
    Selection.EntireColumn.Delete
    Range("A1048576").End(xlUp).Offset(0, n_fty + 1).Select
    Range(Selection, "A4").Select
    Selection.AutoFilter
    Range("A1048576").End(xlUp).Offset(-1, n_fty + 1).Select
    Range(Selection, "A5").Select
    ActiveSheet.Range("$A$4:$K$" & ntgl + nnweek + 6).AutoFilter Field:=1, Criteria1:="=*WEEK*", _
        Operator:=xlOr, Criteria2:="=*OVERTIME*"
    With Selection.Interior
        .Pattern = xlSolid
        .PatternColorIndex = xlAutomatic
        .Color = 65535
    End With
    ActiveSheet.ShowAllData
    
    Range("A1048576").End(xlUp).Offset(0, n_fty + 1).Select
    Range(Selection, "A4").Select
    ActiveSheet.Range("$A$4:$K$" & ntgl + nnweek + 6).AutoFilter Field:=1, Criteria1:="=*WEEK*", _
        Operator:=xlOr, Criteria2:="=*OVERTIME*"
    Selection.Copy
    Range("A1048576").End(xlUp).Offset(5, 0).Select
    ActiveSheet.Paste
    With Selection.Interior
        .Pattern = xlNone
        .TintAndShade = 0
        .PatternTintAndShade = 0
    End With
    Selection.AutoFilter
    
    Range("A1048576").End(xlUp).Offset(-nnweek - 3, 0).Select
    Selection.Value = "RESUME DATA OVERTIME MINGGUAN"
    Selection.Resize(Selection.Rows.Count, Selection.Columns.Count + n_fty + 2).Select
    Selection.Font.Bold = True
    With Selection
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlBottom
        .ReadingOrder = xlContext
        .MergeCells = True
    End With
End Sub