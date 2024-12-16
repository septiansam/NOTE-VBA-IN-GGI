SH_DB.Activate
NAME_RESUME = Range("L9").Value
PATH_RESUME = Range("L10") & Application.PathSeparator & NAME_RESUME & ".xlsx"

SH_InputUser.Activate

SH_InputUser.Copy
Application.DisplayAlerts = False
Set WB_RESUME = ActiveWorkbook
WB_RESUME.Activate
Sheets(1).Select
ActiveSheet.Name = "RESUME"
Range("C:C").Insert
Range("U:U").Copy Range("C:C")
Range("U:U").ClearContents
Cells.EntireColumn.AutoFit

Cells.Font.Name = "Verdana"
Cells.VerticalAlignment = xlCenter

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
    .Interior.Color = RGB(8, 43, 62)
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

LR = ActiveSheet.Cells.Find(What:="*", LookAt:=xlPart, LookIn:=xlFormulas, SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
LC = ActiveSheet.Cells.Find(What:="*", LookAt:=xlPart, LookIn:=xlFormulas, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious).Column

Set RNG_RESUME = Range(Cells(1, 1), Cells(LR, LC))
For Each COL In RNG_RESUME.Columns
    COL.EntireColumn.AutoFit
    COL.ColumnWidth = COL.ColumnWidth + 1
Next COL

Cells.HorizontalAlignment = xlCenter
Cells.VerticalAlignment = xlCenter
Rows("3:3").Select
ActiveWindow.FreezePanes = True
Cells(1, 1).Select

WB_RESUME.SaveAs PATH_RESUME, xlOpenXMLWorkbook
WB_RESUME.Close (True)