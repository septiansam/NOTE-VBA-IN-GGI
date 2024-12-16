Set wb_SI = Workbooks.Open(path_SI): wb_SI.Activate: Sheets(1).Select: ActiveSheet.AutoFilterMODE = False: Cells.EntireColumn.Hidden = False: Cells.EntireRow.Hidden = False

: Sheets(1).Select: ActiveSheet.AutoFilterMode = False: Cells.EntireColumn.Hidden = False: Cells.EntireRow.Hidden = False: Cells.EntireColumn.AutoFit


set wb_src = Workbooks.Open(FullPathFile):wb_Src.Activate:sheets(1).select:ActiveSheet.autofiltermode = false:cells.EntireColumn.Hidden = false:cells.EntireRow.Hidden = false

sheets(1).select:ActiveSheet.autofiltermode = false:cells.EntireColumn.Hidden = false:cells.EntireRow.Hidden = false

.Copy: .PasteSpecial xlPasteValuesAndNumberFormats: Application.CutCopyMode = False: Cells.EntireColumn.AutoFit: Cells(1, 1)Select


LR = Cells.Find("*", , xlFormulas, xlPart, xlByRows, xlPrevious).Row
LC = Cells.Find("*", , xlFormulas, xlPart, xlByColumns, xlPrevious).Column
    
With rng
    .Borders(xlEdgeTop).LineStyle = xlContinuous
    .Borders(xlEdgeTop).Color = RGB(0, 108, 105)
    .Borders(xlEdgeBottom).LineStyle = xlContinuous
    .Borders(xlEdgeBottom).Color = RGB(0, 108, 105)
    .Borders(xlInsideHorizontal).LineStyle = xlContinuous
    .Borders(xlInsideHorizontal).Color = RGB(0, 108, 105)
End With
	
	
Function wsx(sh_Name As String) As Boolean
    On Error Resume Next
        wsx = Not Sheets(sh_Name) Is Nothing
    On Error GoTo 0
End Function

