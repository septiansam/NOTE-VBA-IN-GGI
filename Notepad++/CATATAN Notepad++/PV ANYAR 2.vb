    'PIVOT
    Set pv_Rng = Range(Cells(1, 1), Cells(LR, LC))
    Set pv_Cache = TWB.PivotCaches.Create(SourceType:=xlDatabase, _
                        SourceData:=pv_Rng)
    Set pv_Tb = pv_Cache.CreatePivotTable _
                        (TableDestination:=TEMP3.Range("A1"), _
                        TableName:="pv_table")
    
    TEMP3.Activate
    With pv_Tb.PivotFields("UNIT")
        .Caption = "UNIT"
        .Orientation = xlRowField
        .Position = 1
        .Subtotals = _
            Array(False, False, False, False, False, False, False, False, False, False, False, False)
    End With
    
    With pv_Tb.PivotFields("MONTH NUMBER")
        .Caption = "ID"
        .Orientation = xlRowField
        .Position = 2
        .Subtotals = _
            Array(False, False, False, False, False, False, False, False, False, False, False, False)
    End With
    
    With pv_Tb.PivotFields("MONTH NAME")
        .Caption = "MONTH"
        .Orientation = xlRowField
        .Position = 3
        .Subtotals = _
            Array(False, False, False, False, False, False, False, False, False, False, False, False)
    End With
    
    With pv_Tb.PivotFields("WEEK")
        .Caption = "WEEK"
        .Orientation = xlRowField
        .Position = 4
        .Subtotals = _
            Array(False, False, False, False, False, False, False, False, False, False, False, False)
    End With
    
    With pv_Tb.PivotFields("TEMUAN MESIN")
        .Orientation = xlColumnField
        .Position = 1
    End With
    
    With pv_Tb.PivotFields("TEMUAN MESIN")
        .Orientation = xlDataField
        .Position = 1
        .Function = xlCount
    End With
    
    With pv_Tb
        .RowAxisLayout xlTabularRow
        .RowGrand = False
        .ShowValuesRow = False
        .DisplayErrorString = False
        .RepeatAllLabels xlRepeatLabels
    End With
    
    LR = Cells.Find("*", , xlFormulas, xlPart, xlByRows, xlPrevious).Row
    LC = Cells.Find("*", , xlFormulas, xlPart, xlByColumns, xlPrevious).Column
    
    With Range(Cells(1, 1), Cells(LR, LC))
        .Copy: .PasteSpecial xlPasteValuesAndNumberFormats: Application.CutCopyMode = False: Cells.EntireColumn.AutoFit: Cells(1, 1).Select
    End With
