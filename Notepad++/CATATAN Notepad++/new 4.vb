'[BUAT PIVOT KE SHEET TEMP1]................
Dim PV_RANGE As Range
Dim PV_TABLE As PivotTable
Dim PV_CACHE As PivotCache

Set PV_RANGE = OLAH.Range(Cells(1, 1), Cells(LR, LC))
Set PV_CACHE = TWB.PivotCaches.Create(SourceType:=xlDatabase, _
                    SourceData:=PV_RANGE)
Set PV_TABLE = PV_CACHE.CreatePivotTable _
                    (TableDestination:=TEMP1.Range("A1"), _
                    TableName:="PV_ANALISA")
TEMP1.Activate

'[INSERT FIELD FIELD NYA]...
With PV_TABLE.PivotFields("Buyer")
    .Caption = "Buyer"
    .Orientation = xlRowField
    .Position = 1
End With

With PV_TABLE.PivotFields("KATEGORI")
    .Caption = "Kategori    (Days)"
    .Orientation = xlRowField
    .Position = 2
End With

With PV_TABLE.PivotFields("G/L Cat")
    .Orientation = xlColumnField
    .Position = 1
End With

'[ISI].....
With PV_TABLE.PivotFields("Amount")
    .Orientation = xlDataField
    .Position = 1
    .Function = xlSum
End With

'[SETTING]...
With PV_TABLE
    '.RowGrand = False
    .DisplayErrorString = False
    .NullString = 0
    .PageFieldOrder = 2
    .PreserveFormatting = True
    .PrintTitles = False
    .CompactRowIndent = 1
    .DisplayContextTooltips = True
    .ShowDrillIndicators = True
    .PrintDrillIndicators = False
    .AllowMultipleFilters = False
    .SortUsingCustomLists = True
    .FieldListSortAscending = False
    .ShowValuesRow = False
    .RowAxisLayout xlTabularRow
    .PivotCache.MissingItemsLimit = xlMissingItemsDefault
    .TableStyle2 = "PivotStyleDark7"
End With

LR = TEMP1.Cells.Find(What:="*" _
    , LookAt:=xlPart _
    , LookIn:=xlFormulas _
    , SearchOrder:=xlByRows _
    , SearchDirection:=xlPrevious).Row

LC = TEMP1.Cells.Find(What:="*" _
    , LookAt:=xlPart _
    , LookIn:=xlFormulas _
    , SearchOrder:=xlByColumns _
    , SearchDirection:=xlPrevious).Column

TEMP1.Range(Cells(1, 1), Cells(LR, LC)).Copy
TEMP2.Range("A1").PasteSpecial xlPasteValuesAndNumberFormats: Application.CutCopyMode = False
