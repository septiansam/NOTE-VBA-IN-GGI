Option Explicit
'_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_'
'''''-----------------------------------------------------'''''

Dim TWB As Workbook, WB_TARIKAN As Workbook, WB_RESULTS As Workbook
Dim SH_HOME As Worksheet, SH_TARIKAN As Worksheet, SH_HELP As Worksheet
Dim TEMP1 As Worksheet, TEMP2 As Worksheet, RESULTS As Worksheet
Dim PATH_TARIKAN As String, PATH_RESULTS_EXCEL As String, PATH_RESULTS_PDF As String
Dim RNG As Range, RNG_BORDER As Range, CELL As Range
Dim LR_TARIKAN As Long, LC_TARIKAN As Long, LR As Long, LC As Long, FR As Long, FC As Long
Dim COL_REF As Long
Dim i As Long, j As Long, x As Long, COL_PASTE As Long
Dim RNG_PERIODE As Range, PERIODE_AWAL As Date, PERIODE_AKHIR As Date, TITLE As String
Dim BULAN_AWAL As String, BULAN_AKHIR As String
Dim RNG_ROW As Range, RNG_COLUMN As Range, RNG_RESULTS As Range

'''''-----------------------------------------------------'''''
'_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_'