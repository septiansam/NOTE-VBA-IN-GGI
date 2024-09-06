Attribute VB_Name = "mdl_Process1"
'---------------------------------------------------------------------------------------*
'PUBLIC VARIABLE ------------------------------------------------------------------------
    Dim WB1 As Workbook
    Dim SH1_DM1 As Worksheet
    Dim LR1_DM1 As Long
    Dim SH1_RPA1 As Worksheet
    Dim LR1_RPA1 As Long
    '`.`.`.`.`.`.`.`.`.`.`.`.`.`.`.`.`.`.`.`.`.`.`.`.`.`.`.`.`.`.`.`.`.`.`.`.`.`.`.`.`.'
    
    Dim WB2 As Workbook
    Dim SH2_DM1 As Worksheet, SH2_X As Worksheet
    Dim LR2_DM1 As Long, LR2_X As Long
    '`.`.`.`.`.`.`.`.`.`.`.`.`.`.`.`.`.`.`.`.`.`.`.`.`.`.`.`.`.`.`.`.`.`.`.`.`.`.`.`.`.'
    
    Dim SumResult As Double, TotalResult As Double
    Dim dataRange As Range, CurrentCell  As Range, FoundCell As Range
    Dim NilaiCurrency As Currency
    Dim dataObj As New MSForms.DataObject
    Dim Tanggal As Date
    '`.`.`.`.`.`.`.`.`.`.`.`.`.`.`.`.`.`.`.`.`.`.`.`.`.`.`.`.`.`.`.`.`.`.`.`.`.`.`.`.`.'


Sub Proses1()
'---------------------------------------------------------------------------------------*
'[INISIALISASI]
''''''''''''''''''''''''''''''''''
    Set WB1 = ThisWorkbook
    Set SH1_HOME = WB1.Worksheets("HOME")

'    Set SH2_DM1 = WB2.ActiveSheet
'    LR1_DM1 = SH2_DM1.Range("A" & Rows.Count).End(xlUp).row
    
'    Set SH1_X = Sheets.Add(After:=Sheets(Sheets.Count))
'    SH1_X.Name = "DM1"
'    Set SH2_X = WB2.Worksheets("HEADER")

'    Set WB2 = Workbooks.Open(WB1.Sheets("DB_Dummy").Range("P8").Value)
'    Set SH2_X = WB2.ActiveSheet
'    SH2_X.AutoFilterMode = False
'    SH2_X.Cells.EntireColumn.Hidden = False
'    LR2_X = SH2_X.Range("B" & Rows.Count).End(xlUp).row
    

'[CREATE RPA1]
''''''''''''''''''''''''''''''''''
    '[*]COMMENT NOTE:
    
    '[*]Ref_Konversi SOURCE:
    
    '[*]DEV:
    '.....^]Fromulasi
    '.....^]CoPas
    
    
    
'[PENUTUP TEMP_1]
''''''''''''''''''''''''''''''''''
    WB2.Close (False)
    SH1_RPA1.Activate
    SH1_RPA1.Cells.EntireColumn.AutoFit
End Sub
