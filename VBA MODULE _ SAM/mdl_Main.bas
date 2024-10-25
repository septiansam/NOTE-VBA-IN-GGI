Attribute VB_Name = "mdl_Main"
'_____________________________________________________________________________________________________
'## MACRO_RPA - Report Open PO VS EX Factory By OR   -_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-
'## DEVELOPER := Septian Arif Maulana -_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_
'-----------------------------------------------------------------------------------------------------

'=====================================================================================================
'=> -----------------------------------VARIABLE INITIALIZATION------------------------------------- <=
'=====================================================================================================
    Public WB1 As Workbook
    Public SH1_Home As Worksheet
    
    Public Rng As Range, Cell As Range
    Public LR1 As Long, LC1 As Long
    Public i As Long, j As Long
    '-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-_-
    
    
'+----------------------------------------------------------------------------------------------------+

Sub BTN_PROSES1()
    Application.AskToUpdateLinks = False
    Application.DisplayAlerts = False
    'Application.DisplayStatusBar = False
    Application.EnableEvents = False
    'Application.DisplayScrollBars = False




    Application.AskToUpdateLinks = True
    Application.DisplayAlerts = True
    'Application.DisplayStatusBar = True
    Application.EnableEvents = True
    'Application.DisplayScrollBars = True
End Sub
