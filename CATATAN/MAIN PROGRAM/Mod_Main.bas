Attribute VB_Name = "Mod_Main"
Option Explicit

Sub Main()


OptVBA True
'' program dimulai ''

' ==================================================================================



' ==================================================================================
'' program berakhir ''
ThisWorkbook.Worksheets(1).Activate
ActiveSheet.Cells(1, 1).Select
OptVBA False

ActiveWorkbook.Save
End Sub

Private Sub OptVBA(isOn As Boolean)
    With Application
        .Calculation = IIf(isOn, xlCalculationManual, xlCalculationAutomatic)
        .EnableEvents = Not (isOn)
        .ScreenUpdating = Not (isOn)
        .DisplayAlerts = Not (isOn)
    End With
    ActiveSheet.DisplayPageBreaks = Not (isOn)
End Sub

