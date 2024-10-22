Attribute VB_Name = "mdl_CopyPdfContentToExcel"
'Tools > References : Microsoft Forms 2.0 Object Library


Sub CopyPDFContentToExcel()
    Dim pdfPath As String
    Dim shellApp As Object
    Dim objShell As Object
    Dim clipboardData As Object
    Dim WB As Workbook
    Dim WS As Worksheet
    
    ' Tentukan path file PDF
    pdfPath = "\\10.8.0.35\Bersama\IT\Macro Record Projects\Local Project\MACRO_InputUser - Partial 262 at GCC\.Macro\Backup\LOGIKA ALGORITMA\20241016\3\INV 33 ACC SEWING 24001061 CLN.pdf"
    If pdfPath = "False" Then Exit Sub
    
    ' Set Workbook dan Worksheet
    Set WB = ThisWorkbook
    Set WS = WB.Sheets("Sheet5")
    
    ' Jalankan aplikasi PDF (misalnya, Adobe Reader)
    Set shellApp = CreateObject("Shell.Application")
    shellApp.Open (pdfPath)
    ' Beri waktu agar aplikasi PDF terbuka
    Application.Wait Now + TimeValue("00:00:03")
    
    ' Pastikan aplikasi Adobe Reader berada di depan (aktif)
    Set objShell = CreateObject("WScript.Shell")
    objShell.AppActivate "Adobe Acrobat Reader"
    Application.Wait Now + TimeValue("00:00:02")
    
    ' Copy Data
    objShell.SendKeys "^a", True
    Application.Wait Now + TimeValue("00:00:01")
    objShell.SendKeys "^c", True
    Application.Wait Now + TimeValue("00:00:02")
    
    ' Tutup aplikasi PDF setelah teks disalin
    objShell.SendKeys "%{F4}"
    Application.Wait Now + TimeValue("00:00:02")
    
    ' Tempelkan teks dari clipboard ke sel A1 di Sheet1
    WB.Activate
    WS.Activate
    WS.Range("A1").Select
    objShell.SendKeys "^v", True
End Sub

