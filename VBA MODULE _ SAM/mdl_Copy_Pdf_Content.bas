Attribute VB_Name = "mdl_Copy_Pdf_Content"
'Tools > References : Microsoft Forms 2.0 Object Library


Sub CopyPDFContentToExcel(PathFilePDF As String, SheetTujuan As Worksheet)
    Dim pdfPath As String
    Dim shellApp As Object
    Dim objShell As Object
    Dim clipboardData As Object
    Dim ws As Worksheet
    
    ' Tentukan path file PDF
    pdfPath = PathFilePDF
    If pdfPath = "False" Then Exit Sub
    
    ' Set Worksheet
    Set ws = SheetTujuan
    
    ' Jalankan aplikasi PDF (misalnya, Adobe Reader)
    Set shellApp = CreateObject("Shell.Application")
    shellApp.Open (pdfPath)
    ' Beri waktu agar aplikasi PDF terbuka
    Application.Wait Now + TimeValue("00:00:01")
    
    ' Pastikan aplikasi Adobe Reader berada di depan (aktif)
    Set objShell = CreateObject("WScript.Shell")
    objShell.AppActivate "Adobe Acrobat Reader"
    Application.Wait Now + TimeValue("00:00:01")
    
    ' Copy Data
    objShell.SendKeys "^a", True
    Application.Wait Now + TimeValue("00:00:01")
    objShell.SendKeys "^c", True
    Application.Wait Now + TimeValue("00:00:01")
    
    ' Tutup aplikasi PDF setelah teks disalin
    objShell.SendKeys "%{F4}"
    Application.Wait Now + TimeValue("00:00:01")
    
    ' Tempelkan teks dari clipboard ke sel A1 di Sheet1
    ws.Activate
    ws.Range("A1").Select
    ws.Paste
    ws.Cells.EntireColumn.AutoFit
    ws.Range("A1").Select
    Application.Wait Now + TimeValue("00:00:02")
End Sub
