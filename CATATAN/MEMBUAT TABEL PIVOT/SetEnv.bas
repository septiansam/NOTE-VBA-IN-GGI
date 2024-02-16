Attribute VB_Name = "SetUpEnv"
Option Explicit

Sub SetUpEnvPivot(kondisi As Boolean)
Attribute SetUpEnvPivot.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Setup Environment Pivot Macro
' By Septian
'
'

'
    Application.CutCopyMode = False
    With Application.DefaultPivotTableLayoutOptions
        If kondisi = True Then
            'Menyembunyikan zona penjatuh dalam tabel pivot.
            .InGridDropZones = False
            'Menampilkan judul kolom untuk setiap kolom dalam tabel pivot.
            .DisplayFieldCaptions = True
            'Mengurutkan item dalam daftar bidang tabel pivot secara menurun.
            .FieldListSortAscending = False
            'Menampilkan anggota yang dihitung dalam tabel pivot.
            .ViewCalculatedMembers = True
            'Menampilkan info tooltip untuk item dalam tabel pivot.
            .DisplayContextTooltips = True
            'Menampilkan indikator pengeboran untuk item dalam tabel pivot.
            .ShowDrillIndicators = True
            'Menyembunyikan baris kosong dalam tabel pivot.
            .DisplayEmptyRow = False
            'Menampilkan info tooltip untuk properti anggota dalam tabel pivot.
            .DisplayMemberPropertyTooltips = True
            'Menyembunyikan baris nilai dalam tabel pivot.
            .ShowValuesRow = False
            'Menampilkan string kosong untuk nilai null dalam tabel pivot.
            .DisplayNullString = True
            'Menyembunyikan string kesalahan dalam tabel pivot.
            .DisplayErrorString = False
            'Mengaktifkan penggunaan format otomatis untuk tabel pivot.
            .HasAutoFormat = True
            'Mengatur urutan bidang halaman dalam tabel pivot.
            .PageFieldOrder = True
            'Menyatukan label dalam tabel pivot.
            .MergeLabels = False
            'Mempertahankan format saat mengubah tata letak tabel pivot.
            .PreserveFormatting = True
            'Menyembunyikan indikator pengeboran saat mencetak tabel pivot.
            .PrintDrillIndicators = False
            'Mengulangi item pada setiap halaman cetak tabel pivot.
            .RepeatItemsOnEachPrintedPage = True
            'Menyembunyikan judul saat mencetak tabel pivot.
            .PrintTitles = False
            'Mencegah penggunaan beberapa filter pada tabel pivot.
            .AllowMultipleFilters = False
            'Mengizinkan anggota yang dihitung dalam filter tabel pivot.
            .CalculatedMembersInFilters = True
            'Menampilkan total visual untuk set dalam tabel pivot.
            .VisualTotalsForSets = True
            'Menyembunyikan total visual dalam tabel pivot.
            .VisualTotals = False
            'Menampilkan catatan total dalam tabel pivot.
            .TotalsAnnotation = True
            'Menyembunyikan total baris dalam tabel pivot.
            .RowGrand = False
            'Menampilkan total kolom dalam tabel pivot.
            .ColumnGrand = True
            'Menyembunyikan item halaman subtotal dalam tabel pivot.
            .SubtotalHiddenPageItems = True
            'Menggunakan daftar kustom saat mengurutkan dalam tabel pivot.
            .SortUsingCustomLists = True
            'Menyimpan data saat menyimpan tabel pivot.
            .SaveData = True
            'Mengaktifkan fungsi pengeboran dalam tabel pivot.
            .EnableDrilldown = True
            'Menonaktifkan pembaruan otomatis tabel pivot saat membuka file.
            .RefreshOnFileOpen = False
            'Menyembunyikan subtotal dalam tabel pivot.
            .Subtotals = False
            'Menyembunyikan lokasi subtotal dalam tabel pivot.
            .SubtotalLocation = False
            'Menyembunyikan baris kosong dalam tata letak tabel pivot.
            .LayoutBlankLine = False
            'Mengatur pengaturan item yang hilang ke "Tidak Ada" dalam tabel pivot.
            .xlMissingItemsNone = -1
            'Mengatur jumlah item bidang halaman yang dibungkus dalam tabel pivot.
            .PageFieldWrapCount = 1
            'Mengatur jarak indentasi baris dalam tabel pivot yang dikompakkan.
            .CompactRowIndent = 0
            'Mengatur tata letak sumbu baris dalam tabel pivot.
            .RowAxisLayout = 1
            'Menampilkan item segera dalam tabel pivot.
            .DisplayImmediateItems = True
            'Menonaktifkan penulisan kembali (writeback) dalam tabel pivot.
            .EnableWriteback = False
            'Mengatur pengulangan label pada setiap halaman cetak dalam tabel pivot.
            .RepeatAllLabels = xlRepeatLabels
        Else
            .InGridDropZones = False
            .DisplayFieldCaptions = True
            .FieldListSortAscending = False
            .ViewCalculatedMembers = True
            .DisplayContextTooltips = True
            .ShowDrillIndicators = True
            .DisplayEmptyRow = False
            .DisplayMemberPropertyTooltips = True
            .ShowValuesRow = False
            .DisplayNullString = True
            .DisplayErrorString = False
            .HasAutoFormat = True
            .PageFieldOrder = True
            .MergeLabels = False
            .PreserveFormatting = True
            .PrintDrillIndicators = False
            .RepeatItemsOnEachPrintedPage = True
            .PrintTitles = False
            .AllowMultipleFilters = False
            .CalculatedMembersInFilters = True
            .VisualTotalsForSets = True
            .VisualTotals = False
            .TotalsAnnotation = True
            .RowGrand = True
            .ColumnGrand = True
            .SubtotalHiddenPageItems = True
            .SortUsingCustomLists = True
            .SaveData = True
            .EnableDrilldown = True
            .RefreshOnFileOpen = False
            .Subtotals = True
            .SubtotalLocation = True
            .LayoutBlankLine = False
            .xlMissingItemsNone = -1
            .PageFieldWrapCount = 1
            .CompactRowIndent = 0
            .RowAxisLayout = 0
            .DisplayImmediateItems = True
            .EnableWriteback = False
            .RepeatAllLabels = xlRepeatLabels
        End If
    End With
    
End Sub
