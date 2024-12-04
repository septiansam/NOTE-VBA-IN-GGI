@echo off
rem Periksa apakah direktori Speedtest CLI tersedia
cd /d C:\ooklaspeedtest >nul 2>&1
if %errorlevel% neq 0 (
    echo Direktori Speedtest CLI tidak ditemukan. Pastikan "C:\ooklaspeedtest" tersedia.
    exit /b
)

rem Periksa apakah Speedtest CLI tersedia
speedtest.exe --version >nul 2>&1
if %errorlevel% neq 0 (
    echo Speedtest CLI tidak ditemukan. Pastikan speedtest.exe ada di direktori C:\ooklaspeedtest.
    exit /b
)

rem Hapus file lama jika ada
set output="\\10.8.0.35\Bersama\IT\RPA IT\Internet Speed Test\1201_SpeedTestResults.txt"
if exist %output% del %output%

rem Jalankan Speedtest dan alihkan output ke file sementara
echo Testing speed... Mohon tunggu beberapa saat.
speedtest.exe --server-id=7580 --unit=Mbps > temp.txt

rem Tambahkan date/time ke hasil
set datetime=%date% %time%
echo Date/Time: %datetime% > %output%
type temp.txt >> %output%
del temp.txt

rem Tampilkan hasil di Notepad
start notepad.exe %output%

rem Tunggu beberapa detik sebelum menutup Notepad
timeout /t 5 >nul

rem Tutup Notepad
taskkill /im notepad.exe /f >nul 2>&1

rem Tutup CMD
exit
