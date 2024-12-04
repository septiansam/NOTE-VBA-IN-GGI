@echo off
rem Periksa apakah Python sudah diinstal
python --version >nul 2>&1
if %errorlevel% neq 0 (
    echo Python tidak ditemukan. Silakan instal Python terlebih dahulu.
    pause
    exit /b
)

rem Jalankan speedtest dan simpan hasil ke lokasi jaringan
echo Testing speed... Mohon tunggu beberapa saat.
python -c "import speedtest; import datetime; 
st = speedtest.Speedtest(); 
st.get_best_server(); 
st.download(); 
st.upload(); 
now = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S'); 
results = f'Date/Time: {now}\nDownload: {st.results.download / 1e6:.2f} Mbps\nUpload: {st.results.upload / 1e6:.2f} Mbps\nPing: {st.results.ping:.2f} ms'; 
print(results); 
open(r'\\\\10.8.0.35\\Bersama\\IT\\RPA IT\\Internet Speed Test\\1201_SpeedTestResults.txt', 'w').write(results)" > nul

if exist "\\10.8.0.35\Bersama\IT\RPA IT\Internet Speed Test\1201_SpeedTestResults.txt" (
    echo Speed test selesai. Hasil disimpan di \\10.8.0.35\Bersama\IT\RPA IT\Internet Speed Test\1201_SpeedTestResults.txt
    notepad "\\10.8.0.35\Bersama\IT\RPA IT\Internet Speed Test\1201_SpeedTestResults.txt"
) else (
    echo Gagal menyimpan hasil tes. Periksa akses ke lokasi jaringan.
)

pause