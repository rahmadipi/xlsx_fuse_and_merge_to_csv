@echo off
setlocal enabledelayedexpansion
cls

echo ===================================================
echo    Skrip Otomatis Konversi dan Penggabungan CSV 
echo ===================================================
echo.
echo [1/2]  Memulai proses konversi Excel ke CSV...
echo.

cscript.exe run_macro.vbs "%~dp0ConvertToCSV.xlsm" "ConvertXLS_to_CSV_From_Script"

timeout /t 5 >nul

echo.
echo [Selesai] Konversi Excel telah selesai.
echo ===================================================
echo [2/2]  Memulai proses penggabungan file CSV...
echo.

set "output_file=file_gabungan.csv"

rem Hapus file output jika sudah ada
if exist "%output_file%" del "%output_file%"

rem Ambil nama file pertama yang bukan file output
for /f "delims=" %%f in ('dir /b /a-d-h "*.csv" ^| findstr /i /v /c:"%output_file%"') do (
    set "first_file=%%f"
    goto :process_first
)

echo.
echo    Tidak ada file CSV yang ditemukan untuk digabungkan.
echo    Pastikan file Excel telah dikonversi dengan benar.
echo.
pause
goto :eof

:process_first
echo Menggabungkan file...

rem Proses file pertama secara terpisah (dengan header)
type "%first_file%" > "%output_file%"

rem Proses semua file lainnya tanpa header
for %%f in (*.csv) do (
    rem Lewati file pertama dan file output
    if /i not "%%f"=="%first_file%" (
        if /i not "%%f"=="%output_file%" (
            more +1 "%%f" >> "%output_file%"
        )
    )
)

echo.
echo ===================================================
echo              Penggabungan selesai!
echo ===================================================
echo.
echo File output: %output_file%
echo.

echo.
echo ===================================================
echo   Harap untuk memeriksa file hasil penggabungan
echo ===================================================
pause