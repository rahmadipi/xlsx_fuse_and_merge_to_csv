@echo off
mode con: cols=52 lines=30
setlocal enabledelayedexpansion
cls

title Fusion Tool v1.1
echo ===================================================
echo    Skrip Otomatis Konversi dan Penggabungan CSV 
echo ===================================================
echo.
echo ===================================================
echo [1]  Memulai proses konversi Excel ke CSV...
echo ===================================================
echo.

cscript.exe //nologo run_macro.vbs "%~dp0ConvertToCSV.xlsm" "ConvertXLS_to_CSV_From_Script"

if %errorlevel% neq 0 (
    color 0C
    echo.
    echo ===================================================
    echo     Proses Macro gagal. Konversi dihentikan.
    echo ===================================================
    echo.
    pause
    goto :eof
)

echo.
echo ===================================================
echo [2]  Memulai proses penggabungan file CSV...
echo ===================================================
echo.

set "output_file=file_gabungan.csv"

rem Hapus file output jika sudah ada
if exist "%output_file%" del "%output_file%"

rem Ambil nama file pertama yang bukan file output
for /f "delims=" %%f in ('dir /b /a-d-h "*.csv" ^| findstr /i /v /c:"%output_file%"') do (
    set "first_file=%%f"
    goto :process_first
)

color 0C
echo.
echo ===================================================
echo         Tidak ada file CSV yang ditemukan
echo  Pastikan file Excel telah dikonversi dengan benar
echo ===================================================
echo.
pause
goto :eof

:process_first
echo Menggabungkan file...

rem Proses file pertama secara terpisah (dengan header)
type "%first_file%" > "%output_file%"
del "%first_file%"

rem Proses semua file lainnya tanpa header
for %%f in (*.csv) do (
    rem Lewati file pertama dan file output
    if /i not "%%f"=="%first_file%" (
        if /i not "%%f"=="%output_file%" (
            rem mulai dari baris ke 1
            more +1 "%%f" >> "%output_file%"
            rem hapus csv setelah ditempel ke file output
            del "%%f"
        )
    )
)

echo Penggabungkan selesai...

rem start of opsi pilihan
color 0E
echo.
echo ===================================================
echo Lanjut konversi hasil ke file XLSX ? (y/n)
echo [Otomatis lanjut dalam 10 detik...]
choice /c yn /d y /t 10 /n /m "Tekan Y untuk Lanjut atau N untuk Berhenti: "
echo ===================================================
if errorlevel 2 goto :skipxls
if errorlevel 1 goto :lanjut
rem end of opsi pilihan

:lanjut
echo.
echo ===================================================
echo [3]  Mengonversi hasil gabungan ke Excel...
echo ===================================================
echo.

set "output_file2=file_gabungan.xlsx"

rem Hapus file output jika sudah ada
if exist "%output_file2%" del "%output_file2%"

cscript.exe //nologo run_macro.vbs "%~dp0ConvertToCSV.xlsm" "ConvertFinalCSV_to_XLSX"

if %errorlevel% neq 0 (
    color 0C
    echo.
    echo ===================================================
    echo     Proses Macro gagal. Konversi dihentikan.
    echo ===================================================
    echo.
    pause
    goto :eof
)

rem Hapus file CSV gabungan
if exist "%output_file%" del "%output_file%"

set "output_file=%output_file2%"

:skipxls
cls
echo.
color 0A
echo ===================================================
echo          Proses penggabungan berhasil!
echo   Harap untuk memeriksa file hasil penggabungan
echo ===================================================
echo      *File output: %output_file%
echo ===================================================
echo.
pause