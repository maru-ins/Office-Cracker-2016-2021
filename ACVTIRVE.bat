@echo off
title Ruins Office Activator

:: Logo ASCII
echo ···············································
echo : ______   __    __ _____     __      _  _____ :
echo :(   __ \  ) )  ( ((_   _)   /  \    / )/ ____\:
echo : ) (__) )( (    ) ) | |    / /\ \  / /( (___  :
echo :(    __/  ) )  ( (  | |    ) ) ) ) ) ) \___ \ :
echo : ) \ \  _( (    ) ) | |   ( ( ( ( ( (      ) ):
echo :( ( \ \_))) \__/ ( _| |__ / /  \ \/ /  ___/ / :
echo : )_) \__/ \______//_____((_/    \__/  /____/  :
echo ···············································
echo.

:: Menu pilihan
echo Pilih versi Office yang akan diaktivasi:
echo [1] Office 2016

set /p pilihan=Masukkan pilihan [1-4] : 

:: Proses pilihan
if "%pilihan%"=="1" set "batchfile=crack_office_2016.bat"

:: Cek file batch
if not exist "%batchfile%" (
    echo.
    echo File "%batchfile%" tidak ditemukan!
    pause
    exit /b
)

:: Konfirmasi
set /p jawab=Lanjutkan aktivasi %batchfile%? (y/n) : 
if /i not "%jawab%"=="y" (
    echo Dibatalkan oleh user.
    pause
    exit /b
)

:: Jalankan file batch
echo.
echo Menjalankan %batchfile%...
call "%batchfile%"
pause
