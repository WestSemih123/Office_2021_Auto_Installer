@echo off
setlocal enabledelayedexpansion
title Office 2021 Install

echo Please choose language.
choice /C TE /N /M "Turkish (T) or English (E)"

echo ============================================================================&echo.&echo.

if errorlevel 2 (
    set lan=en-us
) else (
    set lan=tr-TR
)

if "%lan%"=="en-us" (
    echo You have selected English language.
) else (
    echo Turkce dilini sectiniz.
)

echo ============================================================================&echo.&echo.

set arch=!PROCESSOR_ARCHITECTURE!
if "!arch!"=="AMD64" (
    set arch=64
) else if "!arch!"=="x86" (
    set arch=32
) else (
    if "%lan%"=="en-us" (
        echo We could not detect the processor architecture of your computer. Please choose your processor architecture.
    ) else (
        echo Bilgisayarinizin islemci mimarisini tespit edemedik. Lütfen islemci mimarisini secin.
    )
    choice /C AB /N /M "64 (A) or 32 (B)"
    if errorlevel 1 (
        set arch=64
    ) else (
        set arch=32
    )
)

if "%lan%"=="en-us" (
    echo This computer uses a %arch%-bit operating system.
) else (
    echo Bu bilgisayar %arch%-bit isletim sistemini kullaniyor.
)

echo ============================================================================&echo.&echo.

if "%lan%"=="en-us" (
    echo Select your installation option.
) else (
    echo Yukleme seceneginizi secin
)

choice /C FW /N /M "Full_Install (F) or Word_Excel_PowerPoint_Install (W)"

echo ============================================================================&echo.&echo.

if errorlevel 2 (
    set xmlfile=w
    echo ^<Configuration^> >> Word_Excel_PowerPoint.xml
    echo   ^<Add OfficeClientEdition="!arch!" Channel="PerpetualVL2021"^> >> Word_Excel_PowerPoint.xml
    echo     ^<Product ID="ProPlus2021Volume"^> >> Word_Excel_PowerPoint.xml
    echo       ^<Language ID="!lan!" /^> >> Word_Excel_PowerPoint.xml
    echo       ^<ExcludeApp ID="Access" /^> >> Word_Excel_PowerPoint.xml
    echo       ^<ExcludeApp ID="Lync" /^> >> Word_Excel_PowerPoint.xml
    echo       ^<ExcludeApp ID="OneDrive" /^> >> Word_Excel_PowerPoint.xml
    echo       ^<ExcludeApp ID="OneNote" /^> >> Word_Excel_PowerPoint.xml
    echo       ^<ExcludeApp ID="Publisher" /^> >> Word_Excel_PowerPoint.xml
    echo       ^<ExcludeApp ID="Teams" /^> >> Word_Excel_PowerPoint.xml
    echo       ^<ExcludeApp ID="Outlook" /^> >> Word_Excel_PowerPoint.xml
    echo     ^</Product^> >> Word_Excel_PowerPoint.xml
    echo   ^</Add^> >> Word_Excel_PowerPoint.xml
    echo   ^<Remove All="True" /^> >> Word_Excel_PowerPoint.xml
    echo ^</Configuration^> >> Word_Excel_PowerPoint.xml
    set ver=Word_Excel_PowerPoint
) else (
    set xmlfile=f
    echo ^<Configuration^> >> Full.xml
    echo   ^<Add OfficeClientEdition="!arch!" Channel="PerpetualVL2021"^> >> Full.xml
    echo     ^<Product ID="ProPlus2021Volume"^> >> Full.xml
    echo       ^<Language ID="!lan!" /^> >> Full.xml
    echo       ^<ExcludeApp ID="Lync" /^> >> Full.xml
    echo     ^</Product^> >> Full.xml
    echo   ^</Add^> >> Full.xml
    echo   ^<Remove All="True" /^> >> Full.xml
    echo ^</Configuration^> >> Full.xml
    set ver=Full
)

if "%ver%"=="Full" (
    call powershell.exe -Command .\setup.exe /configure .\Full.xml
) else (
    call powershell.exe -Command .\setup.exe /configure .\Word_Excel_PowerPoint.xml
)

echo ============================================================================&echo.&echo.

if "%lan%"=="en-us" (
    echo Your office is installing.
) else (
    echo Office kurulumu yapiliyor.
)

echo ============================================================================&echo.&echo.

:CHECK_INSTALLATION
timeout /t 10 >nul
tasklist /FI "IMAGENAME eq setup.exe" 2>NUL | find /I /N "setup.exe">NUL
if "%ERRORLEVEL%"=="0" (
    goto CHECK_INSTALLATION
) else (
    if "%lan%"=="en-us" (
        echo Installation completed.
    ) else (
        echo Yukleme tamamlandi.
        if "%lan%"=="en-us" (
        msg * Please enter your license key from within the app.
    ) else (
        msg * Lutfen lisans anahtarinizi uygulama icinden girin.
        if "%xmlfile%"=="w"(
            del Word_Excel_PowerPoint.xml
        ) else (
            del Full.xml
        )
    )
    )
    pause
)
