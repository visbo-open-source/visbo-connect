@echo off
setlocal EnableDelayedExpansion
cls

echo ===============================================
echo   VISBO SPE + Connect Installation
echo   (Custom Build Version)
echo ===============================================
echo.

REM ============ Determine Git Repository Path ============
REM This script is located in: <git-ordner>\visbo-connect\docs\admin-doc
REM So git-ordner is 3 levels up

set "SCRIPT_DIR=%~dp0"
pushd "%SCRIPT_DIR%"
cd ..\..\..
set "GIT_ROOT=%CD%"
popd

echo Ermittelter Git-Repository-Pfad:
echo   %GIT_ROOT%
echo.

REM Verify this looks like the correct repository
if not exist "%GIT_ROOT%\visbo-connect" (
    echo FEHLER: Git-Repository-Struktur nicht gefunden!
    echo Erwartet: visbo-connect Ordner in %GIT_ROOT%
    echo.
    echo Bitte stellen Sie sicher, dass dieses Skript im Ordner
    echo   ^<git-ordner^>\visbo-connect\docs\admin-doc
    echo liegt.
    pause
    goto ende_error
)

REM ============ Ask for SPE Publish Folder ============
echo ===============================================
echo   SPE Publish-Ordner angeben
echo ===============================================
echo.
echo Bitte geben Sie den Pfad zum VISBO SPE Publish-Ordner an.
echo (Der Ordner, in den Sie SPE in Visual Studio veroeffentlicht haben)
echo.
set "DEFAULT_PUBLISH=C:\publish\VISBO SPE"
set /p "PUBLISH_DIR=Publish-Ordner [%DEFAULT_PUBLISH%]: "

if "%PUBLISH_DIR%"=="" set "PUBLISH_DIR=%DEFAULT_PUBLISH%"

REM Verify publish folder exists
if not exist "%PUBLISH_DIR%" (
    echo.
    echo FEHLER: Der Ordner "%PUBLISH_DIR%" existiert nicht!
    echo Bitte pruefen Sie den Pfad und fuehren Sie das Skript erneut aus.
    pause
    goto ende_error
)

REM Verify setup.exe exists in publish folder
if not exist "%PUBLISH_DIR%\setup.exe" (
    echo.
    echo FEHLER: "setup.exe" nicht gefunden in:
    echo   %PUBLISH_DIR%
    echo.
    echo Bitte pruefen Sie, ob der SPE-Build korrekt veroeffentlicht wurde.
    echo Die Datei setup.exe sollte im Publish-Ordner vorhanden sein.
    pause
    goto ende_error
)

echo.
echo Verwende Publish-Ordner: %PUBLISH_DIR%
echo.

REM ============ Set Repository Paths ============
set "CONNECT_DIR=%GIT_ROOT%\visbo-connect\docs\admin-doc\VISBO Connect"
set "CONNECT_SETUP_DIR=%CONNECT_DIR%\visbo-connect Setup 0.8.2"

REM Verify Connect setup exists
if not exist "%CONNECT_SETUP_DIR%\visbo-connect Setup.exe" (
    echo FEHLER: VISBO Connect Setup nicht gefunden in:
    echo   %CONNECT_SETUP_DIR%
    pause
    goto ende_error
)

REM ============ Install VISBO SPE ============
echo ===============================================
echo   Installiere VISBO Project Edit (SPE)
echo ===============================================
echo.

echo Starte VISBO SPE Setup...
echo Bitte folgen Sie den Anweisungen des Installationsprogramms.
echo.

pushd "%PUBLISH_DIR%"
start /wait "" "setup.exe"
popd

if errorlevel 1 (
    echo WARNUNG: SPE-Installation wurde moeglicherweise abgebrochen.
)

echo.
echo VISBO SPE Installation abgeschlossen.
echo.

REM ============ Install VISBO Connect ============
echo ===============================================
echo   Installiere VISBO Connect
echo ===============================================
echo.

REM Get Excel path from registry
echo Ermittle Excel-Installationspfad...
set "locEXCEL="
for /f "tokens=2*" %%a in ('reg query "HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\excel.exe" /ve 2^>nul') do set "locEXCEL=%%b"

if "%locEXCEL%"=="" (
    echo WARNUNG: Excel-Pfad konnte nicht automatisch ermittelt werden.
    echo Bitte geben Sie den vollstaendigen Pfad zu EXCEL.EXE an:
    set /p "locEXCEL=Excel-Pfad: "
)

echo Excel gefunden: %locEXCEL%

REM Prepare paths for JSON (escape backslashes)
set "SPE_TARGET=%AppData%\VISBO\VISBO Project Edit"
set "speSHEET=%SPE_TARGET%\VISBO Project Edit.xlsx"
set "speSHEET_JSON=%speSHEET:\=\\%"
set "locEXCEL_JSON=%locEXCEL:\=\\%"

REM Create config directory and JSON file
set "CONNECT_CONFIG_DIR=%AppData%\VISBO\visbo-connect"
if not exist "%CONNECT_CONFIG_DIR%" mkdir "%CONNECT_CONFIG_DIR%"

echo Erstelle Konfigurationsdatei...
echo {"excelExe":"%locEXCEL_JSON%","speSheet":"%speSHEET_JSON%"} > "%CONNECT_CONFIG_DIR%\vcn_config.json"

REM Run VISBO Connect Setup
echo Starte visbo-connect Setup...
echo.

pushd "%CONNECT_SETUP_DIR%"
start /wait "" "visbo-connect Setup.exe"
popd

echo.
echo VISBO Connect erfolgreich installiert!
echo.

REM ============ Configure Excel Trust Center ============
echo ===============================================
echo   Excel Trust Center Konfiguration
echo ===============================================
echo.

REM Detect Office version
set "officeVersion="
for %%v in (16.0 15.0 14.0 12.0) do (
    reg query "HKEY_CURRENT_USER\Software\Microsoft\Office\%%v\Excel\Security\Trusted Locations" >nul 2>&1 && set "officeVersion=%%v"
)

if "%officeVersion%"=="" (
    echo WARNUNG: Es wurde keine unterstuetzte Version von Microsoft Excel gefunden!
    echo Bitte konfigurieren Sie den vertrauenswuerdigen Speicherort manuell.
    echo.
    echo Pfad zum Hinzufuegen: %AppData%\VISBO
    echo.
    pause
    goto ende_success
)

echo Gefundene Office-Version: %officeVersion%
echo.

REM Set trusted location path
set "trustedPath=%APPDATA%\VISBO"

REM Ask user for trust center configuration method
echo Wie moechten Sie den vertrauenswuerdigen Speicherort hinzufuegen?
echo.
echo [1] Automatisch (Registry-Eintrag)
echo [2] Manuell (Anleitung fuer Excel Trust Center)
echo [3] Ueberspringen
echo.
choice /C 123 /N /M "Bitte waehlen Sie [1,2,3]: "

if errorlevel 3 goto ende_success
if errorlevel 2 goto manual_trust

REM ============ Automatic Trust Center Configuration ============
:auto_trust

REM Setup backup
set "backupDir=%USERPROFILE%\Documents\VISBO_Backups"
set "timestamp=%date:~6,4%%date:~3,2%%date:~0,2%_%time:~0,2%%time:~3,2%%time:~6,2%"
set "timestamp=!timestamp: =0!"
set "backupFile=!backupDir!\TrustedLocation_!timestamp!.reg"

if not exist "!backupDir!" mkdir "!backupDir!"

REM Create backup
echo Erstelle Backup der aktuellen Registry-Einstellungen...
reg export "HKEY_CURRENT_USER\Software\Microsoft\Office\%officeVersion%\Excel\Security\Trusted Locations" "!backupFile!" >nul 2>&1

REM Check for existing trusted location
set "pathExists=false"
for /f "tokens=*" %%A in ('reg query "HKEY_CURRENT_USER\Software\Microsoft\Office\%officeVersion%\Excel\Security\Trusted Locations" /s 2^>nul ^| findstr /i /c:"!trustedPath!"') do (
    set "pathExists=true"
)

if "!pathExists!"=="true" (
    echo Der Pfad "!trustedPath!" ist bereits als vertrauenswuerdiger Speicherort eingetragen.
    goto ende_success
)

REM Find next available location number
set "locationNum=0"
for /L %%i in (1,1,100) do (
    reg query "HKEY_CURRENT_USER\Software\Microsoft\Office\!officeVersion!\Excel\Security\Trusted Locations\Location%%i" >nul 2>&1 || (
        set "locationNum=%%i"
        goto :found_location
    )
)
:found_location

if !locationNum!==0 (
    echo FEHLER: Kein freier Location-Slot gefunden
    goto ende_success
)

REM Add trusted location
echo Hinzufuegen des vertrauenswuerdigen Speicherorts...
reg add "HKEY_CURRENT_USER\Software\Microsoft\Office\%officeVersion%\Excel\Security\Trusted Locations\Location!locationNum!" /v Path /t REG_SZ /d "!trustedPath!" /f >nul
reg add "HKEY_CURRENT_USER\Software\Microsoft\Office\%officeVersion%\Excel\Security\Trusted Locations\Location!locationNum!" /v Description /t REG_SZ /d "VISBO Trusted Location" /f >nul
reg add "HKEY_CURRENT_USER\Software\Microsoft\Office\%officeVersion%\Excel\Security\Trusted Locations\Location!locationNum!" /v AllowSubfolders /t REG_DWORD /d 1 /f >nul
reg add "HKEY_CURRENT_USER\Software\Microsoft\Office\%officeVersion%\Excel\Security\Trusted Locations\Location!locationNum!" /v Date /t REG_SZ /d "%date%" /f >nul

echo.
echo Vertrauenswuerdiger Speicherort wurde erfolgreich hinzugefuegt!
echo Backup gespeichert unter: !backupFile!
goto ende_success

REM ============ Manual Trust Center Instructions ============
:manual_trust
echo.
echo ===============================================
echo   Anleitung zum manuellen Hinzufuegen
echo ===============================================
echo.
echo Bitte folgen Sie diesen Schritten in Excel:
echo.
echo 1. Oeffnen Sie Excel
echo 2. Klicken Sie auf 'Datei' -^> 'Optionen'
echo 3. Waehlen Sie 'Sicherheitscenter' -^> 'Einstellungen fuer das Sicherheitscenter'
echo 4. Klicken Sie auf 'Vertrauenswuerdige Speicherorte' -^> 'Neuer Speicherort'
echo 5. Fuegen Sie diesen Pfad hinzu:
echo.
echo    %trustedPath%
echo.
echo 6. Aktivieren Sie 'Unterordner dieses Speicherorts sind ebenfalls vertrauenswuerdig'
echo 7. Klicken Sie auf 'OK'
echo.
goto ende_success

REM ============ Error Handlers ============
:ende_error
endlocal
exit /b 1

REM ============ Success ============
:ende_success
echo.
echo ===============================================
echo   Installation abgeschlossen!
echo ===============================================
echo.
echo Installierte Komponenten:
echo   - VISBO SPE: via Setup aus %PUBLISH_DIR%
echo   - VISBO Connect: Konfiguration in %CONNECT_CONFIG_DIR%
echo.
pause
endlocal
exit /b 0
