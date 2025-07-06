@echo off
setlocal enabledelayedexpansion
echo Organizing release files...

REM Check if output path parameter is provided
if "%~1"=="" (
    echo No output path provided, using default...
    cd /d "%~dp0"
    dotnet build MsgToPdfConverter.sln -c Release
    cd bin\Release\net48
) else (
    echo Using provided output path: %~1
    cd /d "%~1"
)

REM Copy wkhtmltopdf library from Debug if missing in Release
if not exist "libwkhtmltox.dll" (
    set "debugPath=%~dp0bin\Debug\net48\libwkhtmltox.dll"
    if exist "!debugPath!" (
        echo Copying libwkhtmltox.dll from Debug folder...
        copy "!debugPath!" . >nul 2>&1
    ) else (
        echo WARNING: libwkhtmltox.dll not found in Debug folder
        echo You may need to run Debug build first or manually copy wkhtmltopdf binaries
    )
)

REM Create libraries folder
if not exist libraries mkdir libraries

REM Move everything except MsgToPdfConverter.exe and .config to libraries folder
for %%f in (*) do (
    if /i not "%%f"=="MsgToPdfConverter.exe" (
        if /i not "%%~xf"==".config" (
            move "%%f" libraries\ 2>nul
        )
    )
)

REM Move all folders to libraries EXCEPT architecture folders needed by DinkToPdf
for /d %%d in (*) do (
    if /i not "%%d"=="libraries" (
        if /i not "%%d"=="x64" (
            if /i not "%%d"=="x86" (
                if /i not "%%d"=="arm64" (
                    if /i not "%%d"=="win-x64" (
                        if /i not "%%d"=="win-x86" (
                            if /i not "%%d"=="win-arm64" (
                                move "%%d" libraries\ 2>nul
                            )
                        )
                    )
                )
            )
        )    )
)


REM Keep the main exe and config in root
echo.
echo Release organized:
echo - MsgToPdfConverter.exe (main folder)
echo - MsgToPdfConverter.exe.config (main folder)  
echo - All other files and folders moved to libraries\ folder
echo - libwkhtmltox.dll moved to libraries\ folder
echo.
