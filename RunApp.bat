@echo off
echo Building Bulk Mail Sender Application...
dotnet build MailApplication.csproj --verbosity quiet

if %ERRORLEVEL% EQU 0 (
    echo Build successful! Starting application...
    echo.
    start "" ".\bin\Debug\net8.0-windows\MailApplication.exe"
    echo Application started!
) else (
    echo Build failed! Please check the errors above.
    pause
)
