@echo off
setlocal
cd /d "%~dp0"

where dotnet >nul 2>nul
if errorlevel 1 (
  echo [ERROR] .NET 8 SDK was not found.
  echo Install the .NET 8 SDK from Microsoft, then run this file again.
  pause
  exit /b 1
)

echo Restoring and building Word Image Extractor...
dotnet build WordImageExtractor.csproj -c Release
if errorlevel 1 (
  echo.
  echo [ERROR] Build failed. Review the messages above.
  pause
  exit /b 1
)

echo.
echo Starting Word Image Extractor...
dotnet run --project WordImageExtractor.csproj -c Release --no-build
endlocal
