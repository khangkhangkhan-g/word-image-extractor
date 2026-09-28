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

if exist "publish\win-x64" rmdir /s /q "publish\win-x64"

echo Publishing self-contained Windows x64 build...
dotnet publish WordImageExtractor.csproj ^
  -c Release ^
  -r win-x64 ^
  --self-contained true ^
  -p:PublishSingleFile=true ^
  -p:IncludeNativeLibrariesForSelfExtract=true ^
  -p:DebugType=None ^
  -p:DebugSymbols=false ^
  -o "publish\win-x64"

if errorlevel 1 (
  echo.
  echo [ERROR] Publish failed. Review the messages above.
  pause
  exit /b 1
)

echo.
echo Portable build created in:
echo %~dp0publish\win-x64
explorer "%~dp0publish\win-x64"
endlocal
