@echo off
echo ========================================
echo VBA MCP Server Installer Build Script
echo ========================================
echo.

echo Step 1: Publishing VbaMcpServer (CLI)...
dotnet publish ..\src\VbaMcpServer -c Release -r win-x64 --self-contained /p:PublishSingleFile=true
if %ERRORLEVEL% NEQ 0 (
    echo ERROR: Failed to publish VbaMcpServer
    exit /b 1
)
echo.

echo Step 2: Publishing VbaMcpServer.GUI...
dotnet publish ..\src\VbaMcpServer.GUI -c Release -r win-x64 --self-contained /p:PublishSingleFile=true
if %ERRORLEVEL% NEQ 0 (
    echo ERROR: Failed to publish VbaMcpServer.GUI
    exit /b 1
)
echo.

echo Step 3: Building MSI installer with WiX...
wix build Product.wxs -o bin\VbaMcpServer.msi -arch x64
if %ERRORLEVEL% NEQ 0 (
    echo ERROR: Failed to build MSI installer
    exit /b 1
)
echo.

echo ========================================
echo Build completed successfully!
echo Output: installer\bin\VbaMcpServer.msi
echo ========================================
