@echo off
echo Building MCP Teams Connector...

REM Clean previous build
if exist dist rmdir /s /q dist

REM Compile TypeScript
echo Compiling TypeScript...
call npm run build

if %errorlevel% neq 0 (
    echo Build failed!
    exit /b 1
)

echo.
echo Build complete! Files are in the dist/ directory.
echo.
echo To run the server:
echo   npm start
echo.
echo To install dependencies first:
echo   npm install