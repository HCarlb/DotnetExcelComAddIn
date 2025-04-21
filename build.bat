@echo off
:: Build script for the Excel Add-In solution
cls
setlocal

set SOLUTION=TheExcelAddInSolution.sln
set TARGETS=COMContract;HcExcelAddIn;Register
set CONFIG=Debug
set FRAMEWORK=net8.0-windows

for %%P in (x64 x86) do (

    echo --------------------------------------------------------------------------------
    echo Building %%P...
    msbuild %SOLUTION% /t:%TARGETS% /p:Configuration=%CONFIG%;Platform=%%P /v:minimal
    if errorlevel 1 exit /b 1

    
    echo Copying Register.* for %%P...
::    xcopy /Y "Registration\bin\%%P\%CONFIG%\%FRAMEWORK%\Register.*" "ExcelAddIn\bin\%%P\%CONFIG%\%FRAMEWORK%\" >nul
    xcopy /Y "Registration\bin\%%P\%CONFIG%\%FRAMEWORK%\Register.*" "ExcelAddIn\bin\%%P\%CONFIG%\%FRAMEWORK%\"
)

echo --------------------------------------------------------------------------------
endlocal
