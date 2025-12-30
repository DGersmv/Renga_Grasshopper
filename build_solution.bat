@echo off
echo Building solution Renga_Grasshopper.sln...
"C:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe" Renga_Grasshopper.sln /p:Configuration=Release /p:Platform=x64 /t:Build /nologo
if %ERRORLEVEL% EQU 0 (
    echo.
    echo Build completed successfully!
) else (
    echo Build failed
    exit /b 1
)





