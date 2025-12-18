@echo off
echo Building RengaGH...
"C:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe" RengaGH\RengaGH.csproj /p:Configuration=Release /p:Platform=x64 /t:Build /nologo
if %ERRORLEVEL% EQU 0 (
    echo RengaGH build succeeded
    echo.
    echo Building GrasshopperRNG...
    "C:\Program Files\Microsoft Visual Studio\2022\Community\MSBuild\Current\Bin\MSBuild.exe" GrasshopperRNG\GrasshopperRNG.csproj /p:Configuration=Release /p:Platform=AnyCPU /t:Build /nologo
    if %ERRORLEVEL% EQU 0 (
        echo GrasshopperRNG build succeeded
        echo.
        echo Build completed successfully!
    ) else (
        echo GrasshopperRNG build failed
        exit /b 1
    )
) else (
    echo RengaGH build failed
    exit /b 1
)

