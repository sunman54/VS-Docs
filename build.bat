@echo off
setlocal

echo ========================================================
echo   Universal Document Viewer - Production Build Script
echo ========================================================
echo.

:: 1. Node.js kontrolu
where node >nul 2>nul
if %ERRORLEVEL% neq 0 (
    echo [HATA] Node.js bulunamadi! Lutfen Node.js kurulu oldugundan emin olun.
    goto :error
)

:: 2. Eski VSIX dosyalarini temizle
echo [1/4] Eski build ve VSIX dosyalari temizleniyor...
if exist "*.vsix" (
    del /f /q *.vsix >nul 2>nul
)

:: 3. TypeScript Tip Denetimi
echo [2/4] TypeScript tip denetimi yapiliyor (typecheck)...
call npm run typecheck
if %ERRORLEVEL% neq 0 (
    echo.
    echo [HATA] TypeScript tip kontrolunde hatalar tespit edildi!
    goto :error
)

:: 4. Otomatik Test Paketi
echo [3/4] Otomatik test paketi calistiriliyor (npm test)...
call npm test
if %ERRORLEVEL% neq 0 (
    echo.
    echo [HATA] Otomatik testler basarisiz oldu!
    goto :error
)

:: 5. Production VSIX Paketleme
echo [4/4] Production VSIX paketi uretiliyor (vsce package)...
call npm run package
if %ERRORLEVEL% neq 0 (
    echo.
    echo [HATA] VSIX paketi olusturulurken bir hata meydana geldi!
    goto :error
)

echo.
echo ========================================================
echo   BUILD BASARIYLA TAMAMLANDI! (PRODUCTION READY)
echo ========================================================
echo.
echo Uretilen VSIX Paketi:
dir /b *.vsix
echo.
echo Eklentiyi VS Code'a Kurmak Icin:
echo   1. VS Code icinde Extensions sekmesini acin (Ctrl+Shift+X)
echo   2. Sag ustteki (...) menusunden "Install from VSIX..." secin
echo   3. Uretilen .vsix dosyasini isaretleyin.
echo.
goto :end

:error
echo.
echo ========================================================
echo   BUILD BASARISIZ OLDU! Lutfen yukaridaki loglari inceleyin.
echo ========================================================
echo.
exit /b 1

:end
endlocal
