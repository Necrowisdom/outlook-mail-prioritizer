@echo off
setlocal
cd /d "%~dp0"

echo ============================================
echo  Outlook Mail Onceliklendirme
echo ============================================
echo.

where python >nul 2>nul
if errorlevel 1 (
    echo HATA: Python bulunamadi.
    echo Lutfen once Python yukleyin: https://www.python.org/downloads/
    echo Kurulum sirasinda "Add Python to PATH" secenegini isaretleyin.
    echo.
    pause
    exit /b 1
)

python -c "import win32com" >nul 2>nul
if errorlevel 1 (
    echo pywin32 kutuphanesi kuruluyor, lutfen bekleyin...
    python -m pip install --quiet pywin32
    echo.
)

echo Outlook uzerinden okunmamis mailler taraniyor...
echo ^(Outlook'un acik ve hesabina giris yapilmis olmasi gerekir^)
echo.

python "%~dp0outlook_prioritize.py"

echo.
echo ============================================
echo  Islem tamamlandi.
echo ============================================
pause
endlocal
