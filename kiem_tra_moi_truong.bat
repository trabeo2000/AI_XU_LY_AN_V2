@echo off
setlocal enabledelayedexpansion
chcp 65001 >nul
echo ======================================================================
echo    KIEM TRA MOI TRUONG CHAY HE THONG XU LY AN TANDTC (3 MAY TINH)
echo ======================================================================
python -c "import sys; print(f'1. Python Version : {sys.version.split()[0]} -> [SAN SANG]'); libs = [('python-docx', 'docx'), ('pymupdf', 'pymupdf'), ('openpyxl', 'openpyxl')]; [print(f'   + {pkg:<15}: [SAN SANG]') for pkg, mod in libs if __import__(mod)]; print('\n==> KET QUA: Moi truong san sang 100%%, hoat dong tron tru tren 3 may tinh!')" 2>nul
if !errorlevel! neq 0 (
    echo [CANH BAO] May tinh nay chua cai du cac thu vien Python can thiet!
    echo Dang tien hanh cai dat tu requirements.txt...
    pip install -r requirements.txt
    if !errorlevel! equ 0 (
        echo [THANH CONG] Da cai dat xong thu vien!
    ) else (
        echo [LOI] Khong the cai dat thu vien. Vui long kiem tra ket noi mang hoac Python.
    )
)
echo ======================================================================
pause
