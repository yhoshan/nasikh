@echo off
REM ============================================================
REM   بناء ملف ناسخ.exe  (ملف تنفيذي واحد لويندوز)
REM ============================================================
REM   تشغّل هذا الملف مرة واحدة على جهاز ويندوز فيه بايثون.
REM   سينتج ملف:  dist\ناسخ.exe
REM   هذا الملف هو ما توزّعه على الزملاء - ينقرون عليه فيعمل
REM   مباشرة بلا تثبيت أي شيء.
REM ============================================================

chcp 65001 >nul
echo.
echo ====== (1/3) تثبيت المكتبات المطلوبة ======
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
python -m pip install pyinstaller==6.11.1

echo.
echo ====== (2/3) بناء الملف التنفيذي ======
pyinstaller --noconfirm --onefile --windowed ^
  --name "ناسخ" ^
  --collect-all docx ^
  --hidden-import tkinter ^
  app.py

echo.
echo ====== (3/3) تم ======
echo الملف التنفيذي جاهز في المجلد:   dist\ناسخ.exe
echo وزّعه على الزملاء - ينقرون عليه نقرا مزدوجا فيعمل مباشرة.
echo.
pause
