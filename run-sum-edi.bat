@echo off
chcp 65001 >nul
cd /d "%~dp0"
echo ===== SUM-EDI =====
echo วางไฟล์ MF และ CNTRS (.xls / .xlsx) ไว้ในโฟลเดอร์ input ก่อน
echo.

rem หา Python ที่รันได้จริง (ตัว python.exe ของ Microsoft Store บางเครื่องเป็นแค่ทางลัดที่รันไม่ได้)
set PY=
py -3 -c "import sys" >nul 2>nul && set PY=py -3
if not defined PY python -c "import sys" >nul 2>nul && set PY=python
if not defined PY if exist "%LOCALAPPDATA%\Python\bin\python.exe" set PY="%LOCALAPPDATA%\Python\bin\python.exe"
if not defined PY (
  echo !! ไม่พบ Python - ติดตั้งจาก https://www.python.org/downloads/ แล้วติ๊ก "Add python.exe to PATH"
  pause
  exit /b 1
)

%PY% -c "import pymupdf, xlrd" >nul 2>nul
if errorlevel 1 (
  echo กำลังติดตั้งไลบรารีที่ต้องใช้ ครั้งแรกครั้งเดียว...
  %PY% -m pip install -r requirements.txt
)

if not exist output mkdir output
%PY% -X utf8 scripts\sum_mf.py .
echo.
echo ผลลัพธ์อยู่ที่ output\SUM-EDI.pdf และ output\SUMMARY.pdf  (ถ้าเปิดไฟล์ค้างไว้ ให้ปิดก่อนรัน)
pause
