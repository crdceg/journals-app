@echo off

echo =========================
echo Backup databases and output...
echo =========================

IF EXIST dist\app\databases (
    xcopy dist\app\databases databases /E /I /Y
)

IF EXIST dist\app\output (
    xcopy dist\app\output output /E /I /Y
)

echo =========================
echo Cleaning old build...
echo =========================

IF EXIST build rmdir /s /q build
IF EXIST dist rmdir /s /q dist
IF EXIST app.spec del app.spec

echo =========================
echo Building EXE...
echo =========================

pyinstaller --onedir --noconsole --add-data "templates;templates" app.py

echo =========================
echo Copying runtime data...
echo =========================

IF EXIST databases (
    xcopy databases dist\app\databases /E /I /Y
)

IF EXIST output (
    xcopy output dist\app\output /E /I /Y
)

copy reviewers_master.xlsx dist\app\

echo =========================
echo DONE SUCCESSFULLY ✅
echo =========================

pause