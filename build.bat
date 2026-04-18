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

pyinstaller --onedir --noconsole app.py

echo =========================
echo Preparing runtime folders...
echo =========================

IF NOT EXIST dist\app mkdir dist\app

echo Copying templates...
IF EXIST templates (
    xcopy templates dist\app\templates /E /I /Y
) ELSE (
    echo ERROR: templates folder not found!
)

echo Copying databases...
IF EXIST databases (
    xcopy databases dist\app\databases /E /I /Y
) ELSE (
    echo WARNING: databases folder not found!
)

echo Copying output...
IF EXIST output (
    xcopy output dist\app\output /E /I /Y
) ELSE (
    mkdir dist\app\output
)

echo Copying reviewers file...
IF EXIST reviewers_master.xlsx (
    copy reviewers_master.xlsx dist\app\
) ELSE (
    echo ERROR: reviewers_master.xlsx NOT FOUND!
)

echo =========================
echo DONE SUCCESSFULLY ✅
echo =========================

pause