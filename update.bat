@echo off
:: BAT file to clone T-invest repository
:: ------------------------------------

set REPO_URL=https://github.com/Scal1ger/T-invest.git
set FOLDER_NAME=T-invest

echo Checking for Git...
where git >nul 2>&1
if %errorlevel% neq 0 (
    echo Error: Git is not installed or not in PATH
    echo Download Git from https://git-scm.com/downloads
    pause
    exit /b 1
)

echo Checking if %FOLDER_NAME% folder exists...
if exist "%FOLDER_NAME%" (
    echo Error: Folder "%FOLDER_NAME%" already exists
    echo Delete it or choose another location
    pause
    exit /b 1
)

echo Cloning repository %REPO_URL%...
git clone %REPO_URL% %FOLDER_NAME%

if %errorlevel% equ 0 (
    echo Success! Repository cloned to "%FOLDER_NAME%" folder
) else (
    echo Error during cloning
    echo Possible reasons:
    echo 1. No internet connection
    echo 2. Repository access issues
    echo 3. Authentication problems
)

pause