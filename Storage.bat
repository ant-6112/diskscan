@echo off
setlocal enabledelayedexpansion

REM Set default values
set "path=."
set "minimum_size=10.0"
set "unit=MB"

REM Function to convert size to bytes
:ConvertSizeToBytes
set "size=%1"
set "unit=%2"
if /i "%unit%"=="KB" (set /a "size*=1024")
if /i "%unit%"=="GB" (set /a "size*=1024*1024")
exit /b

REM Function to find large files
:FindLargeFiles
set "total_files=0"
for /f %%A in ('dir /s /b /a-d "%path%" 2^>nul ^| find /c /v ""') do set "total_files=%%A"

set "multiplier=1"
if /i "%unit%"=="KB" set "multiplier=1024"
if /i "%unit%"=="GB" set "multiplier=1048576"

set "size_threshold_bytes=%minimum_size%"
call :ConvertSizeToBytes %minimum_size% %unit%

set "index=0"
for /r "%path%" %%F in (*) do (
    set /a "index+=1"
    set "file_path=%%F"
    for %%I in (!file_path!) do set "file_path=%%~dpnI%%~xI"
    for %%F in (!file_path!) do set "file_size=%%~zF"

    REM Check if file size exceeds threshold
    if !file_size! gtr %size_threshold_bytes% (
        set "Files_Found[!index!]=!file_path!"
        set "File_Sizes[!index!]=!file_size!"
        set /a "User_ID[!index!]=%%~aF"
    )
    REM Display progress
    set /a "progress=(index * 100) / total_files"
    echo Searching for large files... !progress!%% complete
)

REM Display found files
for /l %%i in (1,1,%index%) do (
    set "file_path=!Files_Found[%%i]!"
    set "file_size=!File_Sizes[%%i]!"
    set "user_id=!User_ID[%%i]!"
    for /f "tokens=1" %%U in ('net user %user_id% /domain ^| findstr /i "Full Name"') do set "user_name=%%U"
    echo !file_path!: !file_size! bytes (Created by: !user_name!)
)

REM Sort users by storage usage
set "user_storage="
for /l %%i in (1,1,%index%) do (
    set "user_id=!User_ID[%%i]!"
    for /f "tokens=1 delims= " %%U in ('net user %user_id% /domain ^| findstr /i "Full Name"') do set "user_name=%%U"
    set /a "user_storage[!user_name!]+=!File_Sizes[%%i]!"
)

REM Export top users data to Excel
set "temp_file=%temp%\top_users_data.csv"
echo User,Storage (%unit%)> "%temp_file%"
for /f "tokens=2,3 delims== " %%U in ('set user_storage[') do (
    set /a "storage=%%V / multiplier"
    echo %%U,!storage!>> "%temp_file%"
)

REM Open Excel workbook
start excel "%temp_file%"

echo.
echo Top Users Data is Exported
exit /b
