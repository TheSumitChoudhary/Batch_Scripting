@ECHO OFF
SETLOCAL EnableDelayedExpansion
rem -------------------------------------------------------------
rem  LDR_CUBIC_WRAPPER.BAT                         
rem  PURPOSE: Process GL_JOURNAL_ENTRY_TRX2 files with correct backup naming
rem -------------------------------------------------------------

SET LOG_FILE=E:\HULFT\CUBIC\PROG\Proc_Log\LDR_CUBIC_WRAPPER_%DATE:~10,4%%DATE:~4,2%%DATE:~7,2%.log
SET RECV_DIR=E:\HULFT\CUBIC\RECV\
SET PROG_DIR=E:\HULFT\CUBIC\PROG\
SET BACKUP_DIR=E:\HULFT\CUBIC\BACKUP\
SET TEMP_DIR=E:\HULFT\CUBIC\TEMP\

rem --- Set source and target parameters ---
SET SOURCE=GL_JOURNAL_ENTRY_TRX2
SET TARGET=GL_JOURNAL_ENTRY_TRX

ECHO [%DATE% %TIME%] Started LDR_CUBIC_WRAPPER >> %LOG_FILE%
ECHO [%DATE% %TIME%] Parameter: %1 >> %LOG_FILE%

rem --- Only process GL_JOURNAL_ENTRY_TRX2 ---
IF NOT "%1"=="%SOURCE%" (
    ECHO [%DATE% %TIME%] ERROR: This wrapper only handles %SOURCE% >> %LOG_FILE%
    ENDLOCAL
    EXIT /B 1
)

rem --- Ensure temp directory exists ---
IF NOT EXIST "%TEMP_DIR%NUL" MD "%TEMP_DIR%"

rem --- Check for source files ---
SET FILE_FOUND=0

rem --- Handle ZIP file if present ---
IF EXIST "%RECV_DIR%T_%SOURCE%.zip" (
    ECHO [%DATE% %TIME%] Found %RECV_DIR%T_%SOURCE%.zip >> %LOG_FILE%
    SET FILE_FOUND=1
    
    COPY "%RECV_DIR%T_%SOURCE%.zip" "%RECV_DIR%T_%TARGET%.zip" /Y >> %LOG_FILE% 2>&1
    IF ERRORLEVEL 1 (
        ECHO [%DATE% %TIME%] ERROR: Failed to copy ZIP file >> %LOG_FILE%
        ENDLOCAL
        EXIT /B 1
    )
)

rem --- Handle CSV file if present ---
IF EXIST "%RECV_DIR%T_%SOURCE%.csv" (
    ECHO [%DATE% %TIME%] Found %RECV_DIR%T_%SOURCE%.csv >> %LOG_FILE%
    SET FILE_FOUND=1
    
    COPY "%RECV_DIR%T_%SOURCE%.csv" "%RECV_DIR%T_%TARGET%.csv" /Y >> %LOG_FILE% 2>&1
    IF ERRORLEVEL 1 (
        ECHO [%DATE% %TIME%] ERROR: Failed to copy CSV file >> %LOG_FILE%
        ENDLOCAL
        EXIT /B 1
    )
)

IF %FILE_FOUND%==0 (
    ECHO [%DATE% %TIME%] ERROR: No source files found for %SOURCE% >> %LOG_FILE%
    ENDLOCAL
    EXIT /B 1
)

rem --- Ensure source backup directory exists ---
IF NOT EXIST "%BACKUP_DIR%%SOURCE%\" (
    MKDIR "%BACKUP_DIR%%SOURCE%\"
    ECHO [%DATE% %TIME%] Created directory: %BACKUP_DIR%%SOURCE%\ >> %LOG_FILE%
)

rem --- Capture list of files in target backup directory before processing ---
SET "BEFORE_FILE=%TEMP_DIR%before_files.txt"
DIR "%BACKUP_DIR%%TARGET%\*.*" /B /A:-D > "%BEFORE_FILE%" 2>NUL

rem --- Run original script with target parameter ---
ECHO [%DATE% %TIME%] Executing LDR_CUBIC.BAT %TARGET% >> %LOG_FILE%
CALL "%PROG_DIR%LDR_CUBIC.BAT" %TARGET%
SET RESULT=%ERRORLEVEL%
ECHO [%DATE% %TIME%] LDR_CUBIC.BAT completed with exit code: %RESULT% >> %LOG_FILE%

rem --- Capture list of files in target backup directory after processing ---
SET "AFTER_FILE=%TEMP_DIR%after_files.txt"
DIR "%BACKUP_DIR%%TARGET%\*.*" /B /A:-D > "%AFTER_FILE%" 2>NUL

rem --- Find new files by comparing before and after lists ---
ECHO [%DATE% %TIME%] Identifying new backup files... >> %LOG_FILE%
FOR /F "usebackq delims=" %%F IN (`FINDSTR /V /G:"%BEFORE_FILE%" "%AFTER_FILE%"`) DO (
    ECHO [%DATE% %TIME%] Found new backup file: %%F >> %LOG_FILE%
    
    rem --- Determine correct destination filename by replacing TARGET with SOURCE ---
    SET "DEST_FILE=%%F"
    SET "DEST_FILE=!DEST_FILE:%TARGET%=%SOURCE%!"
    
    ECHO [%DATE% %TIME%] Moving "%BACKUP_DIR%%TARGET%\%%F" to "%BACKUP_DIR%%SOURCE%\!DEST_FILE!" >> %LOG_FILE%
    MOVE "%BACKUP_DIR%%TARGET%\%%F" "%BACKUP_DIR%%SOURCE%\!DEST_FILE!" >> %LOG_FILE% 2>&1
)

rem --- Clean up temporary files ---
DEL /F /Q "%BEFORE_FILE%" "%AFTER_FILE%" >> %LOG_FILE% 2>&1
IF EXIST "%RECV_DIR%T_%TARGET%.zip" DEL /F /Q "%RECV_DIR%T_%TARGET%.zip" >> %LOG_FILE% 2>&1
IF EXIST "%RECV_DIR%T_%TARGET%.csv" DEL /F /Q "%RECV_DIR%T_%TARGET%.csv" >> %LOG_FILE% 2>&1

ECHO [%DATE% %TIME%] Completed with exit code: %RESULT% >> %LOG_FILE%

ENDLOCAL
EXIT /B %RESULT%
