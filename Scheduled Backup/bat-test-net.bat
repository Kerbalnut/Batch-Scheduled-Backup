@echo off
REM Set VARS! 
REM Drive to backup (e.g. D:):
set "DRIVETOBAK=D:"
REM Backup Image (*.VHD) file name and full location (e.g. F:\shop_drivebox_d_data.VHD):
set "BAKUPIMG=F:\shop_drivebox_d_data.VHD"
REM Log files path (logs will be auto-named e.g. C:\Users\Shop\Documents\BackupLogs):
set "BAKUPLOGPATH=C:\Users\Shop\Documents\+Reference\Scripts\Robocopy\Scheduled Backup\Logs"
REM Backup drive (location where image is stored) volume label:
set "BAKUPLABEL=LocalDataBackup"
REM Backup drive (location where image is stored) volume serial:
set "BAKUPSERIAL=E693-8EB5"
REM e.g. Run "Vol F:" (F: is whereever your backup *.VHD is located)
REM Volume in drive F is LocalDataBackup
REM Volume Serial Number is E693-8EB5
:Start
REM Getting drive letter of backup image's location...
FOR /F %%G IN ("%BAKUPIMG%") DO (SET BAKUPLOC=%%~dG)
REM Getting just file name of backup image... (for text prompts)
FOR /F %%G IN ("%BAKUPIMG%") DO (SET BAKUPFILE=%%~nxG)
REM Get just letter of drive to be backed-up (no colon)
set DRIVETOBBLETTER=%DRIVETOBAK:~0,1%
REM Get time to show user when script started (in case it starts taking a long time)
FOR %%G IN (%Date%) DO SET Today=%%G
SET Now=%Time%
FOR /F "tokens=1-3 delims=/-" %%G IN ("%Today%") DO (
    SET DayMonth=%%G
    SET MonthDay=%%H
    SET Year=%%I
) 
REM Get volume label and check that drive to be backed up is online
REM ---------------------------------------------------------------
REM Known Issue: if the drive letter being checked is a network drive that's disconnected, it will show up as AVAILABLE
set "VolErrorHandling="
VOL %DRIVETOBAK% >"%Temp%\data-volume-info.txt" 2>&1
REM Redirect all output *and* errors to one file
IF ERRORLEVEL 1 set VolErrorHandling=NECESSARY
IF NOT [%VolErrorHandling%]==[NECESSARY] echo. >%Temp%\data-volume-info.txt
REM If VolErrorHandling is not NECESSARY (there were no errors when querying it), then print a blank line to the captured info.txt so it handles correctly
REM - Collect Vol Data:
FOR /F "tokens=*" %%G IN (%Temp%\data-volume-info.txt) DO (set VolLineOne=%%G&GOTO DataUnoEsc)
:DataUnoEsc
FOR /F "tokens=*" %%G IN (%Temp%\data-volume-info.txt) DO (set VolLineTwo=%%G)
FOR /F "tokens=6*" %%G IN ("%VolLineOne%") DO (set VolLabel=%%G)
FOR /F "tokens=5" %%G IN ("%VolLineTwo%") DO (set VolSerialNum=%%G)
IF "%VolLineOne%"=="The system cannot find the path specified." set "VolLabel=No drive connected."&set "VolSerialNum=No drive connected."
del %Temp%\data-volume-info.txt
REM Volume in drive D is Data
REM Volume Serial Number is 167A-F857
set "DRIVETOBAKLABEL=%VolLabel%"
IF "%VolLineOne%"=="The system cannot find the path specified." (set MainDataDrive=OFFLINE) ELSE (set MainDataDrive=ONLINE)

echo VOL line one = %VolLineOne%
echo VOL line two = %VolLineTwo%
echo Drive to backup label = %DRIVETOBAKLABEL%
echo(
pause

REM first text user sees:
echo ===============================================================================
echo === "%DRIVETOBAK%\ Data" Full Drive (Incremental) Backup (Data-D-backup-script.bat)   ===
echo ===============================================================================
echo(
echo %DRIVETOBAK%\%DRIVETOBAKLABEL% Backup - script started on %Today% at %Now%
echo(

:: ------- Start Script ------- 
:Step1
REM Step 1: Get UAC Admin Rights
REM Note, this will not work if run from a network share.

REM From: https://sites.google.com/site/eneerge/home/BatchGotAdmin
:: BatchGotAdmin
:-------------------------------------
REM  --> Check for permissions
>nul 2>&1 "%SYSTEMROOT%\system32\cacls.exe" "%SYSTEMROOT%\system32\config\system"

REM --> If error flag set, we do not have admin.
if '%errorlevel%' NEQ '0' (
    echo Requesting administrative privileges...
    goto UACPrompt
) else ( goto gotAdmin )

:UACPrompt
    echo Set UAC = CreateObject^("Shell.Application"^) > "%temp%\getadmin.vbs"
    echo UAC.ShellExecute "%~s0", "", "", "runas", 1 >> "%temp%\getadmin.vbs"

    "%temp%\getadmin.vbs"
    exit /B

:gotAdmin
    if exist "%temp%\getadmin.vbs" ( del "%temp%\getadmin.vbs" )
    pushd "%CD%"
    CD /D "%~dp0"
:--------------------------------------
echo %DRIVETOBAK%\%DRIVETOBAKLABEL% Backup - Step 1/4: Admin permissions obtained!
echo(

REM If drive (to be backed-up) was detected as offline or not attached in the beginning:
IF [%MainDataDrive%]==[ONLINE] (GOTO DriveToBBakdOnline) ELSE (GOTO NoDriveToBBakd)
:NoDriveToBBakd
REM Drive to be backed up is not online
echo %DRIVETOBAK%\%DRIVETOBAKLABEL% Backup - Step 2/4: ERROR = %DRIVETOBAK%\ is not online
echo(
echo SOLUTION: Make sure %DRIVETOBAK%\ is connected, then restart script.
goto Fail
:DriveToBBakdOnline
REM Drive to be backed up is online, so continue...

:Step2
REM Step 2: Mount VHD backup
REM check first that the *.vhd file exists in the right location

REM IF EXIST "%BAKUPIMG%" (GOTO FoundVHD) ELSE (GOTO VHDnotfound)
:VHDnotfound
REM VHD not in F:\, so does F:\ exist? is the drive even connected?

REM Known Issue: if the drive letter being checked is a network drive that's disconnected, it will show up as AVAILABLE
set "DriveErrorHandling="
cd /d %BAKUPLOC% 2>%Temp%\data-d-backup-drive-check-temp.txt
IF ERRORLEVEL 1 set DriveErrorHandling=NECESSARY
IF NOT [%DriveErrorHandling%]==[NECESSARY] echo. >%Temp%\data-d-backup-drive-check-temp.txt
REM If DriveErrorHandling is not NECESSARY (there were no errors when switching to it), then print a blank line to the error-report temp.txt so it handles correctly
FOR /F "delims=" %%G IN (%Temp%\data-d-backup-drive-check-temp.txt) DO (set DriveEr="%%G")
del %Temp%\data-d-backup-drive-check-temp.txt
IF %DriveEr%=="" set DriveInUse=TRUE
IF %DriveEr%=="The device is not ready." set DriveInUse=TRUE
IF %DriveEr%=="The system cannot find the drive specified." (set DriveInUse=FALSE) ELSE (set DriveInUse=TRUE)

REM IF [%DriveInUse%]==[TRUE] (GOTO ButFdriveIsOnline) ELSE (GOTO FdriveIsDisconnected)

:ButFdriveIsOnline
REM ? Okay, F drive is online. But no VHD. Check Volume label?
set "VolErrorHandling="
VOL %BAKUPLOC% >%Temp%\data-d-backup-volume-info.txt 2<&1
REM Redirect all output *and* errors to one file
IF ERRORLEVEL 1 set VolErrorHandling=NECESSARY
IF NOT [%VolErrorHandling%]==[NECESSARY] echo. >%Temp%\data-d-backup-volume-info.txt
REM If VolErrorHandling is not NECESSARY (there were no errors when querying it), then print a blank line to the captured info.txt so it handles correctly
REM - Collect Vol Data:
FOR /F "tokens=*" %%G IN (%Temp%\data-d-backup-volume-info.txt) DO (set VolLineOne=%%G&Goto UnoEsc)
:UnoEsc
FOR /F "tokens=*" %%G IN (%Temp%\data-d-backup-volume-info.txt) DO (set VolLineTwo=%%G)
FOR /F "tokens=6,7" %%G IN ("%VolLineOne%") DO (set VolLabel=%%G %%H)
FOR /F "tokens=5" %%G IN ("%VolLineTwo%") DO (set VolSerialNum=%%G)
IF "%VolLineOne%"=="The system cannot find the path specified." set "VolLabel=No drive connected."&set "VolSerialNum=No drive connected."
del %Temp%\data-d-backup-volume-info.txt
REM Volume in drive F is LocalDataBackup
REM Volume Serial Number is E693-8EB5

REM Test VOL command:
echo Vol Line Uno: %VolLineOne%
echo Vol Line Two: %VolLineTwo%
echo(
pause


IF "%VolLabel%"=="%BAKUPLABEL%" (GOTO LabelPass) ELSE (GOTO WrongFlabel)
:LabelPass
REM Label is "LocalDataBackup", now check serials
IF "%VolSerialNum%"=="%BAKUPSERIAL%" (GOTO SerialPass) ELSE (GOTO DifferentF)
:WrongFlabel
REM Label is NOT "LocalDataBackup", but let's check serial anyways.
IF "%VolSerialNum%"=="%BAKUPSERIAL%" (GOTO NoLabelYesSerial) ELSE (GOTO DifferentDrive)
:NoLabelYesSerial
REM Label does not match, but serial *IS* "E693-8EB5" Wtf is going on here?
echo %DRIVETOBAK%\%DRIVETOBAKLABEL% Backup - Step 2/4: ERROR = Volume %BAKUPLOC% Label was changed.
echo(
echo %BAKUPLOC%\ may have had a serious error. (or someone changed the drive label)
echo(
echo Serial numbers match, "%BAKUPSERIAL%"="%VolSerialNum%", but labels do not:
echo "%BAKUPLABEL%"=^!"%VolLabel%"
echo(
echo SOLUTION: Please check the drive. If it is o.k., please change the drive label.
goto Fail
:DifferentF
REM Label matches "LocalDataBackup" but serials do not match
echo %DRIVETOBAK%\%DRIVETOBAKLABEL% Backup - Step 2/4: ERROR = Volume %BAKUPLOC% serial does not match.
echo(
echo %BAKUPLOC%\ "%BAKUPLABEL%" serial number is supposed to be "%BAKUPSERIAL%"
echo Instead %BAKUPLOC%\ "%VolLabel%" serial number is: "%VolSerialNum%"
echo(
echo Another drive may be using the same label: "%BAKUPLABEL%"
echo(
echo SOLUTION: Shutdown, disconnect offending drive, connect %BAKUPLOC%\ "%BAKUPLABEL%",
echo boot up, and run script again.
goto Fail
:DifferentDrive
REM F:\ drive online, but Label and serial do not match. Another drive is using F:\
echo %DRIVETOBAK%\%DRIVETOBAKLABEL% Backup - Step 2/4: ERROR = Drive letter %BAKUPLOC% is in use by another drive.
echo(
echo The backup drive "%BAKUPLABEL%" requires %BAKUPLOC%\ to work.
echo(
echo SOLUTION: Shutdown, disconnect drive using letter %BAKUPLOC%\, 
echo connect %BAKUPLOC%\ "%BAKUPLABEL%", boot up, and run script again.
goto Fail
:SerialPass
REM Drive label and serial match! But no VHD?! Uh-oh.
echo - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
echo %DRIVETOBAK%\%DRIVETOBAKLABEL% Backup - Step 2/4: ERROR = Backup image could not be located in %BAKUPLOC%\
echo(
echo Drive %BAKUPLOC%\"%BAKUPLABEL%" is supposed to contain "%BAKUPIMG%"
echo(
echo This error is not trivial.
echo( 
echo Creating a new backup image from scratch is long, intensive process. It could 
echo cause the drive to fail.
echo(
echo SOLUTION: Search for "%BAKUPFILE%" in %BAKUPLOC%\. Search everywhere to 
echo make sure the image did not get lost or moved. If found, move to the proper 
echo location and run this script again.
echo - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
echo If you've already searched for the backup and this is the *second time* you are
echo seeing this message, stay on the line...
pause
goto FreshBackupImage
:FdriveIsDisconnected
echo %DRIVETOBAK%\%DRIVETOBAKLABEL% Backup - Step 2/4: ERROR = Could not find %BAKUPLOC%\
echo(
echo Drive %BAKUPLOC%\ "%BAKUPLABEL%" appears to be disconnected.
echo(
echo SOLUTION: Shutdown, reconnect %BAKUPLOC%\, boot up, and run script again.
goto Fail
:FoundVHD
echo Found backup image!

REM Ok, so we have the VHD. Now where to mount it? Check what drive letters available.
GOTO SkipLetterCheck
REM ---------------------------FIND DRIVE LETTER---------------------------

REM Look for available drive letter to use

REM As of this writing (6/16/2014), the ShopPC (where this script is supposed to go) has drive letters C, D, F, E, J, K, L, M, N, Y, Z in use and drive letter R is prefered, but not required by this script. That leaves GHI, OPQ, and RSTUVWX available, (Q could possibly be used for QuickBooks in the future) which are what we're going to check.
REM i.e.:
REM ABCDEFGHIJKLMNOPQRSTUVWXYZ
REM       GHI     OP RSTUVWX  

REM "R" is prefered so we will check that one first.
REM R - GHI OP STUVWX

REM Known Issue: if the drive letter being checked is a network drive that's disconnected, it will show up as AVAILABLE

set "AvailableDriveLetters="
set "DriveErrorHandling="
cd /d R:\ 2>%Temp%\scan-tbi-drive-check-temp.txt
IF ERRORLEVEL 1 set DriveErrorHandling=NECESSARY
IF NOT [%DriveErrorHandling%]==[NECESSARY] echo. >%Temp%\scan-tbi-drive-check-temp.txt
REM If DriveErrorHandling is not NECESSARY (there were no errors when switching to it), then print a blank line to the error-report temp.txt so it handles correctly
FOR /F "delims=" %%G IN (%Temp%\scan-tbi-drive-check-temp.txt) DO (set DriveEr="%%G")
IF %DriveEr%=="" set DriveInUse=TRUE
IF %DriveEr%=="The device is not ready." set DriveInUse=TRUE
IF %DriveEr%=="The system cannot find the drive specified." (set DriveInUse=FALSE) ELSE (set DriveInUse=TRUE)
REM IF [%DriveInUse%]==[TRUE] (echo DETECTING: Drive R:\ is in use) ELSE (echo DETECTING: Drive R:\ is AVAILABLE!)
IF [%DriveInUse%]==[FALSE] set AvailableDriveLetters=R,
del %Temp%\scan-tbi-drive-check-temp.txt

set "DriveErrorHandling="
cd /d G:\ 2>%Temp%\scan-tbi-drive-check-temp.txt
IF ERRORLEVEL 1 set DriveErrorHandling=NECESSARY
IF NOT [%DriveErrorHandling%]==[NECESSARY] echo. >%Temp%\scan-tbi-drive-check-temp.txt
REM If DriveErrorHandling is not NECESSARY (there were no errors when switching to it), then print a blank line to the error-report temp.txt so it handles correctly
FOR /F "delims=" %%G IN (%Temp%\scan-tbi-drive-check-temp.txt) DO (set DriveEr="%%G")
IF %DriveEr%=="" set DriveInUse=TRUE
IF %DriveEr%=="The device is not ready." set DriveInUse=TRUE
IF %DriveEr%=="The system cannot find the drive specified." (set DriveInUse=FALSE) ELSE (set DriveInUse=TRUE)
REM IF [%DriveInUse%]==[TRUE] (echo DETECTING: Drive G:\ is in use) ELSE (echo DETECTING: Drive G:\ is AVAILABLE!)
IF [%DriveInUse%]==[FALSE] set AvailableDriveLetters=G,
del %Temp%\scan-tbi-drive-check-temp.txt

set "DriveErrorHandling="
cd /d H:\ 2>%Temp%\scan-tbi-drive-check-temp.txt
IF ERRORLEVEL 1 set DriveErrorHandling=NECESSARY
IF NOT [%DriveErrorHandling%]==[NECESSARY] echo. >%Temp%\scan-tbi-drive-check-temp.txt
REM If DriveErrorHandling is not NECESSARY (there were no errors when switching to it), then print a blank line to the error-report temp.txt so it handles correctly
FOR /F "delims=" %%G IN (%Temp%\scan-tbi-drive-check-temp.txt) DO (set DriveEr="%%G")
IF %DriveEr%=="" set DriveInUse=TRUE
IF %DriveEr%=="The device is not ready." set DriveInUse=TRUE
IF %DriveEr%=="The system cannot find the drive specified." (set DriveInUse=FALSE) ELSE (set DriveInUse=TRUE)
REM IF [%DriveInUse%]==[TRUE] (echo DETECTING: Drive H:\ is in use) ELSE (echo DETECTING: Drive H:\ is AVAILABLE!)
IF [%DriveInUse%]==[FALSE] set AvailableDriveLetters=%AvailableDriveLetters%H,
del %Temp%\scan-tbi-drive-check-temp.txt

set "DriveErrorHandling="
cd /d I:\ 2>%Temp%\scan-tbi-drive-check-temp.txt
IF ERRORLEVEL 1 set DriveErrorHandling=NECESSARY
IF NOT [%DriveErrorHandling%]==[NECESSARY] echo. >%Temp%\scan-tbi-drive-check-temp.txt
REM If DriveErrorHandling is not NECESSARY (there were no errors when switching to it), then print a blank line to the error-report temp.txt so it handles correctly
FOR /F "delims=" %%G IN (%Temp%\scan-tbi-drive-check-temp.txt) DO (set DriveEr="%%G")
IF %DriveEr%=="" set DriveInUse=TRUE
IF %DriveEr%=="The device is not ready." set DriveInUse=TRUE
IF %DriveEr%=="The system cannot find the drive specified." (set DriveInUse=FALSE) ELSE (set DriveInUse=TRUE)
REM IF [%DriveInUse%]==[TRUE] (echo DETECTING: Drive I:\ is in use) ELSE (echo DETECTING: Drive I:\ is AVAILABLE!)
IF [%DriveInUse%]==[FALSE] set AvailableDriveLetters=%AvailableDriveLetters%I,
del %Temp%\scan-tbi-drive-check-temp.txt

set "DriveErrorHandling="
cd /d O:\ 2>%Temp%\scan-tbi-drive-check-temp.txt
IF ERRORLEVEL 1 set DriveErrorHandling=NECESSARY
IF NOT [%DriveErrorHandling%]==[NECESSARY] echo. >%Temp%\scan-tbi-drive-check-temp.txt
REM If DriveErrorHandling is not NECESSARY (there were no errors when switching to it), then print a blank line to the error-report temp.txt so it handles correctly
FOR /F "delims=" %%G IN (%Temp%\scan-tbi-drive-check-temp.txt) DO (set DriveEr="%%G")
IF %DriveEr%=="" set DriveInUse=TRUE
IF %DriveEr%=="The device is not ready." set DriveInUse=TRUE
IF %DriveEr%=="The system cannot find the drive specified." (set DriveInUse=FALSE) ELSE (set DriveInUse=TRUE)
REM IF [%DriveInUse%]==[TRUE] (echo DETECTING: Drive O:\ is in use) ELSE (echo DETECTING: Drive O:\ is AVAILABLE!)
IF [%DriveInUse%]==[FALSE] set AvailableDriveLetters=%AvailableDriveLetters%O,
del %Temp%\scan-tbi-drive-check-temp.txt

set "DriveErrorHandling="
cd /d P:\ 2>%Temp%\scan-tbi-drive-check-temp.txt
IF ERRORLEVEL 1 set DriveErrorHandling=NECESSARY
IF NOT [%DriveErrorHandling%]==[NECESSARY] echo. >%Temp%\scan-tbi-drive-check-temp.txt
REM If DriveErrorHandling is not NECESSARY (there were no errors when switching to it), then print a blank line to the error-report temp.txt so it handles correctly
FOR /F "delims=" %%G IN (%Temp%\scan-tbi-drive-check-temp.txt) DO (set DriveEr="%%G")
IF %DriveEr%=="" set DriveInUse=TRUE
IF %DriveEr%=="The device is not ready." set DriveInUse=TRUE
IF %DriveEr%=="The system cannot find the drive specified." (set DriveInUse=FALSE) ELSE (set DriveInUse=TRUE)
REM IF [%DriveInUse%]==[TRUE] (echo DETECTING: Drive P:\ is in use) ELSE (echo DETECTING: Drive P:\ is AVAILABLE!)
IF [%DriveInUse%]==[FALSE] set AvailableDriveLetters=%AvailableDriveLetters%P,
del %Temp%\scan-tbi-drive-check-temp.txt

set "DriveErrorHandling="
cd /d S:\ 2>%Temp%\scan-tbi-drive-check-temp.txt
IF ERRORLEVEL 1 set DriveErrorHandling=NECESSARY
IF NOT [%DriveErrorHandling%]==[NECESSARY] echo. >%Temp%\scan-tbi-drive-check-temp.txt
REM If DriveErrorHandling is not NECESSARY (there were no errors when switching to it), then print a blank line to the error-report temp.txt so it handles correctly
FOR /F "delims=" %%G IN (%Temp%\scan-tbi-drive-check-temp.txt) DO (set DriveEr="%%G")
IF %DriveEr%=="" set DriveInUse=TRUE
IF %DriveEr%=="The device is not ready." set DriveInUse=TRUE
IF %DriveEr%=="The system cannot find the drive specified." (set DriveInUse=FALSE) ELSE (set DriveInUse=TRUE)
REM IF [%DriveInUse%]==[TRUE] (echo DETECTING: Drive S:\ is in use) ELSE (echo DETECTING: Drive S:\ is AVAILABLE!)
IF [%DriveInUse%]==[FALSE] set AvailableDriveLetters=%AvailableDriveLetters%S,
del %Temp%\scan-tbi-drive-check-temp.txt

set "DriveErrorHandling="
cd /d T:\ 2>%Temp%\scan-tbi-drive-check-temp.txt
IF ERRORLEVEL 1 set DriveErrorHandling=NECESSARY
IF NOT [%DriveErrorHandling%]==[NECESSARY] echo. >%Temp%\scan-tbi-drive-check-temp.txt
REM If DriveErrorHandling is not NECESSARY (there were no errors when switching to it), then print a blank line to the error-report temp.txt so it handles correctly
FOR /F "delims=" %%G IN (%Temp%\scan-tbi-drive-check-temp.txt) DO (set DriveEr="%%G")
IF %DriveEr%=="" set DriveInUse=TRUE
IF %DriveEr%=="The device is not ready." set DriveInUse=TRUE
IF %DriveEr%=="The system cannot find the drive specified." (set DriveInUse=FALSE) ELSE (set DriveInUse=TRUE)
REM IF [%DriveInUse%]==[TRUE] (echo DETECTING: Drive T:\ is in use) ELSE (echo DETECTING: Drive T:\ is AVAILABLE!)
IF [%DriveInUse%]==[FALSE] set AvailableDriveLetters=%AvailableDriveLetters%T,
del %Temp%\scan-tbi-drive-check-temp.txt

set "DriveErrorHandling="
cd /d U:\ 2>%Temp%\scan-tbi-drive-check-temp.txt
IF ERRORLEVEL 1 set DriveErrorHandling=NECESSARY
IF NOT [%DriveErrorHandling%]==[NECESSARY] echo. >%Temp%\scan-tbi-drive-check-temp.txt
REM If DriveErrorHandling is not NECESSARY (there were no errors when switching to it), then print a blank line to the error-report temp.txt so it handles correctly
FOR /F "delims=" %%G IN (%Temp%\scan-tbi-drive-check-temp.txt) DO (set DriveEr="%%G")
IF %DriveEr%=="" set DriveInUse=TRUE
IF %DriveEr%=="The device is not ready." set DriveInUse=TRUE
IF %DriveEr%=="The system cannot find the drive specified." (set DriveInUse=FALSE) ELSE (set DriveInUse=TRUE)
REM IF [%DriveInUse%]==[TRUE] (echo DETECTING: Drive U:\ is in use) ELSE (echo DETECTING: Drive U:\ is AVAILABLE!)
IF [%DriveInUse%]==[FALSE] set AvailableDriveLetters=%AvailableDriveLetters%U,
del %Temp%\scan-tbi-drive-check-temp.txt

set "DriveErrorHandling="
cd /d V:\ 2>%Temp%\scan-tbi-drive-check-temp.txt
IF ERRORLEVEL 1 set DriveErrorHandling=NECESSARY
IF NOT [%DriveErrorHandling%]==[NECESSARY] echo. >%Temp%\scan-tbi-drive-check-temp.txt
REM If DriveErrorHandling is not NECESSARY (there were no errors when switching to it), then print a blank line to the error-report temp.txt so it handles correctly
FOR /F "delims=" %%G IN (%Temp%\scan-tbi-drive-check-temp.txt) DO (set DriveEr="%%G")
IF %DriveEr%=="" set DriveInUse=TRUE
IF %DriveEr%=="The device is not ready." set DriveInUse=TRUE
IF %DriveEr%=="The system cannot find the drive specified." (set DriveInUse=FALSE) ELSE (set DriveInUse=TRUE)
REM IF [%DriveInUse%]==[TRUE] (echo DETECTING: Drive V:\ is in use) ELSE (echo DETECTING: Drive V:\ is AVAILABLE!)
IF [%DriveInUse%]==[FALSE] set AvailableDriveLetters=%AvailableDriveLetters%V,
del %Temp%\scan-tbi-drive-check-temp.txt

set "DriveErrorHandling="
cd /d W:\ 2>%Temp%\scan-tbi-drive-check-temp.txt
IF ERRORLEVEL 1 set DriveErrorHandling=NECESSARY
IF NOT [%DriveErrorHandling%]==[NECESSARY] echo. >%Temp%\scan-tbi-drive-check-temp.txt
REM If DriveErrorHandling is not NECESSARY (there were no errors when switching to it), then print a blank line to the error-report temp.txt so it handles correctly
FOR /F "delims=" %%G IN (%Temp%\scan-tbi-drive-check-temp.txt) DO (set DriveEr="%%G")
IF %DriveEr%=="" set DriveInUse=TRUE
IF %DriveEr%=="The device is not ready." set DriveInUse=TRUE
IF %DriveEr%=="The system cannot find the drive specified." (set DriveInUse=FALSE) ELSE (set DriveInUse=TRUE)
REM IF [%DriveInUse%]==[TRUE] (echo DETECTING: Drive W:\ is in use) ELSE (echo DETECTING: Drive W:\ is AVAILABLE!)
IF [%DriveInUse%]==[FALSE] set AvailableDriveLetters=%AvailableDriveLetters%W,
del %Temp%\scan-tbi-drive-check-temp.txt

set "DriveErrorHandling="
cd /d X:\ 2>%Temp%\scan-tbi-drive-check-temp.txt
IF ERRORLEVEL 1 set DriveErrorHandling=NECESSARY
IF NOT [%DriveErrorHandling%]==[NECESSARY] echo. >%Temp%\scan-tbi-drive-check-temp.txt
REM If DriveErrorHandling is not NECESSARY (there were no errors when switching to it), then print a blank line to the error-report temp.txt so it handles correctly
FOR /F "delims=" %%G IN (%Temp%\scan-tbi-drive-check-temp.txt) DO (set DriveEr="%%G")
IF %DriveEr%=="" set DriveInUse=TRUE
IF %DriveEr%=="The device is not ready." set DriveInUse=TRUE
IF %DriveEr%=="The system cannot find the drive specified." (set DriveInUse=FALSE) ELSE (set DriveInUse=TRUE)
REM IF [%DriveInUse%]==[TRUE] (echo DETECTING: Drive X:\ is in use) ELSE (echo DETECTING: Drive X:\ is AVAILABLE!)
IF [%DriveInUse%]==[FALSE] set AvailableDriveLetters=%AvailableDriveLetters%X,
del %Temp%\scan-tbi-drive-check-temp.txt

REM Finished checking all possible drives!

REM echo(
REM echo Available Drive Letters =%AvailableDriveLetters%
REM echo(

REM LOOP : Count to end-of-string to find number of AvailableDriveLetters:

set /A "CheckLetterPosition=1"
:StartCheckAgain
FOR /F "tokens=%CheckLetterPosition% delims=," %%G IN ("%AvailableDriveLetters%") DO (set CurrentLetter=%%G)
IF [%CurrentLetter%]==[%LastLetter%] set /A "NumberOfAvailDriveLett=CheckLetterPosition-1"
IF [%CurrentLetter%]==[%LastLetter%] GOTO FoundEnd
set LastLetter=%CurrentLetter%
set /A "CheckLetterPosition+=1"
GOTO StartCheckAgain
:FoundEnd

REM Excellent! Now, our available drive letters should be stored in %AvailableDriveLetters%, while the number of drive letters in that variable should be stored in %NumberOfAvailDriveLett%.

:StartOverCheckLetter
:FreshStartLetterCheck
set /A "DriveLetterToCheck=1"
:CheckDriveLetter
FOR /F "tokens=%DriveLetterToCheck% delims=," %%G IN ("%AvailableDriveLetters%") DO (set PossibleDriveLetter=%%G)
GOTO ApprovedDriveLetter
GOTO CheckNextDriveLetter
GOTO Fail
:CheckNextDriveLetter
IF %DriveLetterToCheck% EQU %NumberOfAvailDriveLett% GOTO StartOverCheckLetter
set /A "DriveLetterToCheck+=1"
GOTO CheckDriveLetter
:ApprovedDriveLetter
set UserAcceptedDriveLetter=%PossibleDriveLetter%
REM echo(
REM echo Drive letter "%UserAcceptedDriveLetter%:\" selected!
REM echo(
echo(
echo Located unused drive letter (%PossibleDriveLetter%:)

REM Sweetness. OK, let's review: we now have a list of available drive letters to use (%AvailableDriveLetters%), one of which is confirmed good by the user (%UserAcceptedDriveLetter%), and a path to our hard drive image file (%PathToTBI%).

REM --------------------(/END)-FIND DRIVE LETTER-(/END)--------------------
:SkipLetterCheck
set UserAcceptedDriveLetter=R

REM sel vdisk file="F:\shop_drivebox_d_data.VHD"
REM attach vdisk
REM select partition 1
REM assign letter=R

echo sel vdisk file="%BAKUPIMG%">%Temp%\mount-backup.txt
echo attach vdisk>>%Temp%\mount-backup.txt
REM When attaching a VHD using DISKPART like this, it seems to want to pick the same letter it used last time (in this case R:)
diskpart /s %Temp%\mount-backup.txt>%Temp%\mount-status.log
del "%Temp%\mount-backup.txt"

REM Check that VHD was mounted o.k. (All were 'successful')

echo(
echo %DRIVETOBAK%\%DRIVETOBAKLABEL% Backup - Step 2/4: VHD mounted!

set "SHHHHTESTING="
set "SHHHHTESTING=^/l "
:Step3
REM Step 3: Robocopy all changes from D:\ Data to VHD backup

REM echo This backup is a test...
REM robocopy D:\ R:\ *.* /e /copyall /b /xo /XF Backup-PC-HomeComp3-18-11.TBI pagefile.sys /XD "D:\System Volume Information"  /mir /l /r:10 /eta /tee /log:"C:\Users\Shop\Documents\+Reference\Scripts\Robocopy\Scheduled Backup\Logs\Data-D-backup-testlog.txt"

@echo ON
robocopy %DRIVETOBAK%\ %UserAcceptedDriveLetter%:\ *.* /e /copyall /b /xo /XF Backup-PC-HomeComp3-18-11.TBI pagefile.sys /XD "D:\System Volume Information" /mir /r:10 /eta /tee %SHHHHTESTING%/log+:"%BAKUPLOGPATH%\%Year%-%DayMonth%-%MonthDay%-%DRIVETOBBLETTER%-%DRIVETOBAKLABEL%-backup.log"
@echo OFF

REM robocopy "<source>" "<destination>"
REM /S - Copies subdirectories. Note that this option excludes empty directories.
REM /E - Copies subdirectories. Note that this option includes empty directories. For additional information, see Remarks.
REM /copy:<CopyFlags>
REM Specifies the file properties to be copied. The following are the valid values for this option:
REM D - Data
REM A - Attributes
REM T - Time stamps
REM S - NTFS access control list (ACL)
REM O - Owner information
REM U - Auditing information
REM The default value for CopyFlags is DAT (data, attributes, and time stamps).
REM /copyall - Copies all file information (equivalent to /copy:DATSOU).
REM /MIR - Mirrors a directory tree (equivalent to /e plus /purge), deletes any destination files. Note that when used with /Z (or MAYBE even /XO) it does not delete already copied files at the destination (useful for resuming a copy)
REM /L - Specifies that files are to be listed only (and not copied, deleted, or time stamped).
REM /ETA - Shows the estimated time of arrival (ETA) of the copied files.
REM /log:<LogFile> - Writes the status output to the log file (overwrites the existing log file).
REM /log+:<LogFile> - Writes the status output to the log file (appends the output to the existing log file).
REM /TEE - Output to console window, as well as the log file.
REM /IS - Include Same, overwrite files even if they are already the same.
REM /IT - Include Tweaked files.
REM /X - Report all eXtra files, not just those selected & copied.
REM /FFT - uses fat file timing instead of NTFS. This means the granularity is a bit less precise. For across-network share operations this seems to be much more reliable - just don't rely on the file timings to be completely precise to the second.
REM /Z - ensures Robocopy can resume the transfer of a large file in mid-file instead of restarting. (Restart Mode)(maybe for Network Copys)
REM /B - copies in Backup Modes (overrides ACLs for files it doesn't have access to so it can copy them. Requires User-Level or Admin permissions)
REM /XO - Excludes older files. (Only copies newer and changed files)
REM /XF <FileName>[ ...] Excludes files that match the specified names or paths. Note that FileName can include wildcard characters (* and ?).
REM /XD <Directory>[ ...] Excludes directories that match the specified names and paths.
REM XF and XD can be used in combination  e.g. ROBOCOPY c:\source d:\dest /XF *.doc *.xls /XD c:\unwanted /S
REM e.g. ROBOCOPY C:\source D:\dest /XD "C:\System Volume Information"

echo(
echo %DRIVETOBAK%\%DRIVETOBAKLABEL% Backup - Step 3/4: Robocopy'd all changes from %DRIVETOBAK%\ %DRIVETOBAKLABEL% to VHD backup!

:Step4
REM Step 4: Unmount VHD

REM diskpart sel vdisk file="F:\shop_drivebox_d_data.VHD"
REM diskpart detach vdisk

echo sel vdisk file="%BAKUPIMG%">%Temp%\unmount-backup.txt
echo detach vdisk>>%Temp%\unmount-backup.txt
diskpart /s %Temp%\unmount-backup.txt>%Temp%\unmount-status.log
del "%Temp%\unmount-backup.txt"

echo(
echo %DRIVETOBAK%\%DRIVETOBAKLABEL% Backup - Step 4/4: VHD Unmounted!
GOTO End

:FreshBackupImage
cls
echo ===============================================================================
echo =                             New Backup Image                                =
echo - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
echo(
echo If you cannot find the backup image for drive %DRIVETOBAK%\ "%DRIVETOBAKLABEL%", it will have to be 
echo created again from scratch. ("%BAKUPIMG%")
echo(
echo This process should be treated with *caution*.
echo(
echo Creating a drive image requires a full read of all ^>1 TB of data from %DRIVETOBAK%\.
echo(
echo This operation, if done on an already unstable ^& old drive, could itself cause
echo the drive to fail.
echo(
echo Only start this process if you can be sure this machine will not be shutdown or
echo lose power (for a whole day to be safe). Start it late evening ^& run overnight.
echo(
echo ===============================================================================
echo(
echo If you are confident the backup image is lost, and no one will disturb this new
echo imaging process until it is done, proceed...
echo(
CHOICE /m "PROCEED?"
IF ERRORLEVEL 2 exit
IF ERRORLEVEL 1 GOTO FreshBackupImageStart
GOTO Fail
:FreshBackupImageStart
GOTO Fail
exit

:End
echo(
echo %DRIVETOBAK%\%DRIVETOBAKLABEL% Backup - script completed!
pause
exit

:Fail
echo(
echo The script could not continue.
pause
exit