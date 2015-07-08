@echo off
set "LocalPath=%~dp0"
GOTO Setup
:===============================================================================
System Requirements:
-------------------------------------------------------------------------------
 - Windows 7 and up (untested on Vista)
 - PowerShell v?.
 
To check PowerShell version, Run: 
	$PSVersionTable.PSVersion
Also make sure PowerShell Execution Policy is not restricted:
	Get-ExecutionPolicy
	Set-ExecutionPolicy RemoteSigned

Sub-script Dependencies: (These must stay in the same folder as this script to work)
-------------------------------------------------------------------------------
 - Parse-RobocopyLogs.ps1
 - DateMath.cmd

A Proper Specification:
-------------------------------------------------------------------------------
Disclaimer: This script uses a copy method that can delete files at the destination (ROBOCOPY /MIR). It is designed to DETECT and ABORT itself before undertaking any potentially risky operation (e.g. abnormally large size of changes to copy).

Raison d'être: Backup changes from a SOURCE hard disk to a VHD file located on different hard disk. Designed to be a replacement backup solution that can be scheduled to run repeatedly and complete without user intervention. VHD is used as a compact solution that does not take over all the space on the DESTINATION hard disk. (for Thin Provisioned disks)
Uses built-in tools that work with Vista and later. DISKPART to mount and unmound VHD images, ROBOCOPY to copy only the differences (incremental backup, ROBOCOPY /XO), POWERSHELL to do some find-and-replace one-liners and run log parsing scripts.

USE-CASE (Scenario): A shared computer running Windows 7 PRO SP1 is shut down and rebooted many times throughout the day. New drives are attached and detached regularly, so backup DESTINATION or SOURCE may become disconnected or have their preferred drive letter occupied. 
Script needs to detect this and work around it to make sure both SOURCE and DESTINATION disks match the Unique Identifiers we can get for them (using VOL), and detect that they are mounted and online. Then scan for changes between the two disks and copy differences over using ROBOCOPY.

Non-goals: This version will *NOT* support the following features:
•	Looking for the *.VHD file anywhere besides the location it’s supposed to be at (-> FAIL. Make user look for VHD with a search)
•	Create new *.VHD file for backup if it goes missing (maybe in the future, but this is an intensive process requiring a lot of work to implement that can be abused by ignorant users if automated)
•	Warn user about size of changes for anything less than a gig. We assume here that the user will become concerned if the operation will copy more than just a few gigs. SizeCutoff will only be measured in GB. (Options to check for anything less will not be implemented)

Overview:
1.	Script must accurately detect that both SOURCE and DESTINATION are online, that it is looking at the *RIGHT* disks by using a Unique ID (VOL), and that the VHD exists, in right place, with the right SERIAL and LABEL. If any piece is missing, fail with ERROR.

2. 	After mounting VHD, if after the (DISKPART) mount process the drive gets a different letter than the one usually used, the script must detect this and adjust itself accordingly.
	- This is done by sorting thru *ALL* connected disks' Unique ID's and finding one that matches the indended drive (VOL). If it cannot find where the VHD was mounted to, fail with ERROR.

3.	Copy over ONLY THE DIFFERENCES since last backup (New and Changed files) (ROBOCOPY /XO)
	- Perform TEST COPY to build logs. (ROBOCOPY /L)
	- Analyze logs to get size of changes to be copied, time in days since last backup was performed, and if last backup was successful. (Parse-RobocopyLogs.ps1)
		i.	If the size of changes to be copied is > n GB, WARN
		ii.	If the last operation happened > n days ago, NOTIFY user. (DateMath.cmd)
		iii.If last backup was not successful or completed with errors, NOTIFY
	- If all safety checks are passed, proceed with backup operation automatically.
		
4.  Unmount VHD (DISKPART) and parse final copy logs and find if any files failed (Parse-RobocopyLogs.ps1) and NOTIFY user if there were any errors.

NOTIFY = The script will print important text on-screen, but still proceed automatically with backup.
WARNING = The script will pause itself and wait for USER INTERVENTION, this is to allow somebody check what it's about to do before it proceeds automatically with a potentially dangerous operation.
ERROR = Some essential part is not in place, leaving script no choice but to Fail and stop itself.


THINGS TO ADD LATER (POSSIBLY):
- Create new VHD if the listed one cannot be found (and do fresh ROBOCOPY)
- If Label is detected as changed (drive Letter and Serial still the same) give option to auto-rename it back.
- Make it so by default only test copies are run, require -notest switch for copy to happen.


INSTRUCTIONS:
-------------------------------------------------------------------------------
Make sure PowerShell Execution Policy is not restricted:
	Get-ExecutionPolicy
	Set-ExecutionPolicy RemoteSigned
	
TO CUSTOMIZE THIS SCRIPT: All variables you need peronalize this script are in the "SETUP" section below, between the :Setup and :Start tags.

To get volume label and serial, Run this at command prompt:
	VOL D:

Replace "D:" with whatever drive letter your backup image (*.VHD) is at.
Run command "VOL F:" to get LABEL and SERIAL of volumes (Where "F:" is the letter for the volume you need to get a LABEL and SERIAL for) 

e.g.:
Volume in drive D is Data
Volume Serial Number is 167A-F857
 - or - 
Volume in drive C has no label.
Volume Serial Number is E693-8EB5

If volume has no label, set "LABEL=no label."

- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
To get SERIAL and LABEL of your VHD, you will first have to manually mount it (if it's not mounted already) and run a VOL command on it

you can use these commands: 
	DISKPART
	select vdisk file="C:\TEST\demo.vhd"
	attach vdisk

It should attach successfully and automatically assign itself a letter. Use that letter to set the var VHDPREFERRED=.

After you are done if you still have the same DISKPART window open you can skip the first command, otherwise run:
	select vdisk file="C:\TEST\demo.vhd"
	detach vdisk

- - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
ADDITIONAL CUSTOMIZATION: To customize the ROBOCOPY switches used, (ctrl + F) FIND :ptThreeC
:===============================================================================
:: All Variables are required unless otherwise noted with (Optional).
:Setup
:: SOURCE DRIVE: Drive letter to be backed-up including colon (e.g. D:)
set "DRIVETOBAK=C:"
:: DESTINATION IMAGE FULL PATH: Backup Image (*.VHD) file name and full location (e.g. F:\shop_drivebox_d_data.VHD):
set "BAKUPIMG=D:\test-mount.vhd"
:: (Optional) VHD's preferred mount letter, including colon (:). When mounting VHD, it will try to use same letter every time. If you know what letter this is, it can save time by checking there first. If not found at the preferred location, all other drives will be scanned. Set to NULL to turn off ("VHDPREFERRED=")
set "VHDPREFERRED=F:"
:: Log directory path (e.g. C:\Users\Shop\Documents\Scheduled Backup\Logs) By default, it will use the Logs folder where this script is saved (set "BAKUPLOGPATH=%LocalPath%Logs"):
set "BAKUPLOGPATH=%LocalPath%Logs"
:: SOURCE drive volume LABEL (drive to be backed-up):
set "SOURCELABEL=no label."
:: SOURCE drive volume SERIAL (drive to be backed-up):
set "SOURCESERIAL=5CC8-F408"
:: DESTINATION Backup volume LABEL (drive where *.VHD image is stored):
set "BAKUPLABEL=Data"
:: DESTINATION Backup volume SERIAL (drive where *.VHD image is stored):
set "BAKUPSERIAL=E4F5-E264"
:: DESTINATION Virtual drive volume LABEL (mounted *.VHD drive label):
set "BAKUPVHDLABEL=Test-VHD"
:: DESTINATION Virtual drive volume SERIAL (mounted *.VHD drive serial number):
set "BAKUPVHDSERIAL=5AB6-FDEE"
:: (Optional) Integer representing maximum number of days in age since last backup was performed before warning user. (i.e. IF last backup age >n Days Ago, warn user before proceeding) Set to zero (0) or NULL ("DATECUTOFF=") to turn this check off
set "DATECUTOFF=21"
:: (Optional) Whole integer, representing size in GB, to allow changes to be copied before warning user. (i.e. IF size of change >n GB, warn user before proceeding) Set to zero (0) or NULL ("SIZECUTOFF=") to continue automatically no matter the size of change
set "SIZECUTOFF=5"
:Start
REM Abstract other variables:
REM PARSING: Getting drive letter of backup image's location...
FOR /F %%G IN ("%BAKUPIMG%") DO (SET BAKUPLOC=%%~dG)
REM PARSING: Getting just file name of backup image... (for text prompts)
FOR /F %%G IN ("%BAKUPIMG%") DO (SET BAKUPFILE=%%~nxG)
REM PARSING: Get just letter of SOURCE drive to be backed-up (no colon)
set DRIVETOBBLETTER=%DRIVETOBAK:~0,1%
REM Get time to show user when script started (in case it starts taking a long time)
FOR %%G IN (%Date%) DO SET StartToday=%%G
SET StartNow=%Time%
FOR /F "tokens=1-3 delims=/-" %%G IN ("%StartToday%") DO (
	:: Note: This is actually the Month in numbers
    SET DayMonth=%%G
	:: Note: This is actually the Day, in numbers
    SET MonthDay=%%H
    SET Year=%%I
) 
REM Double-check variables we have are good:
REM Having an incorrect path for the log folder will generate a huge amount of errors.
IF NOT EXIST "%BAKUPLOGPATH%" (
	echo ERROR= Log folder path does not exist. ^(%BAKUPLOGPATH%^)
	echo.
	echo Check the variable BAKUPLOGPATH is set to a valid path.
	GOTO Fail
)
REM The following is an outline of all the tags used below to break up steps:
REM -------------------------------------------------------------------------------
:ptZero
REM Get Admin rights.
:ptOneA
REM Check if SOURCE drive is online.
:ptOneB
REM Check if DESTINATION drive is online.
:ptOneC
REM Check if VHD exists.
:ptTwoA
REM Mount VHD with DISKPART. Check that mount was successful.
:ptTwoB
REM Check if VHD mounted at preferred letter. (check other letters if it's not)
:ptThreeA
REM Was last backup successful?
:ptThreeB
REM Check date of last backup.
:ptThreeC
REM Perform test copy. Find expected size of copy. Is size of changes above limit?
:ptThreeD
REM Start ROBOCOPY.
:ptThreeE
REM Unmount VHD with DISKPART.
:ptFour
REM Last step! Was backup successful? Parse logs to find any copy errors.
REM ===============================================================================

REM BEGINNNNNNNNNNNNNNNNNNNNNNNNN
REM -------------------------------------------------------------------------------

REM first text user sees:
echo ===============================================================================
echo ===      "%DRIVETOBAK%\" Full Drive (Incremental) Backup (Scheduled-Backup.bat)       ===
echo ===============================================================================
echo(
echo %DRIVETOBAK%\%SOURCELABEL% Backup - started on %StartToday% at %StartNow%
echo(

:ptZero
REM Get Admin rights.
REM Step 1: Get UAC Admin Rights
:: ------- Start Script ------- 
REM Note, this will not work if run from a network share.

REM From: https://sites.google.com/site/eneerge/home/BatchGotAdmin
:: BatchGotAdmin
:-------------------------------------------------------------------------------
REM  --> Check for permissions
>nul 2>&1 "%SYSTEMROOT%\system32\cacls.exe" "%SYSTEMROOT%\system32\config\system"

REM --> If error flag set, we do not have admin.
if '%errorlevel%' NEQ '0' (
	REM NOTIFY user:
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
:-------------------------------------------------------------------------------
echo %DRIVETOBAK%\%SOURCELABEL% Backup - Step 0/4: Admin permissions obtained!
echo(
REM End ptZero:

:ptOneA
REM Check if SOURCE drive is online.
REM Step A: Check that drive to be backed up (SOURCE) is online
REM -------------------------------------------------------------------------------
VOL %DRIVETOBAK% >%Temp%\data-volume-info.txt 2>&1
REM Redirect all output *and* errors to one file
REM - Collect Vol Data:
set "VolLabel="
set "VolLabel2="
set "VolSerialNum="
set "VolLineOne="
set "VolLineTwo="
FOR /F "tokens=*" %%G IN (%Temp%\data-volume-info.txt) DO (set VolLineOne=%%G&GOTO DataUnoEsc)
:DataUnoEsc
FOR /F "tokens=*" %%G IN (%Temp%\data-volume-info.txt) DO (set VolLineTwo=%%G)
FOR /F "tokens=6*" %%G IN ("%VolLineOne%") DO (set VolLabel=%%G&set VolLabel2=%%H)
FOR /F "tokens=5" %%G IN ("%VolLineTwo%") DO (set VolSerialNum=%%G)
IF "%VolLineOne%"=="The system cannot find the path specified." (
	set "VolLabel=No drive connected."
	set "VolSerialNum=No drive connected."
)
IF "%VolLabel%"=="no" (
	IF "%VolLabel2%"=="label." set "VolLabel=no label."
)
del %Temp%\data-volume-info.txt
REM Is drive online? Does SOURCE volume's label and serial match?
IF "%VolLabel%"=="No drive connected." (set MainDataDrive=OFFLINE) ELSE (set MainDataDrive=ONLINE)
IF "%VolLabel%"=="%SOURCELABEL%" (set CheckSourceLabel=PASS) ELSE (set CheckSourceLabel=FAIL)
IF "%VolSerialNum%"=="%SOURCESERIAL%" (set CheckSourceSerial=PASS) ELSE (set CheckSourceSerial=FAIL)
REM If drive (to be backed-up)(SOURCE) was detected as offline or not attached in the beginning:
IF [%MainDataDrive%]==[ONLINE] (GOTO DriveToBBakdOnline) ELSE (GOTO NoDriveToBBakd)
:NoDriveToBBakd
REM Drive to be backed up is not online
echo %DRIVETOBAK%\%SOURCELABEL% Backup - Step 1/4: ERROR = %DRIVETOBAK%\ is not online
echo(
echo SOLUTION: Shutdown, make sure %DRIVETOBAK%\ is connected, then restart script.
GOTO Fail
:DriveToBBakdOnline
REM Drive to be backed up (SOURCE) is online, now check if correct drive...
IF "%CheckSourceLabel%"=="FAIL" (GOTO SourceLabelSerialFAIL)
IF "%CheckSourceSerial%"=="FAIL" (GOTO SourceLabelSerialFAIL)
GOTO SourceLabelSerialPASS
:SourceLabelSerialFAIL
REM SOURCE Drive checks failed, either LABEL or SERIAL does not match
echo %DRIVETOBAK%\%SOURCELABEL% Backup - Step 1/4: ERROR = %DRIVETOBAK%\ is not correct drive
echo(
echo Label or Serial for volume "%DRIVETOBAK%" do not match.
echo(
echo Label for %DRIVETOBAK% is "%VolLabel%", was expecting "%SOURCELABEL%"
echo Serial for %DRIVETOBAK% is %VolSerialNum%, was expecting %SOURCESERIAL%
echo(
echo -------------------------------------------------------------------------------
echo SOLUTION: Make sure %DRIVETOBAK%\ is the intended drive, shutdown, connect %DRIVETOBAK%, boot up,
echo and restart script.
GOTO Fail
:SourceLabelSerialPASS
REM End ptOneA: SOURCE drive is online and correct!

:ptOneB
REM Check if DESTINATION drive is online.
REM -------------------------------------------------------------------------------
REM We got more precise error codes in the legacy script. Let's see if we can re-use...
REM First, we'll use CD /D to check if the drive we want is online
REM Then, we'll use VOL like we did before, to also check if drive is online but also get SERIAL and LABEL as well
REM Then, we'll check SERIAL and LABEL against expected values, and give a custom error code for each outcome
REM Let's go... (legacy code separated by \ and =)
REM \=\=\=\=\=\=\=\=\=\=\=\=\=\=\=\=\=\=\=\=\=\=\=\=\=\=\=\=\=\=\=\=\=\=\=\=\=\=\=\
REM Known Issue: if the drive letter being checked is a network drive that's disconnected, it will show up as AVAILABLE
set "DriveErrorHandling="
CD /D %BAKUPLOC% 2>%Temp%\data-d-backup-drive-check-temp.txt
IF ERRORLEVEL 1 set DriveErrorHandling=NECESSARY
IF NOT [%DriveErrorHandling%]==[NECESSARY] echo. >%Temp%\data-d-backup-drive-check-temp.txt
REM If DriveErrorHandling is not NECESSARY (there were no errors when switching to it), then print a blank line to the error-report temp.txt so it handles correctly
FOR /F "delims=" %%G IN (%Temp%\data-d-backup-drive-check-temp.txt) DO (set DriveEr="%%G")
del %Temp%\data-d-backup-drive-check-temp.txt
IF %DriveEr%=="" set DriveInUse=TRUE
IF %DriveEr%=="The device is not ready." set DriveInUse=TRUE
IF %DriveEr%=="The system cannot find the drive specified." (set DriveInUse=FALSE) ELSE (set DriveInUse=TRUE)
IF [%DriveInUse%]==[TRUE] (set "CDDestinationCheque=ONLINE") ELSE (set "CDDestinationCheque=OFFLINE")
REM \=\=\=\=\=\=\=\=\=\=\=\=\=\=\=\=\=\=\=\=\=\=\=\=\=\=\=\=\=\=\=\=\=\=\=\=\=\=\=\
VOL %BAKUPLOC% >%Temp%\DEST-volume-info.txt 2>&1
REM Redirect all output *and* errors to one file
REM - Collect Vol Data:
set "VolLabel="
set "VolLabel2="
set "VolSerialNum="
set "VolLineOne="
set "VolLineTwo="
FOR /F "tokens=*" %%G IN (%Temp%\DEST-volume-info.txt) DO (set VolLineOne=%%G&GOTO DataDESTHOSTEsc)
:DataDESTHOSTEsc
FOR /F "tokens=*" %%G IN (%Temp%\DEST-volume-info.txt) DO (set VolLineTwo=%%G)
FOR /F "tokens=6*" %%G IN ("%VolLineOne%") DO (set VolLabel=%%G&set VolLabel2=%%H)
FOR /F "tokens=5" %%G IN ("%VolLineTwo%") DO (set VolSerialNum=%%G)
IF "%VolLineOne%"=="The system cannot find the path specified." (
	set "VolLabel=No drive connected."
	set "VolSerialNum=No drive connected."
)
IF "%VolLabel%"=="no" (
	IF "%VolLabel2%"=="label." set "VolLabel=no label."
)
del %Temp%\DEST-volume-info.txt
REM Set vars for the tests coming up next
IF "%VolLabel%"=="No drive connected." (set DestHostDrive=OFFLINE) ELSE (set DestHostDrive=ONLINE)
IF "%VolLabel%"=="%BAKUPLABEL%" (set CheckDestLabel=PASS) ELSE (set CheckDestLabel=FAIL)
IF "%VolSerialNum%"=="%BAKUPSERIAL%" (set CheckDestSerial=PASS) ELSE (set CheckDestSerial=FAIL)
:: Volume in drive F is LocalDataBackup
:: Volume Serial Number is E693-8EB5
REM Check what the CD /D and VOL tests gave us. Only continues if BOTH tests pass
IF NOT [%CDDestinationCheque%]==[ONLINE] (GOTO NoDestHost)
IF [%DestHostDrive%]==[ONLINE] (GOTO DestHostOnline) ELSE (GOTO NoDestHost)
:NoDestHost
REM Drive to be backed up is not online
echo %DRIVETOBAK%\%SOURCELABEL% Backup - Step 1/4: ERROR = Could not find %BAKUPLOC%\
echo(
echo Drive %BAKUPLOC%\ "%BAKUPLABEL%" appears to be disconnected.
echo(
echo SOLUTION: Shutdown, reconnect %BAKUPLOC%\, boot up, and run script again.
GOTO Fail
:DestHostOnline
REM Drive to be backed up (SOURCE) is online, now check if LABEL and SERIAL are correct...
REM Check Label
IF "%CheckDestLabel%"=="FAIL" (GOTO DESTLabelFail) ELSE (GOTO DESTLabelPass)
:DESTLabelFail
REM Labels don't match. Check serial anyway so we can specialize error code.
IF "%CheckDestSerial%"=="FAIL" (GOTO DEST3LabelFailSerialFail) ELSE (GOTO DEST2LabelFailSerialPass)
:DEST2LabelFailSerialPass
REM Labels didn't match, but Serial did? Some-thing or -body must've changed the label...
REM Label does not match, but serial *IS* correct. Wtf is going on here?
echo %DRIVETOBAK%\%SOURCELABEL% Backup - Step 1/4: ERROR = Volume %BAKUPLOC% Label was changed.
echo(
echo %BAKUPLOC%\ may have had an error. (or someone changed the drive label)
echo(
echo Serial numbers match, "%BAKUPSERIAL%"="%VolSerialNum%", but labels do not:
echo "%BAKUPLABEL%"=^!"%VolLabel%"
echo(
echo SOLUTION: Please check that the %BAKUPLOC% drive is correct. If it is o.k., please
echo change the drive label.
GOTO Fail
:DEST3LabelFailSerialFail
REM Both failed. Drive letter is in use by another drive, most likely
REM DESTINATION drive letter is online, but Label and serial do not match. Another drive is using it.
echo %DRIVETOBAK%\%SOURCELABEL% Backup - Step 1/4: ERROR = Drive letter %BAKUPLOC% is in use by
echo another drive.
echo(
echo The backup drive "%BAKUPLABEL%" requires %BAKUPLOC%\ to work.
echo(
echo SOLUTION: Shutdown, disconnect the drive occupying letter %BAKUPLOC%\, 
echo connect "%BAKUPLABEL%" to %BAKUPLOC%\, boot up, and run script again.
GOTO Fail
:DESTLabelPass
REM Label passed, now check Serial
IF "%CheckDestSerial%"=="FAIL" (GOTO DEST2LabelPassSerialFail) ELSE (GOTO DESTBothPass)
:DEST2LabelPassSerialFail
REM Labels passed, but Serials didn't. Sounds like different drive using same Label.
REM Label matches, but serials do not match
echo %DRIVETOBAK%\%SOURCELABEL% Backup - Step 1/4: ERROR = Volume %BAKUPLOC% serial does not match.
echo(
echo %BAKUPLOC%\ "%BAKUPLABEL%" serial number is supposed to be "%BAKUPSERIAL%"
echo Instead %BAKUPLOC%\ "%VolLabel%" serial number is: "%VolSerialNum%"
echo(
echo A different drive may be using the same label: "%BAKUPLABEL%"
echo(
echo SOLUTION: Shutdown, disconnect offending drive, connect %BAKUPLOC%\ "%BAKUPLABEL%",
echo boot up, and run script again.
GOTO Fail
:DESTBothPass
REM DESTINATION drive containing the VHD is online, and has correct SERIAL and LABEL!
REM End ptOneB: DESTINATION drive is online and correct!

:ptOneC
REM Check if VHD exists.
REM -------------------------------------------------------------------------------
REM Step 2: Mount VHD backup
REM Step C: check first that the *.vhd file exists in the right location
IF EXIST "%BAKUPIMG%" (GOTO FoundVHD) ELSE (GOTO VHDnotfound)
:VHDnotfound
REM Drive label and serial match! But no VHD?! Uh-oh.
echo %DRIVETOBAK%\%SOURCELABEL% Backup - Step 1/4: ERROR = Backup image could not be located.
echo(
echo Drive %BAKUPLOC%\"%BAKUPLABEL%" is supposed to contain "%BAKUPIMG%"
echo( 
echo Creating a new backup image is an intensive, long process. It could cause the
echo source drive to fail before the backup is finished, resulting in data loss.
echo(
echo SOLUTION: Search for "%BAKUPFILE%" in %BAKUPLOC%\.  
echo Search other locations too to make sure the image did not get deleted or moved.
echo If found, move to expected location or update location and run script again.
echo - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
echo If you've already searched for the backup and this is the *second time* you are
echo seeing this message, stay on the line to create a new backup image from scratch
pause
goto FreshBackupImage
:FoundVHD
REM Found VHD backup image!
echo %DRIVETOBAK%\%SOURCELABEL% Backup - Step 1/4: Found %BAKUPFILE%!
echo(
REM End ptOneC: VHD exists!
REM ===============================================================================
REM End ptOne! All locations Exist and Online! Our variables are good! BEGIN MOUNNNNNNNTT~~~~

:ptTwoA
REM Mount VHD with DISKPART. Check that mount was successful.
REM Step 2B: Attempt to mount VHD file to auto-detected drive letter
:: sel vdisk file="F:\shop_drivebox_d_data.VHD"
:: attach vdisk
:: select partition 1
:: assign letter=R
IF EXIST "%Temp%\mount-status.log" del "%Temp%\mount-status.log"
IF EXIST "%Temp%\mount-backup.txt" del "%Temp%\mount-backup.txt"
echo sel vdisk file="%BAKUPIMG%">%Temp%\mount-backup.txt
echo attach vdisk>>%Temp%\mount-backup.txt
REM When attaching a VHD using DISKPART like this, it seems to want to pick the same letter it used last time (in this case R:)
diskpart /s %Temp%\mount-backup.txt>%Temp%\mount-status.log 2>&1
del "%Temp%\mount-backup.txt"
REM DISKPART mount finished!
REM -------------------------------------------------------------------------------
REM Check DISKPART log to make sure mount was successful!
FOR /F "skip=5 tokens=*" %%G IN (%Temp%\mount-status.log) DO (set DPlineone=%%G&GOTO EscDPlineone)
:EscDPlineone
FOR /F "skip=7 tokens=*" %%G IN (%Temp%\mount-status.log) DO (set DPlinetwo=%%G&GOTO EscDPlinetwo)
:EscDPlinetwo
FOR /F "skip=8 tokens=* delims=" %%G IN (%Temp%\mount-status.log) DO (Call :FindDPlinethree "%%G")
GOTO:EOF
GOTO EscDPlinethree
:FindDPlinethree
IF NOT "%~1"=="" set "DPlinethree=%~1 %~2 %~3 %~4 %~5 %~6 %~7"&GOTO EscDPlinethree
GOTO:EOF
:EscDPlinethree
REM Already Attached:
:: DPlineone="DiskPart successfully selected the virtual disk file."
:: DPlinetwo="Virtual Disk Service error:"
:: DPlinethree="The virtual disk is already attached.      "
REM Successfully Attached:
:: DPlineone="DiskPart successfully selected the virtual disk file."
:: DPlinetwo="  100 percent completed"
:: DPlinethree="DiskPart successfully attached the virtual disk file.      "
FOR /F "tokens=2,3" %%G IN ("%DPlineone%") DO (
	set DPfirstlinechqu1=%%G
	set DPfirstlinechqu2=%%H
)
IF "%DPfirstlinechqu1%"=="successfully" (set MountDPLOne=GOOD) ELSE (set MountDPLOne=BAD)
IF NOT "%DPfirstlinechqu2%"=="selected" set "MountDPLOne=BAD"
REM Line one check finished, Now check third line 'cause that's the important one
REM - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
FOR /F "tokens=2,3,5,6" %%G IN ("%DPlinethree%") DO (
	set DPthirdlinechqu1=%%G
	set DPthirdlinechqu2=%%H
	set DPthirdlinechqu3=%%I
	set DPthirdlinechqu4=%%J
)
IF "%DPthirdlinechqu1%"=="successfully" (
	IF "%DPthirdlinechqu2%"=="attached" (set MountDPsuccess=YES) ELSE (set MountDPsuccess=NO)
) ELSE (set MountDPsuccess=NO)
IF "%DPthirdlinechqu3%"=="already" (
	IF "%DPthirdlinechqu4%"=="attached." (set MountDPalreadyatt=YES) ELSE (set MountDPalreadyatt=NO)
) ELSE (set MountDPalreadyatt=NO)
set "MountDPStatus="
IF "%MountDPsuccess%"=="YES" (set MountDPStatus=Success) ELSE (
	IF "%MountDPalreadyatt%"=="YES" (set MountDPStatus=AlrAttached) ELSE (set MountDPStatus=FAIL)
)
IF "%MountDPStatus%"=="AlrAttached" (
	IF "%MountDPLOne%"=="BAD" echo DISKPART: %DPlineone% & echo.
	echo "%BAKUPIMG%" was already attached. & echo.
)
IF "%MountDPStatus%"=="Success" (
	IF "%MountDPLOne%"=="BAD" echo DISKPART: %DPlineone% & echo.
)
IF NOT "%MountDPStatus%"=="FAIL" GOTO DPMountVHDsuccess
REM DISKPART VHD mount failed. Serious Error. Danger Will Robinson!
echo %DRIVETOBAK%\%SOURCELABEL% Backup - Step 2/4: ERROR = Failed to mount VHD.
echo(
echo DISKPART failed to mount backup image "%BAKUPIMG%"
echo(
echo DISKPART's response:
echo %DPlineone%
echo %DPlinetwo%
echo %DPlinethree%
"%Temp%\mount-status.log"
GOTO Fail
:DPMountVHDsuccess
del "%Temp%\mount-status.log"
REM -------------------------------------------------------------------------------
REM End ptTwoA! VHD was successfully mounted.

:ptTwoB
REM Check if VHD mounted at preferred letter. (check other letters if it's not)
REM Step 2C: Check that VHD was mounted o.k. (All were 'successful')
REM If we have a preferred drive letter for the VHD to be mounted to, check there first.
IF [%VHDPREFERRED%]==[] GOTO NoPreferredDriveLetter
REM -------------------------------------------------------------------------------
VOL %VHDPREFERRED% >%Temp%\vhd-preferred-check.txt 2>&1
REM Redirect all output *and* errors to one file
set "VolLabel="
set "VolLabel2="
set "VolSerialNum="
set "VolLineOne="
set "VolLineTwo="
FOR /F "tokens=*" %%G IN (%Temp%\vhd-preferred-check.txt) DO (set VolLineOne=%%G&GOTO VHDPreferredEsc)
:VHDPreferredEsc
FOR /F "tokens=*" %%G IN (%Temp%\vhd-preferred-check.txt) DO (set VolLineTwo=%%G)
FOR /F "tokens=6*" %%G IN ("%VolLineOne%") DO (set VolLabel=%%G&set VolLabel2=%%H)
FOR /F "tokens=5" %%G IN ("%VolLineTwo%") DO (set VolSerialNum=%%G)
IF "%VolLineOne%"=="The system cannot find the path specified." (
	set "VolLabel=No drive connected."
	set "VolSerialNum=No drive connected."
)
IF "%VolLabel%"=="no" (
	IF "%VolLabel2%"=="label." set "VolLabel=no label."
)
del %Temp%\vhd-preferred-check.txt
REM Is VHD here? Does VHD's volume label and serial match?
IF "%VolLabel%"=="%BAKUPVHDLABEL%" (set CheckTestVHDLabel=PASS) ELSE (set CheckTestVHDLabel=FAIL)
IF "%VolSerialNum%"=="%BAKUPVHDSERIAL%" (set CheckTestVHDSerial=PASS) ELSE (set CheckTestVHDSerial=FAIL)
IF "%CheckTestVHDSerial%"=="PASS" (
	set VHDDiscoveredLetter=%VHDPREFERRED%
	GOTO VHDSerialMatch
)
:NoPreferredDriveLetter
REM NOTIFY user:
echo Preferred VHD letter not found. Scanning all letters...
echo(
REM -------------------------------------------------------------------------------
REM The Preferred drive letter cannot help us. Now we must search everywhere for the VHD.
REM Since it's true the VHD could be mounted at A, or B, or even C, I see no other option but to scan every possible letter one-by-one.
:: i.e.:
:: ABCDEFGHIJKLMNOPQRSTUVWXYZ
REM Let's go!
REM - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
set "VHDscanletter=A:"
VOL %VHDscanletter% >%Temp%\vhd-iteration-check.txt 2>&1
REM Redirect all output *and* errors to one file
set "VolLabel="
set "VolLabel2="
set "VolSerialNum="
set "VolLineOne="
set "VolLineTwo="
FOR /F "tokens=*" %%G IN (%Temp%\vhd-iteration-check.txt) DO (set VolLineOne=%%G&GOTO VHDAIterationEsc)
:VHDAIterationEsc
FOR /F "tokens=*" %%G IN (%Temp%\vhd-iteration-check.txt) DO (set VolLineTwo=%%G)
FOR /F "tokens=6*" %%G IN ("%VolLineOne%") DO (set VolLabel=%%G&set VolLabel2=%%H)
FOR /F "tokens=5" %%G IN ("%VolLineTwo%") DO (set VolSerialNum=%%G)
IF "%VolLineOne%"=="The system cannot find the path specified." (
	set "VolLabel=No drive connected."
	set "VolSerialNum=No drive connected."
)
IF "%VolLabel%"=="no" (
	IF "%VolLabel2%"=="label." set "VolLabel=no label."
)
del %Temp%\vhd-iteration-check.txt
REM Is VHD here? Does VHD's volume label and serial match?
IF "%VolLabel%"=="%BAKUPVHDLABEL%" (set CheckTestVHDLabel=PASS) ELSE (set CheckTestVHDLabel=FAIL)
IF "%VolSerialNum%"=="%BAKUPVHDSERIAL%" (set CheckTestVHDSerial=PASS) ELSE (set CheckTestVHDSerial=FAIL)
IF "%CheckTestVHDSerial%"=="PASS" set VHDDiscoveredLetter=%VHDscanletter%&GOTO VHDSerialMatch
REM - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
set "VHDscanletter=B:"
VOL %VHDscanletter% >%Temp%\vhd-iteration-check.txt 2>&1
REM Redirect all output *and* errors to one file
set "VolLabel="
set "VolLabel2="
set "VolSerialNum="
set "VolLineOne="
set "VolLineTwo="
FOR /F "tokens=*" %%G IN (%Temp%\vhd-iteration-check.txt) DO (set VolLineOne=%%G&GOTO VHDBIterationEsc)
:VHDBIterationEsc
FOR /F "tokens=*" %%G IN (%Temp%\vhd-iteration-check.txt) DO (set VolLineTwo=%%G)
FOR /F "tokens=6*" %%G IN ("%VolLineOne%") DO (set VolLabel=%%G&set VolLabel2=%%H)
FOR /F "tokens=5" %%G IN ("%VolLineTwo%") DO (set VolSerialNum=%%G)
IF "%VolLineOne%"=="The system cannot find the path specified." (
	set "VolLabel=No drive connected."
	set "VolSerialNum=No drive connected."
)
IF "%VolLabel%"=="no" (
	IF "%VolLabel2%"=="label." set "VolLabel=no label."
)
del %Temp%\vhd-iteration-check.txt
REM Is VHD here? Does VHD's volume label and serial match?
IF "%VolLabel%"=="%BAKUPVHDLABEL%" (set CheckTestVHDLabel=PASS) ELSE (set CheckTestVHDLabel=FAIL)
IF "%VolSerialNum%"=="%BAKUPVHDSERIAL%" (set CheckTestVHDSerial=PASS) ELSE (set CheckTestVHDSerial=FAIL)
IF "%CheckTestVHDSerial%"=="PASS" set VHDDiscoveredLetter=%VHDscanletter%&GOTO VHDSerialMatch
REM - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
set "VHDscanletter=C:"
VOL %VHDscanletter% >%Temp%\vhd-iteration-check.txt 2>&1
REM Redirect all output *and* errors to one file
set "VolLabel="
set "VolLabel2="
set "VolSerialNum="
set "VolLineOne="
set "VolLineTwo="
FOR /F "tokens=*" %%G IN (%Temp%\vhd-iteration-check.txt) DO (set VolLineOne=%%G&GOTO VHDCIterationEsc)
:VHDCIterationEsc
FOR /F "tokens=*" %%G IN (%Temp%\vhd-iteration-check.txt) DO (set VolLineTwo=%%G)
FOR /F "tokens=6*" %%G IN ("%VolLineOne%") DO (set VolLabel=%%G&set VolLabel2=%%H)
FOR /F "tokens=5" %%G IN ("%VolLineTwo%") DO (set VolSerialNum=%%G)
IF "%VolLineOne%"=="The system cannot find the path specified." (
	set "VolLabel=No drive connected."
	set "VolSerialNum=No drive connected."
)
IF "%VolLabel%"=="no" (
	IF "%VolLabel2%"=="label." set "VolLabel=no label."
)
del %Temp%\vhd-iteration-check.txt
REM Is VHD here? Does VHD's volume label and serial match?
IF "%VolLabel%"=="%BAKUPVHDLABEL%" (set CheckTestVHDLabel=PASS) ELSE (set CheckTestVHDLabel=FAIL)
IF "%VolSerialNum%"=="%BAKUPVHDSERIAL%" (set CheckTestVHDSerial=PASS) ELSE (set CheckTestVHDSerial=FAIL)
IF "%CheckTestVHDSerial%"=="PASS" set VHDDiscoveredLetter=%VHDscanletter%&GOTO VHDSerialMatch
REM - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
set "VHDscanletter=D:"
VOL %VHDscanletter% >%Temp%\vhd-iteration-check.txt 2>&1
REM Redirect all output *and* errors to one file
set "VolLabel="
set "VolLabel2="
set "VolSerialNum="
set "VolLineOne="
set "VolLineTwo="
FOR /F "tokens=*" %%G IN (%Temp%\vhd-iteration-check.txt) DO (set VolLineOne=%%G&GOTO VHDDIterationEsc)
:VHDDIterationEsc
FOR /F "tokens=*" %%G IN (%Temp%\vhd-iteration-check.txt) DO (set VolLineTwo=%%G)
FOR /F "tokens=6*" %%G IN ("%VolLineOne%") DO (set VolLabel=%%G&set VolLabel2=%%H)
FOR /F "tokens=5" %%G IN ("%VolLineTwo%") DO (set VolSerialNum=%%G)
IF "%VolLineOne%"=="The system cannot find the path specified." (
	set "VolLabel=No drive connected."
	set "VolSerialNum=No drive connected."
)
IF "%VolLabel%"=="no" (
	IF "%VolLabel2%"=="label." set "VolLabel=no label."
)
del %Temp%\vhd-iteration-check.txt
REM Is VHD here? Does VHD's volume label and serial match?
IF "%VolLabel%"=="%BAKUPVHDLABEL%" (set CheckTestVHDLabel=PASS) ELSE (set CheckTestVHDLabel=FAIL)
IF "%VolSerialNum%"=="%BAKUPVHDSERIAL%" (set CheckTestVHDSerial=PASS) ELSE (set CheckTestVHDSerial=FAIL)
IF "%CheckTestVHDSerial%"=="PASS" set VHDDiscoveredLetter=%VHDscanletter%&GOTO VHDSerialMatch
REM - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
set "VHDscanletter=E:"
VOL %VHDscanletter% >%Temp%\vhd-iteration-check.txt 2>&1
REM Redirect all output *and* errors to one file
set "VolLabel="
set "VolLabel2="
set "VolSerialNum="
set "VolLineOne="
set "VolLineTwo="
FOR /F "tokens=*" %%G IN (%Temp%\vhd-iteration-check.txt) DO (set VolLineOne=%%G&GOTO VHDEIterationEsc)
:VHDEIterationEsc
FOR /F "tokens=*" %%G IN (%Temp%\vhd-iteration-check.txt) DO (set VolLineTwo=%%G)
FOR /F "tokens=6*" %%G IN ("%VolLineOne%") DO (set VolLabel=%%G&set VolLabel2=%%H)
FOR /F "tokens=5" %%G IN ("%VolLineTwo%") DO (set VolSerialNum=%%G)
IF "%VolLineOne%"=="The system cannot find the path specified." (
	set "VolLabel=No drive connected."
	set "VolSerialNum=No drive connected."
)
IF "%VolLabel%"=="no" (
	IF "%VolLabel2%"=="label." set "VolLabel=no label."
)
del %Temp%\vhd-iteration-check.txt
REM Is VHD here? Does VHD's volume label and serial match?
IF "%VolLabel%"=="%BAKUPVHDLABEL%" (set CheckTestVHDLabel=PASS) ELSE (set CheckTestVHDLabel=FAIL)
IF "%VolSerialNum%"=="%BAKUPVHDSERIAL%" (set CheckTestVHDSerial=PASS) ELSE (set CheckTestVHDSerial=FAIL)
IF "%CheckTestVHDSerial%"=="PASS" set VHDDiscoveredLetter=%VHDscanletter%&GOTO VHDSerialMatch
REM - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
set "VHDscanletter=F:"
VOL %VHDscanletter% >%Temp%\vhd-iteration-check.txt 2>&1
REM Redirect all output *and* errors to one file
set "VolLabel="
set "VolLabel2="
set "VolSerialNum="
set "VolLineOne="
set "VolLineTwo="
FOR /F "tokens=*" %%G IN (%Temp%\vhd-iteration-check.txt) DO (set VolLineOne=%%G&GOTO VHDFIterationEsc)
:VHDFIterationEsc
FOR /F "tokens=*" %%G IN (%Temp%\vhd-iteration-check.txt) DO (set VolLineTwo=%%G)
FOR /F "tokens=6*" %%G IN ("%VolLineOne%") DO (set VolLabel=%%G&set VolLabel2=%%H)
FOR /F "tokens=5" %%G IN ("%VolLineTwo%") DO (set VolSerialNum=%%G)
IF "%VolLineOne%"=="The system cannot find the path specified." (
	set "VolLabel=No drive connected."
	set "VolSerialNum=No drive connected."
)
IF "%VolLabel%"=="no" (
	IF "%VolLabel2%"=="label." set "VolLabel=no label."
)
del %Temp%\vhd-iteration-check.txt
REM Is VHD here? Does VHD's volume label and serial match?
IF "%VolLabel%"=="%BAKUPVHDLABEL%" (set CheckTestVHDLabel=PASS) ELSE (set CheckTestVHDLabel=FAIL)
IF "%VolSerialNum%"=="%BAKUPVHDSERIAL%" (set CheckTestVHDSerial=PASS) ELSE (set CheckTestVHDSerial=FAIL)
IF "%CheckTestVHDSerial%"=="PASS" set VHDDiscoveredLetter=%VHDscanletter%&GOTO VHDSerialMatch
REM - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
set "VHDscanletter=G:"
VOL %VHDscanletter% >%Temp%\vhd-iteration-check.txt 2>&1
REM Redirect all output *and* errors to one file
set "VolLabel="
set "VolLabel2="
set "VolSerialNum="
set "VolLineOne="
set "VolLineTwo="
FOR /F "tokens=*" %%G IN (%Temp%\vhd-iteration-check.txt) DO (set VolLineOne=%%G&GOTO VHDGIterationEsc)
:VHDGIterationEsc
FOR /F "tokens=*" %%G IN (%Temp%\vhd-iteration-check.txt) DO (set VolLineTwo=%%G)
FOR /F "tokens=6*" %%G IN ("%VolLineOne%") DO (set VolLabel=%%G&set VolLabel2=%%H)
FOR /F "tokens=5" %%G IN ("%VolLineTwo%") DO (set VolSerialNum=%%G)
IF "%VolLineOne%"=="The system cannot find the path specified." (
	set "VolLabel=No drive connected."
	set "VolSerialNum=No drive connected."
)
IF "%VolLabel%"=="no" (
	IF "%VolLabel2%"=="label." set "VolLabel=no label."
)
del %Temp%\vhd-iteration-check.txt
REM Is VHD here? Does VHD's volume label and serial match?
IF "%VolLabel%"=="%BAKUPVHDLABEL%" (set CheckTestVHDLabel=PASS) ELSE (set CheckTestVHDLabel=FAIL)
IF "%VolSerialNum%"=="%BAKUPVHDSERIAL%" (set CheckTestVHDSerial=PASS) ELSE (set CheckTestVHDSerial=FAIL)
IF "%CheckTestVHDSerial%"=="PASS" set VHDDiscoveredLetter=%VHDscanletter%&GOTO VHDSerialMatch
REM - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
set "VHDscanletter=H:"
VOL %VHDscanletter% >%Temp%\vhd-iteration-check.txt 2>&1
REM Redirect all output *and* errors to one file
set "VolLabel="
set "VolLabel2="
set "VolSerialNum="
set "VolLineOne="
set "VolLineTwo="
FOR /F "tokens=*" %%G IN (%Temp%\vhd-iteration-check.txt) DO (set VolLineOne=%%G&GOTO VHDHIterationEsc)
:VHDHIterationEsc
FOR /F "tokens=*" %%G IN (%Temp%\vhd-iteration-check.txt) DO (set VolLineTwo=%%G)
FOR /F "tokens=6*" %%G IN ("%VolLineOne%") DO (set VolLabel=%%G&set VolLabel2=%%H)
FOR /F "tokens=5" %%G IN ("%VolLineTwo%") DO (set VolSerialNum=%%G)
IF "%VolLineOne%"=="The system cannot find the path specified." (
	set "VolLabel=No drive connected."
	set "VolSerialNum=No drive connected."
)
IF "%VolLabel%"=="no" (
	IF "%VolLabel2%"=="label." set "VolLabel=no label."
)
del %Temp%\vhd-iteration-check.txt
REM Is VHD here? Does VHD's volume label and serial match?
IF "%VolLabel%"=="%BAKUPVHDLABEL%" (set CheckTestVHDLabel=PASS) ELSE (set CheckTestVHDLabel=FAIL)
IF "%VolSerialNum%"=="%BAKUPVHDSERIAL%" (set CheckTestVHDSerial=PASS) ELSE (set CheckTestVHDSerial=FAIL)
IF "%CheckTestVHDSerial%"=="PASS" set VHDDiscoveredLetter=%VHDscanletter%&GOTO VHDSerialMatch
REM - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
set "VHDscanletter=I:"
VOL %VHDscanletter% >%Temp%\vhd-iteration-check.txt 2>&1
REM Redirect all output *and* errors to one file
set "VolLabel="
set "VolLabel2="
set "VolSerialNum="
set "VolLineOne="
set "VolLineTwo="
FOR /F "tokens=*" %%G IN (%Temp%\vhd-iteration-check.txt) DO (set VolLineOne=%%G&GOTO VHDIIterationEsc)
:VHDIIterationEsc
FOR /F "tokens=*" %%G IN (%Temp%\vhd-iteration-check.txt) DO (set VolLineTwo=%%G)
FOR /F "tokens=6*" %%G IN ("%VolLineOne%") DO (set VolLabel=%%G&set VolLabel2=%%H)
FOR /F "tokens=5" %%G IN ("%VolLineTwo%") DO (set VolSerialNum=%%G)
IF "%VolLineOne%"=="The system cannot find the path specified." (
	set "VolLabel=No drive connected."
	set "VolSerialNum=No drive connected."
)
IF "%VolLabel%"=="no" (
	IF "%VolLabel2%"=="label." set "VolLabel=no label."
)
del %Temp%\vhd-iteration-check.txt
REM Is VHD here? Does VHD's volume label and serial match?
IF "%VolLabel%"=="%BAKUPVHDLABEL%" (set CheckTestVHDLabel=PASS) ELSE (set CheckTestVHDLabel=FAIL)
IF "%VolSerialNum%"=="%BAKUPVHDSERIAL%" (set CheckTestVHDSerial=PASS) ELSE (set CheckTestVHDSerial=FAIL)
IF "%CheckTestVHDSerial%"=="PASS" set VHDDiscoveredLetter=%VHDscanletter%&GOTO VHDSerialMatch
REM - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
set "VHDscanletter=J:"
VOL %VHDscanletter% >%Temp%\vhd-iteration-check.txt 2>&1
REM Redirect all output *and* errors to one file
set "VolLabel="
set "VolLabel2="
set "VolSerialNum="
set "VolLineOne="
set "VolLineTwo="
FOR /F "tokens=*" %%G IN (%Temp%\vhd-iteration-check.txt) DO (set VolLineOne=%%G&GOTO VHDJIterationEsc)
:VHDJIterationEsc
FOR /F "tokens=*" %%G IN (%Temp%\vhd-iteration-check.txt) DO (set VolLineTwo=%%G)
FOR /F "tokens=6*" %%G IN ("%VolLineOne%") DO (set VolLabel=%%G&set VolLabel2=%%H)
FOR /F "tokens=5" %%G IN ("%VolLineTwo%") DO (set VolSerialNum=%%G)
IF "%VolLineOne%"=="The system cannot find the path specified." (
	set "VolLabel=No drive connected."
	set "VolSerialNum=No drive connected."
)
IF "%VolLabel%"=="no" (
	IF "%VolLabel2%"=="label." set "VolLabel=no label."
)
del %Temp%\vhd-iteration-check.txt
REM Is VHD here? Does VHD's volume label and serial match?
IF "%VolLabel%"=="%BAKUPVHDLABEL%" (set CheckTestVHDLabel=PASS) ELSE (set CheckTestVHDLabel=FAIL)
IF "%VolSerialNum%"=="%BAKUPVHDSERIAL%" (set CheckTestVHDSerial=PASS) ELSE (set CheckTestVHDSerial=FAIL)
IF "%CheckTestVHDSerial%"=="PASS" set VHDDiscoveredLetter=%VHDscanletter%&GOTO VHDSerialMatch
REM - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
set "VHDscanletter=K:"
VOL %VHDscanletter% >%Temp%\vhd-iteration-check.txt 2>&1
REM Redirect all output *and* errors to one file
set "VolLabel="
set "VolLabel2="
set "VolSerialNum="
set "VolLineOne="
set "VolLineTwo="
FOR /F "tokens=*" %%G IN (%Temp%\vhd-iteration-check.txt) DO (set VolLineOne=%%G&GOTO VHDKIterationEsc)
:VHDKIterationEsc
FOR /F "tokens=*" %%G IN (%Temp%\vhd-iteration-check.txt) DO (set VolLineTwo=%%G)
FOR /F "tokens=6*" %%G IN ("%VolLineOne%") DO (set VolLabel=%%G&set VolLabel2=%%H)
FOR /F "tokens=5" %%G IN ("%VolLineTwo%") DO (set VolSerialNum=%%G)
IF "%VolLineOne%"=="The system cannot find the path specified." (
	set "VolLabel=No drive connected."
	set "VolSerialNum=No drive connected."
)
IF "%VolLabel%"=="no" (
	IF "%VolLabel2%"=="label." set "VolLabel=no label."
)
del %Temp%\vhd-iteration-check.txt
REM Is VHD here? Does VHD's volume label and serial match?
IF "%VolLabel%"=="%BAKUPVHDLABEL%" (set CheckTestVHDLabel=PASS) ELSE (set CheckTestVHDLabel=FAIL)
IF "%VolSerialNum%"=="%BAKUPVHDSERIAL%" (set CheckTestVHDSerial=PASS) ELSE (set CheckTestVHDSerial=FAIL)
IF "%CheckTestVHDSerial%"=="PASS" set VHDDiscoveredLetter=%VHDscanletter%&GOTO VHDSerialMatch
REM - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
set "VHDscanletter=L:"
VOL %VHDscanletter% >%Temp%\vhd-iteration-check.txt 2>&1
REM Redirect all output *and* errors to one file
set "VolLabel="
set "VolLabel2="
set "VolSerialNum="
set "VolLineOne="
set "VolLineTwo="
FOR /F "tokens=*" %%G IN (%Temp%\vhd-iteration-check.txt) DO (set VolLineOne=%%G&GOTO VHDLIterationEsc)
:VHDLIterationEsc
FOR /F "tokens=*" %%G IN (%Temp%\vhd-iteration-check.txt) DO (set VolLineTwo=%%G)
FOR /F "tokens=6*" %%G IN ("%VolLineOne%") DO (set VolLabel=%%G&set VolLabel2=%%H)
FOR /F "tokens=5" %%G IN ("%VolLineTwo%") DO (set VolSerialNum=%%G)
IF "%VolLineOne%"=="The system cannot find the path specified." (
	set "VolLabel=No drive connected."
	set "VolSerialNum=No drive connected."
)
IF "%VolLabel%"=="no" (
	IF "%VolLabel2%"=="label." set "VolLabel=no label."
)
del %Temp%\vhd-iteration-check.txt
REM Is VHD here? Does VHD's volume label and serial match?
IF "%VolLabel%"=="%BAKUPVHDLABEL%" (set CheckTestVHDLabel=PASS) ELSE (set CheckTestVHDLabel=FAIL)
IF "%VolSerialNum%"=="%BAKUPVHDSERIAL%" (set CheckTestVHDSerial=PASS) ELSE (set CheckTestVHDSerial=FAIL)
IF "%CheckTestVHDSerial%"=="PASS" set VHDDiscoveredLetter=%VHDscanletter%&GOTO VHDSerialMatch
REM - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
set "VHDscanletter=M:"
VOL %VHDscanletter% >%Temp%\vhd-iteration-check.txt 2>&1
REM Redirect all output *and* errors to one file
set "VolLabel="
set "VolLabel2="
set "VolSerialNum="
set "VolLineOne="
set "VolLineTwo="
FOR /F "tokens=*" %%G IN (%Temp%\vhd-iteration-check.txt) DO (set VolLineOne=%%G&GOTO VHDMIterationEsc)
:VHDMIterationEsc
FOR /F "tokens=*" %%G IN (%Temp%\vhd-iteration-check.txt) DO (set VolLineTwo=%%G)
FOR /F "tokens=6*" %%G IN ("%VolLineOne%") DO (set VolLabel=%%G&set VolLabel2=%%H)
FOR /F "tokens=5" %%G IN ("%VolLineTwo%") DO (set VolSerialNum=%%G)
IF "%VolLineOne%"=="The system cannot find the path specified." (
	set "VolLabel=No drive connected."
	set "VolSerialNum=No drive connected."
)
IF "%VolLabel%"=="no" (
	IF "%VolLabel2%"=="label." set "VolLabel=no label."
)
del %Temp%\vhd-iteration-check.txt
REM Is VHD here? Does VHD's volume label and serial match?
IF "%VolLabel%"=="%BAKUPVHDLABEL%" (set CheckTestVHDLabel=PASS) ELSE (set CheckTestVHDLabel=FAIL)
IF "%VolSerialNum%"=="%BAKUPVHDSERIAL%" (set CheckTestVHDSerial=PASS) ELSE (set CheckTestVHDSerial=FAIL)
IF "%CheckTestVHDSerial%"=="PASS" set VHDDiscoveredLetter=%VHDscanletter%&GOTO VHDSerialMatch
REM - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
set "VHDscanletter=N:"
VOL %VHDscanletter% >%Temp%\vhd-iteration-check.txt 2>&1
REM Redirect all output *and* errors to one file
set "VolLabel="
set "VolLabel2="
set "VolSerialNum="
set "VolLineOne="
set "VolLineTwo="
FOR /F "tokens=*" %%G IN (%Temp%\vhd-iteration-check.txt) DO (set VolLineOne=%%G&GOTO VHDNIterationEsc)
:VHDNIterationEsc
FOR /F "tokens=*" %%G IN (%Temp%\vhd-iteration-check.txt) DO (set VolLineTwo=%%G)
FOR /F "tokens=6*" %%G IN ("%VolLineOne%") DO (set VolLabel=%%G&set VolLabel2=%%H)
FOR /F "tokens=5" %%G IN ("%VolLineTwo%") DO (set VolSerialNum=%%G)
IF "%VolLineOne%"=="The system cannot find the path specified." (
	set "VolLabel=No drive connected."
	set "VolSerialNum=No drive connected."
)
IF "%VolLabel%"=="no" (
	IF "%VolLabel2%"=="label." set "VolLabel=no label."
)
del %Temp%\vhd-iteration-check.txt
REM Is VHD here? Does VHD's volume label and serial match?
IF "%VolLabel%"=="%BAKUPVHDLABEL%" (set CheckTestVHDLabel=PASS) ELSE (set CheckTestVHDLabel=FAIL)
IF "%VolSerialNum%"=="%BAKUPVHDSERIAL%" (set CheckTestVHDSerial=PASS) ELSE (set CheckTestVHDSerial=FAIL)
IF "%CheckTestVHDSerial%"=="PASS" set VHDDiscoveredLetter=%VHDscanletter%&GOTO VHDSerialMatch
REM - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
set "VHDscanletter=O:"
VOL %VHDscanletter% >%Temp%\vhd-iteration-check.txt 2>&1
REM Redirect all output *and* errors to one file
set "VolLabel="
set "VolLabel2="
set "VolSerialNum="
set "VolLineOne="
set "VolLineTwo="
FOR /F "tokens=*" %%G IN (%Temp%\vhd-iteration-check.txt) DO (set VolLineOne=%%G&GOTO VHDOIterationEsc)
:VHDOIterationEsc
FOR /F "tokens=*" %%G IN (%Temp%\vhd-iteration-check.txt) DO (set VolLineTwo=%%G)
FOR /F "tokens=6*" %%G IN ("%VolLineOne%") DO (set VolLabel=%%G&set VolLabel2=%%H)
FOR /F "tokens=5" %%G IN ("%VolLineTwo%") DO (set VolSerialNum=%%G)
IF "%VolLineOne%"=="The system cannot find the path specified." (
	set "VolLabel=No drive connected."
	set "VolSerialNum=No drive connected."
)
IF "%VolLabel%"=="no" (
	IF "%VolLabel2%"=="label." set "VolLabel=no label."
)
del %Temp%\vhd-iteration-check.txt
REM Is VHD here? Does VHD's volume label and serial match?
IF "%VolLabel%"=="%BAKUPVHDLABEL%" (set CheckTestVHDLabel=PASS) ELSE (set CheckTestVHDLabel=FAIL)
IF "%VolSerialNum%"=="%BAKUPVHDSERIAL%" (set CheckTestVHDSerial=PASS) ELSE (set CheckTestVHDSerial=FAIL)
IF "%CheckTestVHDSerial%"=="PASS" set VHDDiscoveredLetter=%VHDscanletter%&GOTO VHDSerialMatch
REM - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
set "VHDscanletter=P:"
VOL %VHDscanletter% >%Temp%\vhd-iteration-check.txt 2>&1
REM Redirect all output *and* errors to one file
set "VolLabel="
set "VolLabel2="
set "VolSerialNum="
set "VolLineOne="
set "VolLineTwo="
FOR /F "tokens=*" %%G IN (%Temp%\vhd-iteration-check.txt) DO (set VolLineOne=%%G&GOTO VHDPIterationEsc)
:VHDPIterationEsc
FOR /F "tokens=*" %%G IN (%Temp%\vhd-iteration-check.txt) DO (set VolLineTwo=%%G)
FOR /F "tokens=6*" %%G IN ("%VolLineOne%") DO (set VolLabel=%%G&set VolLabel2=%%H)
FOR /F "tokens=5" %%G IN ("%VolLineTwo%") DO (set VolSerialNum=%%G)
IF "%VolLineOne%"=="The system cannot find the path specified." (
	set "VolLabel=No drive connected."
	set "VolSerialNum=No drive connected."
)
IF "%VolLabel%"=="no" (
	IF "%VolLabel2%"=="label." set "VolLabel=no label."
)
del %Temp%\vhd-iteration-check.txt
REM Is VHD here? Does VHD's volume label and serial match?
IF "%VolLabel%"=="%BAKUPVHDLABEL%" (set CheckTestVHDLabel=PASS) ELSE (set CheckTestVHDLabel=FAIL)
IF "%VolSerialNum%"=="%BAKUPVHDSERIAL%" (set CheckTestVHDSerial=PASS) ELSE (set CheckTestVHDSerial=FAIL)
IF "%CheckTestVHDSerial%"=="PASS" set VHDDiscoveredLetter=%VHDscanletter%&GOTO VHDSerialMatch
REM - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
set "VHDscanletter=Q:"
VOL %VHDscanletter% >%Temp%\vhd-iteration-check.txt 2>&1
REM Redirect all output *and* errors to one file
set "VolLabel="
set "VolLabel2="
set "VolSerialNum="
set "VolLineOne="
set "VolLineTwo="
FOR /F "tokens=*" %%G IN (%Temp%\vhd-iteration-check.txt) DO (set VolLineOne=%%G&GOTO VHDQIterationEsc)
:VHDQIterationEsc
FOR /F "tokens=*" %%G IN (%Temp%\vhd-iteration-check.txt) DO (set VolLineTwo=%%G)
FOR /F "tokens=6*" %%G IN ("%VolLineOne%") DO (set VolLabel=%%G&set VolLabel2=%%H)
FOR /F "tokens=5" %%G IN ("%VolLineTwo%") DO (set VolSerialNum=%%G)
IF "%VolLineOne%"=="The system cannot find the path specified." (
	set "VolLabel=No drive connected."
	set "VolSerialNum=No drive connected."
)
IF "%VolLabel%"=="no" (
	IF "%VolLabel2%"=="label." set "VolLabel=no label."
)
del %Temp%\vhd-iteration-check.txt
REM Is VHD here? Does VHD's volume label and serial match?
IF "%VolLabel%"=="%BAKUPVHDLABEL%" (set CheckTestVHDLabel=PASS) ELSE (set CheckTestVHDLabel=FAIL)
IF "%VolSerialNum%"=="%BAKUPVHDSERIAL%" (set CheckTestVHDSerial=PASS) ELSE (set CheckTestVHDSerial=FAIL)
IF "%CheckTestVHDSerial%"=="PASS" set VHDDiscoveredLetter=%VHDscanletter%&GOTO VHDSerialMatch
REM - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
set "VHDscanletter=R:"
VOL %VHDscanletter% >%Temp%\vhd-iteration-check.txt 2>&1
REM Redirect all output *and* errors to one file
set "VolLabel="
set "VolLabel2="
set "VolSerialNum="
set "VolLineOne="
set "VolLineTwo="
FOR /F "tokens=*" %%G IN (%Temp%\vhd-iteration-check.txt) DO (set VolLineOne=%%G&GOTO VHDRIterationEsc)
:VHDRIterationEsc
FOR /F "tokens=*" %%G IN (%Temp%\vhd-iteration-check.txt) DO (set VolLineTwo=%%G)
FOR /F "tokens=6*" %%G IN ("%VolLineOne%") DO (set VolLabel=%%G&set VolLabel2=%%H)
FOR /F "tokens=5" %%G IN ("%VolLineTwo%") DO (set VolSerialNum=%%G)
IF "%VolLineOne%"=="The system cannot find the path specified." (
	set "VolLabel=No drive connected."
	set "VolSerialNum=No drive connected."
)
IF "%VolLabel%"=="no" (
	IF "%VolLabel2%"=="label." set "VolLabel=no label."
)
del %Temp%\vhd-iteration-check.txt
REM Is VHD here? Does VHD's volume label and serial match?
IF "%VolLabel%"=="%BAKUPVHDLABEL%" (set CheckTestVHDLabel=PASS) ELSE (set CheckTestVHDLabel=FAIL)
IF "%VolSerialNum%"=="%BAKUPVHDSERIAL%" (set CheckTestVHDSerial=PASS) ELSE (set CheckTestVHDSerial=FAIL)
IF "%CheckTestVHDSerial%"=="PASS" set VHDDiscoveredLetter=%VHDscanletter%&GOTO VHDSerialMatch
REM - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
set "VHDscanletter=S:"
VOL %VHDscanletter% >%Temp%\vhd-iteration-check.txt 2>&1
REM Redirect all output *and* errors to one file
set "VolLabel="
set "VolLabel2="
set "VolSerialNum="
set "VolLineOne="
set "VolLineTwo="
FOR /F "tokens=*" %%G IN (%Temp%\vhd-iteration-check.txt) DO (set VolLineOne=%%G&GOTO VHDSIterationEsc)
:VHDSIterationEsc
FOR /F "tokens=*" %%G IN (%Temp%\vhd-iteration-check.txt) DO (set VolLineTwo=%%G)
FOR /F "tokens=6*" %%G IN ("%VolLineOne%") DO (set VolLabel=%%G&set VolLabel2=%%H)
FOR /F "tokens=5" %%G IN ("%VolLineTwo%") DO (set VolSerialNum=%%G)
IF "%VolLineOne%"=="The system cannot find the path specified." (
	set "VolLabel=No drive connected."
	set "VolSerialNum=No drive connected."
)
IF "%VolLabel%"=="no" (
	IF "%VolLabel2%"=="label." set "VolLabel=no label."
)
del %Temp%\vhd-iteration-check.txt
REM Is VHD here? Does VHD's volume label and serial match?
IF "%VolLabel%"=="%BAKUPVHDLABEL%" (set CheckTestVHDLabel=PASS) ELSE (set CheckTestVHDLabel=FAIL)
IF "%VolSerialNum%"=="%BAKUPVHDSERIAL%" (set CheckTestVHDSerial=PASS) ELSE (set CheckTestVHDSerial=FAIL)
IF "%CheckTestVHDSerial%"=="PASS" set VHDDiscoveredLetter=%VHDscanletter%&GOTO VHDSerialMatch
REM - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
set "VHDscanletter=T:"
VOL %VHDscanletter% >%Temp%\vhd-iteration-check.txt 2>&1
REM Redirect all output *and* errors to one file
set "VolLabel="
set "VolLabel2="
set "VolSerialNum="
set "VolLineOne="
set "VolLineTwo="
FOR /F "tokens=*" %%G IN (%Temp%\vhd-iteration-check.txt) DO (set VolLineOne=%%G&GOTO VHDTIterationEsc)
:VHDTIterationEsc
FOR /F "tokens=*" %%G IN (%Temp%\vhd-iteration-check.txt) DO (set VolLineTwo=%%G)
FOR /F "tokens=6*" %%G IN ("%VolLineOne%") DO (set VolLabel=%%G&set VolLabel2=%%H)
FOR /F "tokens=5" %%G IN ("%VolLineTwo%") DO (set VolSerialNum=%%G)
IF "%VolLineOne%"=="The system cannot find the path specified." (
	set "VolLabel=No drive connected."
	set "VolSerialNum=No drive connected."
)
IF "%VolLabel%"=="no" (
	IF "%VolLabel2%"=="label." set "VolLabel=no label."
)
del %Temp%\vhd-iteration-check.txt
REM Is VHD here? Does VHD's volume label and serial match?
IF "%VolLabel%"=="%BAKUPVHDLABEL%" (set CheckTestVHDLabel=PASS) ELSE (set CheckTestVHDLabel=FAIL)
IF "%VolSerialNum%"=="%BAKUPVHDSERIAL%" (set CheckTestVHDSerial=PASS) ELSE (set CheckTestVHDSerial=FAIL)
IF "%CheckTestVHDSerial%"=="PASS" set VHDDiscoveredLetter=%VHDscanletter%&GOTO VHDSerialMatch
REM - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
set "VHDscanletter=U:"
VOL %VHDscanletter% >%Temp%\vhd-iteration-check.txt 2>&1
REM Redirect all output *and* errors to one file
set "VolLabel="
set "VolLabel2="
set "VolSerialNum="
set "VolLineOne="
set "VolLineTwo="
FOR /F "tokens=*" %%G IN (%Temp%\vhd-iteration-check.txt) DO (set VolLineOne=%%G&GOTO VHDUIterationEsc)
:VHDUIterationEsc
FOR /F "tokens=*" %%G IN (%Temp%\vhd-iteration-check.txt) DO (set VolLineTwo=%%G)
FOR /F "tokens=6*" %%G IN ("%VolLineOne%") DO (set VolLabel=%%G&set VolLabel2=%%H)
FOR /F "tokens=5" %%G IN ("%VolLineTwo%") DO (set VolSerialNum=%%G)
IF "%VolLineOne%"=="The system cannot find the path specified." (
	set "VolLabel=No drive connected."
	set "VolSerialNum=No drive connected."
)
IF "%VolLabel%"=="no" (
	IF "%VolLabel2%"=="label." set "VolLabel=no label."
)
del %Temp%\vhd-iteration-check.txt
REM Is VHD here? Does VHD's volume label and serial match?
IF "%VolLabel%"=="%BAKUPVHDLABEL%" (set CheckTestVHDLabel=PASS) ELSE (set CheckTestVHDLabel=FAIL)
IF "%VolSerialNum%"=="%BAKUPVHDSERIAL%" (set CheckTestVHDSerial=PASS) ELSE (set CheckTestVHDSerial=FAIL)
IF "%CheckTestVHDSerial%"=="PASS" set VHDDiscoveredLetter=%VHDscanletter%&GOTO VHDSerialMatch
REM - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
set "VHDscanletter=V:"
VOL %VHDscanletter% >%Temp%\vhd-iteration-check.txt 2>&1
REM Redirect all output *and* errors to one file
set "VolLabel="
set "VolLabel2="
set "VolSerialNum="
set "VolLineOne="
set "VolLineTwo="
FOR /F "tokens=*" %%G IN (%Temp%\vhd-iteration-check.txt) DO (set VolLineOne=%%G&GOTO VHDVIterationEsc)
:VHDVIterationEsc
FOR /F "tokens=*" %%G IN (%Temp%\vhd-iteration-check.txt) DO (set VolLineTwo=%%G)
FOR /F "tokens=6*" %%G IN ("%VolLineOne%") DO (set VolLabel=%%G&set VolLabel2=%%H)
FOR /F "tokens=5" %%G IN ("%VolLineTwo%") DO (set VolSerialNum=%%G)
IF "%VolLineOne%"=="The system cannot find the path specified." (
	set "VolLabel=No drive connected."
	set "VolSerialNum=No drive connected."
)
IF "%VolLabel%"=="no" (
	IF "%VolLabel2%"=="label." set "VolLabel=no label."
)
del %Temp%\vhd-iteration-check.txt
REM Is VHD here? Does VHD's volume label and serial match?
IF "%VolLabel%"=="%BAKUPVHDLABEL%" (set CheckTestVHDLabel=PASS) ELSE (set CheckTestVHDLabel=FAIL)
IF "%VolSerialNum%"=="%BAKUPVHDSERIAL%" (set CheckTestVHDSerial=PASS) ELSE (set CheckTestVHDSerial=FAIL)
IF "%CheckTestVHDSerial%"=="PASS" set VHDDiscoveredLetter=%VHDscanletter%&GOTO VHDSerialMatch
REM - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
set "VHDscanletter=W:"
VOL %VHDscanletter% >%Temp%\vhd-iteration-check.txt 2>&1
REM Redirect all output *and* errors to one file
set "VolLabel="
set "VolLabel2="
set "VolSerialNum="
set "VolLineOne="
set "VolLineTwo="
FOR /F "tokens=*" %%G IN (%Temp%\vhd-iteration-check.txt) DO (set VolLineOne=%%G&GOTO VHDWIterationEsc)
:VHDWIterationEsc
FOR /F "tokens=*" %%G IN (%Temp%\vhd-iteration-check.txt) DO (set VolLineTwo=%%G)
FOR /F "tokens=6*" %%G IN ("%VolLineOne%") DO (set VolLabel=%%G&set VolLabel2=%%H)
FOR /F "tokens=5" %%G IN ("%VolLineTwo%") DO (set VolSerialNum=%%G)
IF "%VolLineOne%"=="The system cannot find the path specified." (
	set "VolLabel=No drive connected."
	set "VolSerialNum=No drive connected."
)
IF "%VolLabel%"=="no" (
	IF "%VolLabel2%"=="label." set "VolLabel=no label."
)
del %Temp%\vhd-iteration-check.txt
REM Is VHD here? Does VHD's volume label and serial match?
IF "%VolLabel%"=="%BAKUPVHDLABEL%" (set CheckTestVHDLabel=PASS) ELSE (set CheckTestVHDLabel=FAIL)
IF "%VolSerialNum%"=="%BAKUPVHDSERIAL%" (set CheckTestVHDSerial=PASS) ELSE (set CheckTestVHDSerial=FAIL)
IF "%CheckTestVHDSerial%"=="PASS" set VHDDiscoveredLetter=%VHDscanletter%&GOTO VHDSerialMatch
REM - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
set "VHDscanletter=X:"
VOL %VHDscanletter% >%Temp%\vhd-iteration-check.txt 2>&1
REM Redirect all output *and* errors to one file
set "VolLabel="
set "VolLabel2="
set "VolSerialNum="
set "VolLineOne="
set "VolLineTwo="
FOR /F "tokens=*" %%G IN (%Temp%\vhd-iteration-check.txt) DO (set VolLineOne=%%G&GOTO VHDXIterationEsc)
:VHDXIterationEsc
FOR /F "tokens=*" %%G IN (%Temp%\vhd-iteration-check.txt) DO (set VolLineTwo=%%G)
FOR /F "tokens=6*" %%G IN ("%VolLineOne%") DO (set VolLabel=%%G&set VolLabel2=%%H)
FOR /F "tokens=5" %%G IN ("%VolLineTwo%") DO (set VolSerialNum=%%G)
IF "%VolLineOne%"=="The system cannot find the path specified." (
	set "VolLabel=No drive connected."
	set "VolSerialNum=No drive connected."
)
IF "%VolLabel%"=="no" (
	IF "%VolLabel2%"=="label." set "VolLabel=no label."
)
del %Temp%\vhd-iteration-check.txt
REM Is VHD here? Does VHD's volume label and serial match?
IF "%VolLabel%"=="%BAKUPVHDLABEL%" (set CheckTestVHDLabel=PASS) ELSE (set CheckTestVHDLabel=FAIL)
IF "%VolSerialNum%"=="%BAKUPVHDSERIAL%" (set CheckTestVHDSerial=PASS) ELSE (set CheckTestVHDSerial=FAIL)
IF "%CheckTestVHDSerial%"=="PASS" set VHDDiscoveredLetter=%VHDscanletter%&GOTO VHDSerialMatch
REM - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
set "VHDscanletter=Y:"
VOL %VHDscanletter% >%Temp%\vhd-iteration-check.txt 2>&1
REM Redirect all output *and* errors to one file
set "VolLabel="
set "VolLabel2="
set "VolSerialNum="
set "VolLineOne="
set "VolLineTwo="
FOR /F "tokens=*" %%G IN (%Temp%\vhd-iteration-check.txt) DO (set VolLineOne=%%G&GOTO VHDYIterationEsc)
:VHDYIterationEsc
FOR /F "tokens=*" %%G IN (%Temp%\vhd-iteration-check.txt) DO (set VolLineTwo=%%G)
FOR /F "tokens=6*" %%G IN ("%VolLineOne%") DO (set VolLabel=%%G&set VolLabel2=%%H)
FOR /F "tokens=5" %%G IN ("%VolLineTwo%") DO (set VolSerialNum=%%G)
IF "%VolLineOne%"=="The system cannot find the path specified." (
	set "VolLabel=No drive connected."
	set "VolSerialNum=No drive connected."
)
IF "%VolLabel%"=="no" (
	IF "%VolLabel2%"=="label." set "VolLabel=no label."
)
del %Temp%\vhd-iteration-check.txt
REM Is VHD here? Does VHD's volume label and serial match?
IF "%VolLabel%"=="%BAKUPVHDLABEL%" (set CheckTestVHDLabel=PASS) ELSE (set CheckTestVHDLabel=FAIL)
IF "%VolSerialNum%"=="%BAKUPVHDSERIAL%" (set CheckTestVHDSerial=PASS) ELSE (set CheckTestVHDSerial=FAIL)
IF "%CheckTestVHDSerial%"=="PASS" set VHDDiscoveredLetter=%VHDscanletter%&GOTO VHDSerialMatch
REM - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
set "VHDscanletter=Z:"
VOL %VHDscanletter% >%Temp%\vhd-iteration-check.txt 2>&1
REM Redirect all output *and* errors to one file
set "VolLabel="
set "VolLabel2="
set "VolSerialNum="
set "VolLineOne="
set "VolLineTwo="
FOR /F "tokens=*" %%G IN (%Temp%\vhd-iteration-check.txt) DO (set VolLineOne=%%G&GOTO VHDZIterationEsc)
:VHDZIterationEsc
FOR /F "tokens=*" %%G IN (%Temp%\vhd-iteration-check.txt) DO (set VolLineTwo=%%G)
FOR /F "tokens=6*" %%G IN ("%VolLineOne%") DO (set VolLabel=%%G&set VolLabel2=%%H)
FOR /F "tokens=5" %%G IN ("%VolLineTwo%") DO (set VolSerialNum=%%G)
IF "%VolLineOne%"=="The system cannot find the path specified." (
	set "VolLabel=No drive connected."
	set "VolSerialNum=No drive connected."
)
IF "%VolLabel%"=="no" (
	IF "%VolLabel2%"=="label." set "VolLabel=no label."
)
del %Temp%\vhd-iteration-check.txt
REM Is VHD here? Does VHD's volume label and serial match?
IF "%VolLabel%"=="%BAKUPVHDLABEL%" (set CheckTestVHDLabel=PASS) ELSE (set CheckTestVHDLabel=FAIL)
IF "%VolSerialNum%"=="%BAKUPVHDSERIAL%" (set CheckTestVHDSerial=PASS) ELSE (set CheckTestVHDSerial=FAIL)
IF "%CheckTestVHDSerial%"=="PASS" set VHDDiscoveredLetter=%VHDscanletter%&GOTO VHDSerialMatch
REM - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
REM ===============================================================================
REM ===============================================================================
REM -------------------------------------------------------------------------------
:VHDNoSerialNoMatch
REM Not a single drive's serial number matched what we should has for the VHD's serial.
echo %DRIVETOBAK%\%SOURCELABEL% Backup - Step 2/4: ERROR = Could not find VHD letter
echo(
echo "%BAKUPFILE%" was supposedly mounted successfully.
echo(
echo But a drive matching serial number "%BAKUPVHDSERIAL%" could not be found.
echo(
echo SOLUTION: Mount VHD manually, run VOL command, and ensure correct serial number
echo defined in the beginning of this script is set for variable "BAKUPVHDSERIAL".
GOTO Fail
:VHDSerialMatch
REM We got a match for the SERIAL number at drive letter %VHDDiscoveredLetter% Wooo!
REM Now lets check that the Labels match too just for shits.
IF "%CheckTestVHDLabel%"=="PASS" GOTO VHDLabelMatchToo
REM Uh-oh. Serials do match, but Labels don't? Let's see what user wants to do:
echo %DRIVETOBAK%\%SOURCELABEL% Backup - Step 2/4: WARNING = VHD labels do not match
echo(
echo However, serial number for "%BAKUPFILE%" does match.
echo(
echo Serial number %BAKUPVHDSERIAL% matched with %VolSerialNum% at %VHDDiscoveredLetter%
echo(
echo However, expected label "%BAKUPVHDLABEL%" does not match "%VolLabel%"
echo(
echo Would you like to Continue regardless, or Abort?
CHOICE /C CA /M "'C'ontinue, or 'A'bort?"
IF ERRORLEVEL 2 GOTO Fail
IF ERRORLEVEL 1 GOTO VHDLabelMatchOverride
GOTO Fail
:VHDLabelMatchOverride
echo(
:VHDLabelMatchToo
REM We found the VHD! It's at %VHDDiscoveredLetter% and we got a SERIAL and LABEL match on that location!
echo %DRIVETOBAK%\%SOURCELABEL% Backup - Step 2/4: VHD mounted at %VHDDiscoveredLetter%!
echo(
REM ===============================================================================
REM End ptTwoB! VHD is mounted AND we know where it is for certain :)
REM End ptTwo! Ready for ROBOCOPY!

:ptThreeA
REM Was last backup successful? Did it FAIL?
REM Check if a "last backup log" exists, and parse it to find any FAILED FILES
IF NOT EXIST "%BAKUPLOGPATH%\last-backup-%DRIVETOBBLETTER%.log" (
	REM NOTIFY user:
	echo Previous backup log "last-backup-%DRIVETOBBLETTER%.log" could not be found. Is this the first
	echo backup ever done?
	echo.
	GOTO NoLastBackupLog
	REM :NoLastBackupLog is after :ptThreeB, as that section also requires last-backup-Var.log
)
GOTO SkipRobocopyOutput
REM ROBOCOPY start text: 
:-------------------------------------------------------------------------------
-------------------------------------------------------------------------------
   ROBOCOPY     ::     Robust File Copy for Windows                              
-------------------------------------------------------------------------------

  Started : Fri Oct 31 15:11:20 2014

   Source : D:\
     Dest : R:\

    Files : *.*
	    
Exc Files : Backup-PC-HomeComp3-18-11.TBI
	    
 Exc Dirs : D:\System Volume Information
	    
  Options : *.* /TEE /S /E /COPYALL /PURGE /MIR /B /ETA /XO /R:10 /W:30 

------------------------------------------------------------------------------

	                   7	D:\
:-------------------------------------------------------------------------------
REM ROBOCOPY output text (end):
:-------------------------------------------------------------------------------
	                   1	D:\Virtual Hard Disk XP\

------------------------------------------------------------------------------

               Total    Copied   Skipped  Mismatch    FAILED    Extras
    Dirs :      7791         1      7790         0         0         0
   Files :    146833         2    146831         0         0         0
   Bytes : 848.175 g   3.660 g 844.514 g         0         0         0
   Times :   0:02:26   0:00:36                       0:00:00   0:01:49


   Speed :           106742705 Bytes/sec.
   Speed :            6107.866 MegaBytes/min.

   Ended : Fri Oct 24 15:13:48 2014

:-------------------------------------------------------------------------------
:SkipRobocopyOutput
IF EXIST "%BAKUPLOGPATH%\%Year%-%DayMonth%-%MonthDay%-%DRIVETOBBLETTER%-PREVBACKUP-Results.csv" del "%BAKUPLOGPATH%\%Year%-%DayMonth%-%MonthDay%-%DRIVETOBBLETTER%-PREVBACKUP-Results.csv"
IF EXIST "%Temp%\RoboParseOutput.log" del "%Temp%\RoboParseOutput.log"
REM Using Parse-RobocopyLogs.ps1 requires PowerShell v?
PowerShell . '%~dp0\Parse-RobocopyLogs.ps1' -fp '%BAKUPLOGPATH%\last-backup-%DRIVETOBBLETTER%.log' -outputfile '%BAKUPLOGPATH%\%Year%-%DayMonth%-%MonthDay%-%DRIVETOBBLETTER%-PREVBACKUP-Results' > "%Temp%\RoboParseOutput.log"

REM Using Parse-RobocopyLogs.ps1 creates *.csv files with some columns empty/missing.
REM But the format of the *.csv it creates leaves no space or tab in-between empty values, creating blocks of consecutive commas (,,,,,) which don't parse well with FOR /F
CD /D "%BAKUPLOGPATH%"
PowerShell (Get-Content %Year%-%DayMonth%-%MonthDay%-%DRIVETOBBLETTER%-PREVBACKUP-Results.csv) ^| ForEach-Object { $_ -replace ',,', ', ,' } ^| Set-Content %Year%-%DayMonth%-%MonthDay%-%DRIVETOBBLETTER%-PREVBACKUP-Results_FIX.csv
REM Now since will skip to the next block of characters after a replace and not re-analyze what it just replaced, there will still be some consecutive commas (i.e. ,, ,, ,,)
REM So, it just needs to be run a second time! This works out perfectly for our situation.
del "%Year%-%DayMonth%-%MonthDay%-%DRIVETOBBLETTER%-PREVBACKUP-Results.csv"
PowerShell (Get-Content %Year%-%DayMonth%-%MonthDay%-%DRIVETOBBLETTER%-PREVBACKUP-Results_FIX.csv) ^| ForEach-Object { $_ -replace ',,', ', ,' } ^| Set-Content %Year%-%DayMonth%-%MonthDay%-%DRIVETOBBLETTER%-PREVBACKUP-Results.csv
del "%Year%-%DayMonth%-%MonthDay%-%DRIVETOBBLETTER%-PREVBACKUP-Results_FIX.csv"

REM Now that we got the stats from the last backup in a .CSV file, it'll be super easy to parse to find out if anything went wrong with the last backup operation.
FOR /F "tokens=17 delims=," %%G IN (%Year%-%DayMonth%-%MonthDay%-%DRIVETOBBLETTER%-PREVBACKUP-Results.csv) DO (set CorrectColumn=%%G&GOTO ESCAPEGOBLINS)
:ESCAPEGOBLINS
set LastFailedFiles=0
FOR /F "skip=1 tokens=17 delims=," %%G IN (%Year%-%DayMonth%-%MonthDay%-%DRIVETOBBLETTER%-PREVBACKUP-Results.csv) DO (set LastFailedFiles=%%G)
CD /D "%~dp0"
REM we have three results here:
REM first, if parse fails the "Files Failed" column name test. That means some *SRS* error with how our powershell script Parse-RobocopyLogs.ps1 produces its .csv output, or how our FOR /F command parses THAT .csv output is going on.
REM next, say that passes but our "LastFailedFiles" does not equal zero. This means two things, either there actually was a failed file in the last robocopy log, or (more likely) the blank spaces in the .csv output of Parse-RobocopyLogs.ps1 changed and now our FOR /F command is picking up a different column than we intended.
REM last, say both those checks pass. That means the last backup completed without failures! Continue with no error message.

REM So, there were some files that failed in the last backup! However:
REM 	- The last Robocopy completed well enough to generate a log
REM 	- The log was complete enough that it could be parsed by Parse-Robocopylogs.ps1
REM 	- Even if files failed, does that mean we now want to abort this new backup operation which is in progress? Personally I think that means it's all the more important to get a new backup operation done ASAP, place a special priority on requiring user intervention, and
IF "%LastFailedFiles%"=="" set LastFailedFiles=0
IF NOT %LastFailedFiles% EQU 0 (
	REM NOTIFY user:
	echo Last backup operation had %LastFailedFiles% files fail to copy.
	echo.
)
IF NOT "%CorrectColumn%"=="Files Failed" GOTO FunkyFailRoboParse
GOTO LastBackupSuccessful
:FunkyFailRoboParse
REM The column name check failed. The log that Parse-RobocopyLogs.ps1 is trying to parse may be incomplete or be formatted differently (different version of Robocopy) than what Parse-RobocopyLogs.ps1 expects.
REM If the last backup failed, I dunno what the user should do, besides try again. 
REM NOTIFY user:
echo Last backup operation may have had serious failure or was interrupted: 
echo(
type "%Temp%\RoboParseOutput.log"
echo(
:LastBackupSuccessful
del "%Temp%\RoboParseOutput.log"
REM Found last backup log, parse'd log, found # of files failed (supposedly), we ready to proceed!
REM -------------------------------------------------------------------------------
REM End ptThreeA! Informed user of any failed files found last backup!

:ptThreeB
REM Check date of last backup.
CD /D "%BAKUPLOGPATH%"
REM Token 5 is start date/time, Token 6 is end date/time
FOR /F "skip=1 tokens=5 delims=," %%G IN (%Year%-%DayMonth%-%MonthDay%-%DRIVETOBBLETTER%-PREVBACKUP-Results.csv) DO (set LastBackupDate=%%G)
CD /D "%~dp0"
FOR /F "tokens=1-3" %%G IN ("%LastBackupDate%") DO (
	set LastBackupDay=%%G
	set LastBackupMonth=%%H
	set LastBackupYear=%%I
)
IF "%LastBackupMonth%"=="Jan" set LastBackupMonth=01
IF "%LastBackupMonth%"=="Feb" set LastBackupMonth=02
IF "%LastBackupMonth%"=="Mar" set LastBackupMonth=03
IF "%LastBackupMonth%"=="Apr" set LastBackupMonth=04
IF "%LastBackupMonth%"=="May" set LastBackupMonth=05
IF "%LastBackupMonth%"=="Jun" set LastBackupMonth=06
IF "%LastBackupMonth%"=="Jul" set LastBackupMonth=07
IF "%LastBackupMonth%"=="Aug" set LastBackupMonth=08
IF "%LastBackupMonth%"=="Sep" set LastBackupMonth=09
IF "%LastBackupMonth%"=="Oct" set LastBackupMonth=10
IF "%LastBackupMonth%"=="Nov" set LastBackupMonth=11
IF "%LastBackupMonth%"=="Dec" set LastBackupMonth=12

CALL DateMath %Year% %DayMonth% %MonthDay% - %LastBackupYear% %LastBackupMonth% %LastBackupDay% > "%Temp%\DateMathOutput.log"
del "%Temp%\DateMathOutput.log"
REM NULL Check :: If var is zero = Success. No hits.
REM NULL Check :: If var is null = Success. Works. All NULL triggers hit successfully.
IF NOT DEFINED DATECUTOFF set DATECUTOFF=0
REM Zero Check :: If var is zero = Success. IF evaluation works, hit.
REM Zero Check :: If var is null = FAILURE. Cmd crash.
set WARNDATEOLD=no
IF "%_dd_int%"=="" set "_dd_int=0"
IF NOT %DATECUTOFF% EQU 0 (
	REM DateCutoff is Defined!
	IF %_dd_int% GTR %DATECUTOFF% set WARNDATEOLD=yes
)
set DaysSinceLastBackup=%_dd_int%
del "%BAKUPLOGPATH%\%Year%-%DayMonth%-%MonthDay%-%DRIVETOBBLETTER%-PREVBACKUP-Results.csv"
:NoLastBackupLog
REM -------------------------------------------------------------------------------
REM End ptThreeB! Days since last backup evaluated, ready for warning message after size test.

:ptThreeC
REM Perform test copy. Find expected size of copy. Is size of changes above limit?
REM These are settings/switches only used for the TEST robocopy. They will change again when doing actual backup copy operation, later on. Do not modify these individual switches unless you know what you are doing.
set "SHHHHTESTING="
set "SHHHHTESTING=/l "
set "PRINTTOSCREEN=/eta "
set "PRINTTOSCREEN="
set "NOPROGRESS="
set "NOPROGRESS=/NP "
set "RETRYSETTINS=/r:1 /w:2 "
set "RETRYSETTINS=/r:0 /w:0 "

REM If you would like to change ROBOCOPY options/switches, below here is the place to do it.
::
::
:: Define the options/switches you want here (will apply to real copy as well):
set "ROBOSWITCHES=/mir /copyall /B /XF pagefile.sys hiberfil.sys Backup-PC-HomeComp3-18-11.TBI /XD "%DRIVETOBAK%\System Volume Information^" %DRIVETOBAK%\$Recycle.Bin ^/XJ ^/XO ^/tee"
::
::
::

REM This first ROBOCOPY is a test... (DO NOT MODIFY BELOW THIS LINE)
IF EXIST "%BAKUPLOGPATH%\TEST-COPY-%DRIVETOBBLETTER%.log" del "%BAKUPLOGPATH%\TEST-COPY-%DRIVETOBBLETTER%.log"
IF EXIST "%BAKUPLOGPATH%\RoboParseOutput.log" del "%BAKUPLOGPATH%\RoboParseOutput.log"
REM If SizeCutoff is not set, escape here after the important VARS here are set.
IF NOT DEFINED SIZECUTOFF set SIZECUTOFF=0
IF %SIZECUTOFF% EQU 0 GOTO NoSizeLimit
REM NOTIFY user:
echo Measuring size of changes to copy...
echo(

START "Measuring size of difference..." /W "robocopy" %DRIVETOBAK%\ %VHDDiscoveredLetter%\ *.* %ROBOSWITCHES% %NOPROGRESS%%RETRYSETTINS%%PRINTTOSCREEN%%SHHHHTESTING%/log:"%BAKUPLOGPATH%\TEST-COPY-%DRIVETOBBLETTER%.log"

REM ROBOCOPY Options/Switches Descriptions: (NOTE THIS LIST IS NOT COMPLETE)
:: robocopy "<source>" "<destination>"
:: e.g. ROBOCOPY C:\source D:\dest *.* /XD "C:\System Volume Information" /XF pagefile.sys hiberfil.sys
:: /S - Copies subdirectories. Note that this option excludes empty directories.
:: /E - Copies subdirectories. Note that this option includes empty directories. For additional information, see Remarks.
:: /copy:<CopyFlags>
:: Specifies the file properties to be copied. The following are the valid values for this option:
:: D - Data
:: A - Attributes
:: T - Time stamps
:: S - NTFS access control list (ACL)
:: O - Owner information
:: U - Auditing information
:: The default value for CopyFlags is DAT (data, attributes, and time stamps).
:: /copyall - Copies all file information (equivalent to /copy:DATSOU).
:: /DCOPY:T - COPY Directory Timestamps.
:: /MIR - Mirrors a directory tree (equivalent to /e plus /purge), deletes any destination files. Note that when used with /Z (or MAYBE even /XO) it does not delete already copied files at the destination (useful for resuming a copy)
:: /PURGE - delete dest files/dirs that no longer exist in source.
:: /L - Specifies that files are to be listed only (and not copied, deleted, or time stamped).
:: /NP - No Progress – don’t display % copied.
:: /ETA - Shows the estimated time of arrival (ETA) of the copied files.
:: /BYTES - Print sizes as bytes.
:: /log:<LogFile> - Writes the status output to the log file (overwrites the existing log file).
:: /log+:<LogFile> - Writes the status output to the log file (appends the output to the existing log file).
:: /TEE - Output to console window, as well as the log file.
:: /IS - Include Same, overwrite files even if they are already the same.
:: /IT - Include Tweaked files.
:: /X - Report all eXtra files, not just those selected & copied.
:: /FFT - uses fat file timing instead of NTFS. This means the granularity is a bit less precise. For across-network share operations this seems to be much more reliable - just don't rely on the file timings to be completely precise to the second.
:: /Z - ensures Robocopy can resume the transfer of a large file in mid-file instead of restarting. (Restart Mode)(maybe for Network Copys)
:: /B - copies in Backup Modes (overrides ACLs for files it doesn't have access to so it can copy them. Requires User-Level or Admin permissions)
:: /ZB : Use restartable mode; if access denied use Backup mode.
:: /R:n - Number of Retries on failed copies - default is 1 million.
:: /W:n - Wait time between retries - default is 30 seconds.
:: /REG - Save /R:n and /W:n in the Registry as default settings.
:: /XO - Excludes older files. (Only copies newer and changed files)
:: /XJ - eXclude Junction points. (normally included by default). In Windows 7 Junction Points were introduced which adds symbolic-like-links for "Documents and Settings" which redirect to "C:\User\Documents" for old program compatibility. Sometimes they can throw ROBOCOPY into a loop and it will copy the same files more than once.
:: /XF <FileName>[ ...] Excludes files that match the specified names or paths. Note that FileName can include wildcard characters (* and ?).
:: /XD <Directory>[ ...] Excludes directories that match the specified names and paths.
:: XF and XD can be used in combination  e.g. ROBOCOPY c:\source d:\dest /XF *.doc *.xls /XD c:\unwanted /S
REM - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
REM Check ROBOCOPY test log to find size of changes before performing copy.
IF EXIST "%BAKUPLOGPATH%\RoboParseOutput.log" del "%BAKUPLOGPATH%\RoboParseOutput.log"
PowerShell . '%~dp0\Parse-RobocopyLogs.ps1' -fp '%BAKUPLOGPATH%\TEST-COPY-%DRIVETOBBLETTER%.log' -outputfile '%BAKUPLOGPATH%\%Year%-%DayMonth%-%MonthDay%-%DRIVETOBBLETTER%-TESTCOPY-Results' > "%BAKUPLOGPATH%\RoboParseOutput.log"

REM Add spaces between commas for blank column entries
CD /D "%BAKUPLOGPATH%"
PowerShell (Get-Content %Year%-%DayMonth%-%MonthDay%-%DRIVETOBBLETTER%-TESTCOPY-Results.csv) ^| ForEach-Object { $_ -replace ',,', ', ,' } ^| Set-Content %Year%-%DayMonth%-%MonthDay%-%DRIVETOBBLETTER%-TESTCOPY-Results_FIX.csv
del "%Year%-%DayMonth%-%MonthDay%-%DRIVETOBBLETTER%-TESTCOPY-Results.csv"
PowerShell (Get-Content %Year%-%DayMonth%-%MonthDay%-%DRIVETOBBLETTER%-TESTCOPY-Results_FIX.csv) ^| ForEach-Object { $_ -replace ',,', ', ,' } ^| Set-Content %Year%-%DayMonth%-%MonthDay%-%DRIVETOBBLETTER%-TESTCOPY-Results.csv
del "%Year%-%DayMonth%-%MonthDay%-%DRIVETOBBLETTER%-TESTCOPY-Results_FIX.csv"
REM Grab Bytes Copied to find expect size of copy
FOR /F "tokens=20 delims=," %%G IN (%Year%-%DayMonth%-%MonthDay%-%DRIVETOBBLETTER%-TESTCOPY-Results.csv) DO (set CorrectBTColumn=%%G&GOTO ESCAPEWEREWIMINS)
:ESCAPEWEREWIMINS
FOR /F "skip=1 tokens=20 delims=," %%G IN (%Year%-%DayMonth%-%MonthDay%-%DRIVETOBBLETTER%-TESTCOPY-Results.csv) DO (set BytesCopiedTest=%%G)
CD /D "%~dp0"
set SizeCopied=%BytesCopiedTest:~0,-2%
set SizeUnits=%BytesCopiedTest:~-1%
:: CorrectBTColumn="Bytes Copied"
:: BytesCopiedTest="1.2 t"
:: BytesCopiedTest="369.194 g"
:: BytesCopiedTest="16.78 m"
:: BytesCopiedTest="1.2 k"
set WARNSIZE=no
IF "%BytesCopiedTest%"=="0" set "SizeCopied=0"
IF "%SizeCopied%"=="" GOTO NoSizeLimit
IF /I NOT "%SizeUnits%"=="g" GOTO NoSizeLimit
IF %SizeCopied% EQU 0 GOTO NoSizeLimit
IF /I "%SizeUnits%"=="g" (
	IF %SizeCopied% GTR %SIZECUTOFF% (
		set WARNSIZE=yes
		SET /A "SizeDiff=SizeCopied-SIZECUTOFF"
	)
)
:NoSizeLimit
IF %SIZECUTOFF% EQU 0 set WARNSIZE=no
REM WARNDATEOLD gets set above after the check if DATECUTOFF is active. If it is, it is.
REM Same with WARNSIZE, it only gets set if results from the size check are set and do FAIL.
IF "%WARNDATEOLD%"=="yes" (
	IF "%WARNSIZE%"=="yes" (set SystemsChecks=ZizeAnDate) ELSE (set SystemsChecks=JustDateErr)
) ELSE (
	IF "%WARNSIZE%"=="yes" (set SystemsChecks=SizeErr) ELSE (set SystemsChecks=NoWong)
)
IF "%SystemsChecks%"=="JustDateErr" GOTO NOTIFYDATEOLD
IF "%SystemsChecks%"=="NoWong" GOTO NoTingWong
REM Size [and Date] Warning (comBINED)
echo %DRIVETOBAK%\%SOURCELABEL% Backup - Step 3/4: WARNING = Size of changeset very large.
echo(
echo Size of planned copy is %BytesCopiedTest%, which is %SizeDiff% GB above the warning (%SIZECUTOFF% GB).
echo(
IF "%WARNDATEOLD%"=="yes" echo Additionally, the last backup was performed %DaysSinceLastBackup% days ago.
IF "%WARNDATEOLD%"=="yes" echo.
IF "%WARNDATEOLD%"=="yes" (
	echo This operation can delete files at the destination! It may have been a long
) ELSE (
	echo This operation can DELETE files at the destination! If the size looks odd, make
)
IF "%WARNDATEOLD%"=="yes" (
	echo time since the last backup for %DRIVETOBAK% was done. Make sure the destination ^(%VHDDiscoveredLetter%^) is
) ELSE (
	echo sure the destination ^(%VHDDiscoveredLetter%^) is the intended backup location for the source ^(%DRIVETOBAK%^).
)
IF "%WARNDATEOLD%"=="yes" (
	echo correct.
)
echo(
echo Would you like to Continue regardless, or Abort?
CHOICE /C CA /M "'C'ontinue, or 'A'bort?"
IF ERRORLEVEL 2 GOTO Fail
IF ERRORLEVEL 1 GOTO CopyOK
GOTO Fail
:NOTIFYDATEOLD
REM Just Date Warning
REM NOTIFY user:
echo Last backup completed %DaysSinceLastBackup% days ago.
:CopyOK
echo(
:NoTingWong
REM NOTIFY user:
IF %SIZECUTOFF% EQU 0 (
	echo %DRIVETOBAK%\%SOURCELABEL% Backup - Step 3/4: Backup is working...
) ELSE (
	IF %SizeCopied% EQU 0 (
		echo %DRIVETOBAK%\%SOURCELABEL% Backup - Step 3/4: No changes to copy!
	) ELSE (
		echo %DRIVETOBAK%\%SOURCELABEL% Backup - Step 3/4: Copying %BytesCopiedTest% to backup...
	)
)
echo(
IF EXIST "%BAKUPLOGPATH%\TEST-COPY-%DRIVETOBBLETTER%.log" del "%BAKUPLOGPATH%\TEST-COPY-%DRIVETOBBLETTER%.log"
IF EXIST "%BAKUPLOGPATH%\%Year%-%DayMonth%-%MonthDay%-%DRIVETOBBLETTER%-TESTCOPY-Results.csv" del "%BAKUPLOGPATH%\%Year%-%DayMonth%-%MonthDay%-%DRIVETOBBLETTER%-TESTCOPY-Results.csv"
IF EXIST "%BAKUPLOGPATH%\RoboParseOutput.log" del "%BAKUPLOGPATH%\RoboParseOutput.log"
REM -------------------------------------------------------------------------------
REM End ptThreeC! Test copy performed, size of changes found, and they do not break any rules. 

:ptThreeD
REM Start ROBOCOPY.
REM Step 3: Robocopy all changes from SOURCE to VHD backup
REM ===============================================================================
set "SHHHHTESTING=/l "
set "SHHHHTESTING="
set "PRINTTOSCREEN="
set "PRINTTOSCREEN=/eta "
set "NOPROGRESS=/NP "
set "NOPROGRESS="
set "RETRYSETTINS=/r:0 /w:0 "
set "RETRYSETTINS=/r:1 /w:2 "

REM Here we go...
IF %SizeCopied% EQU 0 GOTO NoChangesToCopy
IF EXIST "%BAKUPLOGPATH%\last-backup-%DRIVETOBBLETTER%.log" del "%BAKUPLOGPATH%\last-backup-%DRIVETOBBLETTER%.log"

START "Copying new and changed files to backup..." /W "robocopy" %DRIVETOBAK%\ %VHDDiscoveredLetter%\ *.* %ROBOSWITCHES% %NOPROGRESS%%RETRYSETTINS%%PRINTTOSCREEN%%SHHHHTESTING%/log:"%BAKUPLOGPATH%\last-backup-%DRIVETOBBLETTER%.log"

COPY "%BAKUPLOGPATH%\last-backup-%DRIVETOBBLETTER%.log" "%BAKUPLOGPATH%\%Year%-%DayMonth%-%MonthDay%-%DRIVETOBBLETTER%-%SOURCELABEL%-BACKUP.log" >nul
:NoChangesToCopy
echo %DRIVETOBAK%\%SOURCELABEL% Backup - Step 3/4: All changes copied to backup!
echo(
REM ===============================================================================
REM ROBOCOPY complete!

:ptThreeE
REM Unmount VHD with DISKPART.
REM Step 4: Unmount VHD, maybe show user log file?
:: diskpart sel vdisk file="F:\shop_drivebox_d_data.VHD"
:: diskpart detach vdisk
IF EXIST "%Temp%\unmount-backup.txt" del "%Temp%\unmount-backup.txt"
IF EXIST "%Temp%\unmount-status.log" del "%Temp%\unmount-status.log"
echo sel vdisk file="%BAKUPIMG%">%Temp%\unmount-backup.txt
echo detach vdisk>>%Temp%\unmount-backup.txt
diskpart /s %Temp%\unmount-backup.txt>%Temp%\unmount-status.log
del "%Temp%\unmount-backup.txt"
REM DISKPART dismount finished!
REM -------------------------------------------------------------------------------
REM Check DISKPART log to make sure dismount was successful!
set "DPlineone="
set "DPlinetwo="
FOR /F "skip=5 tokens=*" %%G IN (%Temp%\unmount-status.log) DO (set DPlineone=%%G&GOTO EscUnmountUno)
:EscUnmountUno
FOR /F "skip=7 tokens=*" %%G IN (%Temp%\unmount-status.log) DO (set DPlinetwo=%%G&GOTO EscUnmountDos)
:EscUnmountDos
REM Successfully Detached:
:: DPlineone="DiskPart successfully selected the virtual disk file."
:: DPlinetwo="DiskPart successfully detached the virtual disk file."
FOR /F "tokens=2,3" %%G IN ("%DPlineone%") DO (
	set DPfirstlinechqu1=%%G
	set DPfirstlinechqu2=%%H
)
IF "%DPfirstlinechqu1%"=="successfully" (set DismountDPLOne=GOOD) ELSE (set DismountDPLOne=BAD)
IF NOT "%DPfirstlinechqu2%"=="selected" set "DismountDPLOne=BAD"
REM - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
FOR /F "tokens=2,3,5,6" %%G IN ("%DPlinetwo%") DO (
	set DPtwolinechqu1=%%G
	set DPtwolinechqu2=%%H
	set DPtwolinechqu3=%%I
	set DPtwolinechqu4=%%J
)
IF "%DPtwolinechqu1%"=="successfully" (
	IF "%DPtwolinechqu2%"=="detached" (set DismountDPsuccess=YES) ELSE (set DismountDPsuccess=NO)
) ELSE (set DismountDPsuccess=NO)
set UnmountVHDStatus=NOT
IF "%DismountDPLOne%"=="GOOD" (
	IF "%DismountDPsuccess%"=="YES" set UnmountVHDStatus=GOOD
)
IF [%UnmountVHDStatus%]==[GOOD] GOTO DismountDone
REM DISKPART VHD unmount failed. At this point we assume everything else completed swimmingly, so if we had trouble dismounting the VHD, let's just NOTIFY the user and carry on.
REM NOTIFY user:
echo Failed to un-mount backup image "%BAKUPIMG%"
echo(
echo DISKPART:
echo %DPlineone%
echo %DPlinetwo%
echo(
:DismountDone
REM -------------------------------------------------------------------------------
REM End ptThreeE! Unmounted VHD with DISKPART, and if there were any errors user will know about it.

:ptFour
REM Last step! Was backup successful? Parse logs to find any copy errors.
REM -------------------------------------------------------------------------------
IF EXIST "%BAKUPLOGPATH%\%Year%-%DayMonth%-%MonthDay%-%DRIVETOBBLETTER%-Results.csv" del "%BAKUPLOGPATH%\%Year%-%DayMonth%-%MonthDay%-%DRIVETOBBLETTER%-Results.csv"
IF EXIST "%Temp%\RoboParseOutput.log" del "%Temp%\RoboParseOutput.log"
REM Using Parse-RobocopyLogs.ps1 requires PowerShell v?
PowerShell . '%~dp0\Parse-RobocopyLogs.ps1' -fp '%BAKUPLOGPATH%\last-backup-%DRIVETOBBLETTER%.log' -outputfile '%BAKUPLOGPATH%\%Year%-%DayMonth%-%MonthDay%-%DRIVETOBBLETTER%-Results' > "%Temp%\RoboParseOutput.log"
REM Using Parse-RobocopyLogs.ps1 creates *.csv files with some columns empty/missing.
REM But the format of the *.csv it creates leaves no space or tab in-between empty values, creating blocks of consecutive commas (,,,,,) which don't parse well with FOR /F
CD /D "%BAKUPLOGPATH%"
PowerShell (Get-Content %Year%-%DayMonth%-%MonthDay%-%DRIVETOBBLETTER%-Results.csv) ^| ForEach-Object { $_ -replace ',,', ', ,' } ^| Set-Content %Year%-%DayMonth%-%MonthDay%-%DRIVETOBBLETTER%-Results_FIX.csv
del "%Year%-%DayMonth%-%MonthDay%-%DRIVETOBBLETTER%-Results.csv"
PowerShell (Get-Content %Year%-%DayMonth%-%MonthDay%-%DRIVETOBBLETTER%-Results_FIX.csv) ^| ForEach-Object { $_ -replace ',,', ', ,' } ^| Set-Content %Year%-%DayMonth%-%MonthDay%-%DRIVETOBBLETTER%-Results.csv
del "%Year%-%DayMonth%-%MonthDay%-%DRIVETOBBLETTER%-Results_FIX.csv"
set "CorrectColumn="
FOR /F "tokens=17 delims=," %%G IN (%Year%-%DayMonth%-%MonthDay%-%DRIVETOBBLETTER%-Results.csv) DO (set CorrectColumn=%%G&GOTO ESCAPEWARRIORS)
:ESCAPEWARRIORS
set FinalFailedFiles=0
FOR /F "skip=1 tokens=17 delims=," %%G IN (%Year%-%DayMonth%-%MonthDay%-%DRIVETOBBLETTER%-Results.csv) DO (set FinalFailedFiles=%%G)
IF "%FinalFailedFiles%"=="" set FinalFailedFiles=0
set "SizeCopied="
set "SizeUnits="
FOR /F "skip=1 tokens=20 delims=," %%G IN (%Year%-%DayMonth%-%MonthDay%-%DRIVETOBBLETTER%-Results.csv) DO (set RealBytesCopied=%%G)
CD /D "%~dp0"
set SizeCopied=%RealBytesCopied:~0,-2%
set SizeUnits=%RealBytesCopied:~-1%
REM Time for final error messages!!!
REM We've got some options here, but at this point every message will let the script run to it's end
REM First we already dismounted the VHD, so we can make sure than went off alright.
REM Next we've got any errors parse'd from this most recent backup, in two parts: the Column Name check and the actual Failed Files.
REM While it's true that if the Column Name check failed, there's probably a serious error with the parse.
:: 1- Everything went spectacularly. No errors.
:: 2- VHD unmounted fine, but parse errors. Either there were failed files, or PARSE FAILED
:: 3- VHD dismount failed, no parse errors. NOTIFY user to unmount VHD manually
:: 4- VHD dismount failed AND parse errors. 
IF NOT "%CorrectColumn%"=="Files Failed" GOTO StrangeRoboParse
GOTO RealBackupSuccessful
:StrangeRoboParse
REM NOTIFY user:
echo Backup operation may have had serious failure or was interrupted: 
echo(
type "%Temp%\RoboParseOutput.log"
echo(
:RealBackupSuccessful
IF %FinalFailedFiles% EQU 0 (
	IF [%UnmountVHDStatus%]==[GOOD] (
		REM Unmount successful and no files were failed!
		echo %DRIVETOBAK%\%SOURCELABEL% Backup - Step 4/4: Backup completed and VHD unmounted!
		IF /I "%SizeUnits%"=="g" echo.
		IF /I "%SizeUnits%"=="g" echo %RealBytesCopied% copied.
	) ELSE (
		REM No files failed to copy, but VHD dismount was unsuccessful.
		echo %DRIVETOBAK%\%SOURCELABEL% Backup - Step 4/4: Backup completed successfully!
		IF /I "%SizeUnits%"=="g" echo.
		IF /I "%SizeUnits%"=="g" echo %RealBytesCopied% copied.
		echo.
		echo VHD could not be un-mounted after backup completed, dismount manually.
	)
) ELSE (
	IF [%UnmountVHDStatus%]==[GOOD] (
		REM Some files failed, but unmount was successful!
		echo %DRIVETOBAK%\%SOURCELABEL% Backup - Step 4/4: %FinalFailedFiles% files failed to copy.
		IF /I "%SizeUnits%"=="g" echo.
		IF /I "%SizeUnits%"=="g" echo %RealBytesCopied% copied successfully.
	) ELSE (
		REM Some files failed (or parse went bad) and dismount failed.
		echo %DRIVETOBAK%\%SOURCELABEL% Backup - Step 4/4: %FinalFailedFiles% files failed to copy.
		IF /I "%SizeUnits%"=="g" echo.
		IF /I "%SizeUnits%"=="g" echo %RealBytesCopied% copied successfully.
		echo.
		echo VHD could not be un-mounted after backup completed, dismount manually.
	)
)
REM -------------------------------------------------------------------------------
REM End ptFour! That's it folks!
GOTO End

:FreshBackupImage
cls
echo ===============================================================================
echo =                             New Backup Image                                =
echo - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - - -
echo(
echo If you cannot find the backup image for drive %DRIVETOBAK%\ "%SOURCELABEL%", it will have to be 
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
REM Get time to show user when script finished (so they can see on average how long it takes)
FOR %%G IN (%Date%) DO SET EndedToday=%%G
SET EndedNow=%Time%
echo(
echo %DRIVETOBAK%\%SOURCELABEL% Backup - completed on %EndedToday% at %EndedNow%
echo(
REM -------------------------------------------------------------------------------
REM Show any error logs from ptThreeE or ptFour if we found 'em
IF EXIST "%Temp%\RoboParseOutput.log" del "%Temp%\RoboParseOutput.log"
IF NOT [%UnmountVHDStatus%]==[GOOD] "%Temp%\unmount-status.log"
IF EXIST "%Temp%\unmount-status.log" del "%Temp%\unmount-status.log"
IF "%CorrectColumn%"=="Files Failed" (
	IF %FinalFailedFiles% EQU 0 (
		REM delete this one, unless there were any errors with parse
		del "%BAKUPLOGPATH%\%Year%-%DayMonth%-%MonthDay%-%DRIVETOBBLETTER%-Results.csv"
	)
)
REM -------------------------------------------------------------------------------
REM Of course, always show the log afterwards no matter what!
"%BAKUPLOGPATH%\last-backup-%DRIVETOBBLETTER%.log"
IF NOT %FinalFailedFiles% EQU 0 (
	REM Some files failed! Show explorer window of logs no matter what if files failed.
	%SystemRoot%\explorer.exe "%BAKUPLOGPATH%"
) ELSE (
	IF NOT "%CorrectColumn%"=="Files Failed" (
		REM Files did not fail, but parse did? show log location!(Open location for .csv parse)
		%SystemRoot%\explorer.exe "%BAKUPLOGPATH%"
	)
)
pause
exit

:Fail
echo(
echo The script could not continue.
echo(
pause
exit