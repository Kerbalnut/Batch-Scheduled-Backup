@echo off
REM This script has 4 parts, and must have someone click "Yes" to the UAC Administer rights prompt (if we don't have admin rights already)
REM Step 1: Get UAC Admin Rights
REM Step 2: Mount VHD backup
REM Step 3: Robocopy all changes from D:\ Data to VHD backup
REM Step 4: Unmount VHD
echo D:\Data Backup - script started!
echo(

:: ------- Start Script ------- 
:Step1
REM Step 1: Get UAC Admin Rights

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
echo D:\Data Backup - Step 1/4: Admin permissions obtained!


:Step2
REM Step 2: Mount VHD backup

diskpart /s mount-d-backup-vhd.diskpart
echo(
echo D:\Data Backup - Step 2/4: VHD mounted!


:Step3
REM Step 3: Robocopy all changes from D:\ Data to VHD backup

REM echo This backup is a test...
REM robocopy D:\ R:\ *.* /e /copyall /b /xo /XF Backup-PC-HomeComp3-18-11.TBI pagefile.sys /XD "D:\System Volume Information"  /mir /l /r:10 /eta /tee /log:"C:\Users\Shop\Documents\+Reference\Scripts\Robocopy\Scheduled Backup\Logs\Data-D-backup-testlog.txt"

robocopy D:\ R:\ *.* /e /copyall /b /xo /XF Backup-PC-HomeComp3-18-11.TBI pagefile.sys /XD "D:\System Volume Information" /mir /r:10 /eta /tee /log+:"C:\Users\Shop\Documents\+Reference\Scripts\Robocopy\Scheduled Backup\Logs\Data-D-backup-log.txt"

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
echo D:\Data Backup - Step 3/4: Robocopy'd all changes from D:\ Data to VHD backup!

:Step4
REM Step 4: Unmount VHD

diskpart /s unmount-d-backup-vhd.diskpart
echo(
echo D:\Data Backup - Step 4/4: VHD Unmounted!


:End
echo(
echo D:\Data Backup - script completed!
pause
exit