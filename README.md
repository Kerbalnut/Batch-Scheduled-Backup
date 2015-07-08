# Scheduled-Backup.bat
A batch script designed to copy only differences (new and changed files) from a SOURCE drive to a VHD file located on a different DESTINATION disk. Designed to be fully automated, so it can be set to repeat as a Scheduled Task and can complete without user intervention. A poor man's backup plan.

## USE-CASE (raison d'être):
A workplace machine used to image hard drives requires regular backup. But because this machine is used to image new hard drives, the DESTINATION disk containing the backup image VHD may get disconnected, or the VHD may assume a different driver letter than is normal when re-mounted. This script is designed to detect these changes and work around them or ABORT itself if it cannot.

## Overview:
This script uses native Windows tools to avoid having to install additional software on the target machine. It uses CD /D and VOL commands to test if the necessary drives are online, DISKPART to mount the VHD file, and ROBOCOPY to discover the differences between the two drives and copy over only the changes. DISKPART then dismounts the VHD. POWERSHELL is used to parse the ROBOCOPY logs and determine if the last backup was successful, if any files failed, and to test the size of changes about to be copied and WARN the user if it detects a risky operation.

Since this backup operation can delete files in the destination VHD (ROBOCOPY /MIR), extra checks are put into place to make sure all targets are correct. If the SOURCE disk, DESTINATION disk, or VHD fail any tests this script ABORTS the operation with a failure.

It's a given in this situation that other people will be using the computer this script is scheduled to run on. To reconcile this, custom error messages are written for nearly every known possibility, designed to describe in the simplest way necessary what went wrong and in most cases provide an easy SOLUTION for the user to fix the problem.

## Special Thanks:
To Simon Sheppard for the superb resource that is SS64.com, as well as [DateMath.cmd](http://ss64.com/nt/syntax-datemath.html)!

To Guy Chapman for [Parse-RobocopyLogs.ps1](http://www.chapmancentral.co.uk/cloudy/2013/02/23/parsing-robocopy-logs-in-powershell/)!

To all the friends and great help over at [StackExchange](http://stackexchange.com/)!

# v1.0 SHIPPED

### System Requirements:
 - Windows 7 and up (untested on Vista or earlier)
 - PowerShell v?.0 (only tested with v4)

### Dependencies: 
(These are included, but must stay in the same folder as the script)

 - `Parse-RobocopyLogs.ps1`
 - `DateMath.cmd`

### Non-goals: 
This version will *NOT* support the following features:

 - Create a *.VHD for the user. This script assumes you've already made a VHD for your backup, or are capable of creating one manually.
 - Looking for the *.VHD file anywhere besides the location it’s supposed to be at (DESTINATION disk must keep same letter and VHD must stay in same location always. Make user fix drive letters or update location.)
 - Create new *.VHD file for backup if it goes missing (maybe in the future, but this is an intensive process requiring a lot of work to implement that can be abused by ignorant users if automated)
 - Warn user about size of copy for anything less than a gig. We assume here that the user will only be concerned if the operation will copy more than just a few gigs. SizeCutoff will only be measured in GB. (Options to check for anything less will not be implemented)
	
### THINGS TO (POSSIBLY) ADD LATER: (NO PROMISES)
 - Create new VHD if the listed one cannot be found (and do fresh ROBOCOPY backup)
 - If LABEL is detected as changed (drive LETTER and SERIAL still the same) give option to auto-rename it back.

### Instructions:
 1. Copy folder `Scheduled Backup` to your hard drive.
 2. Edit `Scheduled-Backup.bat`
 3. Read instructions and customize variables to your application.
 4. Make sure proper PowerShell version is installed (to check PS version run `$PSVersionTable.PSVersion`)
 5. `Get-ExecutionPolicy` and make sure it's not Restricted, if so `Set-ExecutionPolicy` to something like `RemoteSigned`.
 6. Test run.

### Bug Reports:
To file a bug report, head over to the "Issues" section for this project here on GitHub create a new issue. Fill out the following:

 1. Steps to reproduce the bug.
 2. What you expected to happen, and
 3. What actually happened.

Disclaimer: There are absolutely no garauntees on a timeline for any bug to get fixed, or that they will ever get fixed at all. This is a free public project now, meaning I don't get paid for this and you are more than welcome to clone this repository, find & fix the bug yourself, and commit your changes back to master here. All that being said, I will do my best to continually update this as I find bugs, and test & fix bugs others find if you are kind enough to post them here. :)

# v0.1 IN PRODUCTION
Yes, this is a functional script that has been running in production as a backup solution for one of our drive machines in the shop.

It's short, basic, dirty, has hard-coded variables, but it gets the job done and has been running at about 95% success rate for the past 14 months.

The script fails if it can't find all of the drives. The machine we use this on gets shut down and booted up several times a day, with drives to be imaged connected and disconnected regularly. Some (hard-coded) needed drive letters may be occupied or the drive may be disconnected completely, so the script blindly continues through the failures.
