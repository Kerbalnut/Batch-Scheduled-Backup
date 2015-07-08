# Scheduled-Backup.bat
A batch script designed to copy only the differences (new and changed files) from a SOURCE drive to a VHD file located on a different DESTINATION disk. Designed to be fully automated so it can be set to repeat as a Scheduled Task and can complete without user intervention. A poor man's backup plan.

## USE-CASE:
A workplace machine used to image hard drives requires a good backup of the drive images it contains. But because this machine is used regularly to image new hard drives, it means the DESTINATION disk containing the backup image VHD may become disconnected, or the VHD may assume a different driver letter than is normal when mounted. This script is designed to detect these changes, and work around them or ABORT itself if it cannot.

## Overview:
This script uses native Windows tools to avoid having to install additional software on the target machine. It uses CD /D and VOL commands to test if the necessary drives are online, DISKPART to mount the VHD file, and ROBOCOPY to discover the differences between the two drives and copy over only the changes. DISKPART then dismounts the VHD. POWERSHELL is used to parse the ROBOCOPY logs and determine if the last backup was successful, if any files failed, and to test the size of changes about to be copied and WARN the user if it detects a risky operation.

Since this backup operation can delete files in the destination VHD (ROBOCOPY /MIR), extra checks are put into place to make sure all targets are correct. If the SOURCE disk, DESTINATION disk, or VHD fail any tests this script ABORTS the operation with a failure.

It's a given in this situation that other people will be using the computer this is being run on. To reconcile this, custom error messages are written for nearly every known possibility, designed to describe in the simplest way necessary what went wrong and in most cases provide an easy SOLUTION for the user to fix the problem.

It works. It's simple. If you're on Windows, don't use the Windows backup tool, and want a quick backup solution you can set up and forget, this is it.

# v1.0 SHIPPED

### System Requirements:
 - Windows 7 and up (untested on Vista or earlier)
 - PowerShell v?.0

### Dependencies: (These must stay in the same folder as Scheduled-Backup.bat)
 - Parse-RobocopyLogs.ps1
 - DateMath.cmd

### Non-goals: 
This version will *NOT* support the following features:
	- Looking for the *.VHD file anywhere besides the location it’s supposed to be at (DESTINATION disk must keep same letter and VHD must stay in same location always. Make user fix drive letters or update location)
	- Create new *.VHD file for backup if it goes missing (maybe in the future, but this is an intensive process requiring a lot of work to implement that can be abused by ignorant users if automated)
	- Warn user about size of copy for anything less than a gig. We assume here that the user will only be concerned if the operation will copy more than just a few gigs. SizeCutoff will only be measured in GB. (Options to check for anything less will not be implemented)
	
### THINGS (POSSIBLY) TO ADD LATER: (NO PROMISES)
	- Create new VHD if the listed one cannot be found (and do fresh ROBOCOPY backup)
	- If LABEL is detected as changed (drive LETTER and SERIAL still the same) give option to auto-rename it back.

# v0.1 IN PRODUCTION
Yes, this is a functional script that has been running in production as a backup solution for one of our drive machines in the shop.

It's short, basic, dirty, has hard-coded variables, but it gets the job done and has been running at about 95% success rate for the past 14 months.

The script fails if it can't find all of the drives. The machine we use this on gets shut down and booted up several times a day, with drives to be imaged connected and disconnected regularly. Some (hard-coded) needed drive letters may be occupied or the drive may be disconnected completely, so the script blindly continues through the failures.
