# Scheduled-Backup
A batch script designed to copy source drive changes to a VHD file located on a different drive. Designed to be fully automated so it can be set as a Scheduled Task and complete without user intervention. A poor man's backup plan.

This script tries to use native Windows tools to avoid installing any additional software on the target machine. This means using CD and VOL commands to check if all needed drives are online, DISKPART to mount the VHD file, and ROBOCOPY to discover the differences between the two drives and copy over only the changes. DISKPART then dismounts the VHD.

It's not a true Incremental Backup but it acts like one. You won't have a "history" of all backups like a true Incremental Backup solution would, but you will have an exact copy of your SOURCE drive (current to date of last backup operation) convienintly packaged in a VHD, ready to burn to another drive if the source fails.

But it works. It's simple. If you're on Windows, don't use the Windows backup tool, and want a quick backup solution you can set and forget, this is it.

## v0.1 IN PRODUCTION
Yes, this is a functional script that has been running in production as a backup solution for one of our drive machines in the shop.

It's short, basic, dirty, has hard-coded variables, but it gets the job done and has been running at about 95% success rate for the past 14 months.

The script fails if it can't find all of the drives. The machine we use this on gets shut down and booted up several times a day, with drives to be imaged connected and disconnected regularly. Some (hard-coded) needed drive letters may be occupied or the drive may be disconnected completely, so the script blindly continues through the failures.

v1.0 which I'm working on now, will detect if the drives it's looking at are the right ones, search for the right ones if they're not, and either warn the user before continuing through a potentially risky operation (changes to copy >*n*GB) or abort itself if it detects a problem.
