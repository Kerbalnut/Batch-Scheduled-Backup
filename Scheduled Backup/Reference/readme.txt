http://nicj.net/mounting-vhds-in-windows-7-from-a-command-line-script/

These commands work fine on an ad-hoc basis, but I had the need to automate loading a VHD from a script.  Luckily, diskpart takes a single parameter, /s, which specifies a diskpart “script”.  The script is simply the command you would have typed in above:

C:\> diskpart /s [diskpart script file]

I’ve created two simple scripts, MountVHD.cmd and UnmountVHD.cmd that create a “diskpart script”, run it, then remove the temporary file.  This way, you can simply run MountVHD.cmd and point it to your VHD:

C:\> MountVHD.cmd [location of vhd] [drive letter - optional]

Or unmount the same VHD:

C:\> UnMountVHD.cmd [location of vhd]

