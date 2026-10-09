Remote Server Cleanup
Windows PowerShell tool for cleaning selected junk off one or more remote servers. It uses PowerShell remoting, shows the run in a Process Viewer, and writes a log and a CSV on each target.

v1.0.0 is the first published baseline. Later changes belong in later commits and tags, not in a rewrite of this release.

Platform
Windows PowerShell 5.1
Windows Server 2016 or newer
WinRM enabled on the target
An account that can run the cleanup elevated over remoting
Run it from an elevated PowerShell session on your machine. If a target does not answer WinRM, the usual fix on that server is Enable-PSRemoting -Force.

What it can clean
User temp for the remoting account, not every profile on the box
Windows temp, plus CBS and DISM logs
Recycle Bin on the selected drive
SCCM / ConfigMgr cache
Windows Update download cache (SoftwareDistribution\Download)
DISM component-store cleanup
IIS logs older than the retention value
Extra paths you supply
Each run also records free space before and after. Output on the target goes to C:\ProgramData\ServerCleanup as a timestamped .log and .csv. Those files include the server name and the paths that were touched.

Before you point it at a server
This tool deletes data. It can also stop wuauserv, bits, and cryptsvc while it clears the Windows Update cache, and DISM can run for a long time.

Start with one server.
Turn WhatIf on for the first pass. WhatIf logs the intended deletes and does not run DISM.
Treat a live run as a change: authorized, tested, and inside a maintenance window.
Do not put a drive root or a Windows directory in Custom paths. The script deletes the children of whatever path you give it.
Run
Open the script in Windows PowerShell 5.1.
Enter one server name per line.
Set the drive letter, temp age, and IIS retention.
Select the cleanup options. Uncheck anything you do not want.
Use WhatIf, or an alternate credential, if you need either.
Run Cleanup and watch the Process Viewer.
A failed WinRM check skips that server and continues with the next one.
