# =================================================================================================# ================================================================= more remote servers
# - Select cleanup actions
# - Run remote cleanup through PowerShell remoting
# - Review live step-by-step output in a separate Process Viewer window
# - Access built-in Help that explains every field, option, and command
#
# OUTPUT ON EACH REMOTE SERVER
# ----------------------------
#   C:\ProgramData\ServerCleanup
#
# REMOTE OUTPUT FILES
# -------------------
# - Timestamped .log file
# - Timestamped .csv file
#
# NOTES
# -----
# - Best run from Windows PowerShell 5.1
# - Run elevated
# - Target servers must allow PowerShell remoting
# - If needed, run on the target:
#     Enable-PSRemoting -Force
#
# =================================================================================================

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

# =========================================
# DEFAULTS
# =========================================
$DefaultDriveLetter = "C:"
$DefaultTempFileAgeDays = 2
$DefaultIISLogRetentionDays = 30
$DefaultOutputRoot = "C:\ProgramData\ServerCleanup"

$DefaultCleanUserTemp = $true
$DefaultCleanWindowsTemp = $true
$DefaultCleanRecycleBin = $true
$DefaultCleanSCCMCache = $true
$DefaultCleanWindowsUpdate = $true
$DefaultRunDismCleanup = $true
$DefaultCleanIISLogs = $false
$DefaultCleanCustomPaths = $false
$DefaultWhatIfMode = $false

# Layout offset so controls sit below the menu strip
$TopOffset = 28

function New-UiPoint {
    param(
        [int]$X,
        [int]$Y
    )

    return (New-Object System.Drawing.Point -ArgumentList $X, ($Y + $TopOffset))
}

# =========================================
# REMOTE CLEANUP SCRIPTBLOCK
# =========================================
$RemoteCleanupScript = {
    param(
        [string]$DriveLetter,
        [int]$TempFileAgeDays,
        [int]$IISLogRetentionDays,
        [bool]$CleanUserTemp,
        [bool]$CleanWindowsTemp,
        [bool]$CleanRecycleBin,
        [bool]$CleanSCCMCache,
        [bool]$CleanWindowsUpdate,
        [bool]$RunDISM,
        [bool]$CleanIISLogs,
        [bool]$CleanCustomPaths,
        [string[]]$CustomExtraPaths,
        [bool]$WhatIfMode,
        [string]$OutputRoot
    )

    if (-not (Test-Path -LiteralPath $OutputRoot)) {
        New-Item -Path $OutputRoot -ItemType Directory -Force | Out-Null
    }

    $RunTimeStamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $LogFile = Join-Path $OutputRoot ("ServerCleanup_REMOTE_{0}_{1}.log" -f $env:COMPUTERNAME, $RunTimeStamp)
    $CsvFile = Join-Path $OutputRoot ("ServerCleanup_REMOTE_{0}_{1}.csv" -f $env:COMPUTERNAME, $RunTimeStamp)

    $script:ResultRows = New-Object System.Collections.ArrayList

    function Add-ResultRow {
        param(
            [string]$Step,
            [string]$Status,
            [string]$Target,
            [string]$Details
        )

        $row = [pscustomobject]@{
            TimeStamp    = (Get-Date -Format "yyyy-MM-dd HH:mm:ss")
            ComputerName = $env:COMPUTERNAME
            Step         = $Step
            Status       = $Status
            Target       = $Target
            Details      = $Details
        }

        [void]$script:ResultRows.Add($row)
    }

    function Write-LogLine {
        param(
            [string]$Level,
            [string]$Message
        )

        $line = "[{0}] [{1}] [{2}] {3}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $Level, $env:COMPUTERNAME, $Message

        Add-Content -Path $LogFile -Value $line
        Write-Output $line
    }

    function Write-StepLine {
        param(
            [int]$StepNumber,
            [int]$TotalSteps,
            [string]$StatusText
        )

        Write-LogLine -Level "STEP" -Message ("{0} of {1} - {2}" -f $StepNumber, $TotalSteps, $StatusText)
    }

    function Get-FreeSpaceGB {
        param([string]$Drive = "C:")

        try {
            $disk = Get-WmiObject -Class Win32_LogicalDisk -Filter ("DeviceID='{0}'" -f $Drive.TrimEnd('\'))
            if ($disk -and $disk.FreeSpace -ne $null) {
                return [math]::Round(($disk.FreeSpace / 1GB), 2)
            }
        }
        catch {
            Write-LogLine -Level "WARN" -Message "Unable to read free space for drive $Drive - $($_.Exception.Message)"
            Add-ResultRow -Step "FreeSpaceCheck" -Status "WARN" -Target $Drive -Details $_.Exception.Message
        }

        return $null
    }

    function Remove-ItemsSafe {
        param(
            [string]$Path,
            [int]$OlderThanDays = 0,
            [string]$StepName = "Cleanup"
        )

        if ([string]::IsNullOrWhiteSpace($Path)) {
            Write-LogLine -Level "WARN" -Message "Empty path received in $StepName"
            Add-ResultRow -Step $StepName -Status "WARN" -Target "<empty>" -Details "Empty path received"
            return
        }

        if (-not (Test-Path -LiteralPath $Path)) {
            Write-LogLine -Level "WARN" -Message "Path not found: $Path"
            Add-ResultRow -Step $StepName -Status "WARN" -Target $Path -Details "Path not found"
            return
        }

        try {
            $items = Get-ChildItem -LiteralPath $Path -Force -ErrorAction SilentlyContinue

            if ($OlderThanDays -gt 0) {
                $cutoff = (Get-Date).AddDays(-1 * $OlderThanDays)
                $items = $items | Where-Object { $_.LastWriteTime -lt $cutoff }
            }

            $processed = 0
            $removed = 0
            $skipped = 0

            foreach ($item in $items) {
                $processed++

                try {
                    if ($WhatIfMode) {
                        Write-LogLine -Level "INFO" -Message "WhatIf - Would remove: $($item.FullName)"
                        $removed++
                    }
                    else {
                        Remove-Item -LiteralPath $item.FullName -Recurse -Force -ErrorAction Stop
                        $removed++
                    }
                }
                catch {
                    $skipped++
                    Write-LogLine -Level "WARN" -Message "Could not remove: $($item.FullName) - $($_.Exception.Message)"
                }
            }

            $details = "Processed=$processed Removed=$removed Skipped=$skipped"
            Write-LogLine -Level "OK" -Message "Processed path: $Path | $details"
            Add-ResultRow -Step $StepName -Status "OK" -Target $Path -Details $details
        }
        catch {
            Write-LogLine -Level "ERROR" -Message "Failed to process path: $Path - $($_.Exception.Message)"
            Add-ResultRow -Step $StepName -Status "ERROR" -Target $Path -Details $_.Exception.Message
        }
    }

    function Clear-UserTemp {
        Write-LogLine -Level "INFO" -Message "Cleaning current user temp paths"

        $tempPaths = @()
        if ($env:TEMP) { $tempPaths += $env:TEMP }
        if ($env:TMP -and $env:TMP -ne $env:TEMP) { $tempPaths += $env:TMP }

        $tempPaths = $tempPaths | Select-Object -Unique

        if (-not $tempPaths -or $tempPaths.Count -eq 0) {
            Write-LogLine -Level "WARN" -Message "No user temp paths found"
            Add-ResultRow -Step "UserTemp" -Status "WARN" -Target "" -Details "No user temp paths found"
            return
        }

        foreach ($path in $tempPaths) {
            Remove-ItemsSafe -Path $path -OlderThanDays $TempFileAgeDays -StepName "UserTemp"
        }
    }

    function Clear-WindowsTemp {
        Write-LogLine -Level "INFO" -Message "Cleaning Windows temp and servicing log paths"

        $paths = @(
            (Join-Path $env:windir "Temp"),
            (Join-Path $env:windir "Logs\CBS"),
            (Join-Path $env:windir "Logs\DISM")
        )

        foreach ($path in $paths) {
            Remove-ItemsSafe -Path $path -OlderThanDays $TempFileAgeDays -StepName "WindowsTemp"
        }
    }

    function Clear-RecycleBinSafe {
        $recyclePath = "$($DriveLetter.TrimEnd('\'))\`$Recycle.Bin"
        Write-LogLine -Level "INFO" -Message "Cleaning Recycle Bin path: $recyclePath"
        Remove-ItemsSafe -Path $recyclePath -OlderThanDays 0 -StepName "RecycleBin"
    }

    function Clear-SCCMCache {
        Write-LogLine -Level "INFO" -Message "Checking for SCCM / ConfigMgr cache"

        $found = $false

        try {
            $uiResource = New-Object -ComObject UIResource.UIResourceMgr -ErrorAction Stop
            $cacheInfo = $uiResource.GetCacheInfo()
            $elements = $cacheInfo.GetCacheElements()
            $found = $true

            if ($elements -and $elements.Count -gt 0) {
                foreach ($element in $elements) {
                    try {
                        if ($WhatIfMode) {
                            Write-LogLine -Level "INFO" -Message "WhatIf - Would delete SCCM cache element ID: $($element.CacheElementID)"
                        }
                        else {
                            $cacheInfo.DeleteCacheElement($element.CacheElementID)
                            Write-LogLine -Level "OK" -Message "Deleted SCCM cache element ID: $($element.CacheElementID)"
                        }
                    }
                    catch {
                        Write-LogLine -Level "WARN" -Message "Failed to remove SCCM cache element ID $($element.CacheElementID) - $($_.Exception.Message)"
                    }
                }
            }
            else {
                Write-LogLine -Level "INFO" -Message "SCCM cache found but no cache elements were returned"
            }
        }
        catch {
            Write-LogLine -Level "WARN" -Message "SCCM COM cleanup method not available on this server"
        }

        $ccmCachePath = Join-Path $env:windir "ccmcache"
        if (Test-Path -LiteralPath $ccmCachePath) {
            $found = $true
            Remove-ItemsSafe -Path $ccmCachePath -OlderThanDays 0 -StepName "SCCMCache"
        }

        if (-not $found) {
            Write-LogLine -Level "WARN" -Message "SCCM / ConfigMgr client cache not detected"
            Add-ResultRow -Step "SCCMCache" -Status "WARN" -Target "" -Details "SCCM / ConfigMgr client cache not detected"
        }
    }

    function Clear-WindowsUpdateDownload {
        $wuPath = Join-Path $env:windir "SoftwareDistribution\Download"

        if (-not (Test-Path -LiteralPath $wuPath)) {
            Write-LogLine -Level "WARN" -Message "Windows Update download cache path not found: $wuPath"
            Add-ResultRow -Step "WindowsUpdateCache" -Status "WARN" -Target $wuPath -Details "Path not found"
            return
        }

        $services = @("wuauserv", "bits", "cryptsvc")

        Write-LogLine -Level "INFO" -Message "Stopping Windows Update related services"
        foreach ($serviceName in $services) {
            try {
                $service = Get-Service -Name $serviceName -ErrorAction SilentlyContinue
                if ($service -and $service.Status -ne "Stopped") {
                    if ($WhatIfMode) {
                        Write-LogLine -Level "INFO" -Message "WhatIf - Would stop service: $serviceName"
                    }
                    else {
                        Stop-Service -Name $serviceName -Force -ErrorAction SilentlyContinue
                    }
                    Write-LogLine -Level "OK" -Message "Service stop processed: $serviceName"
                }
                else {
                    Write-LogLine -Level "INFO" -Message "Service already stopped or not found: $serviceName"
                }
            }
            catch {
                Write-LogLine -Level "WARN" -Message "Could not stop service $serviceName - $($_.Exception.Message)"
            }
        }

        Start-Sleep -Seconds 2
        Remove-ItemsSafe -Path $wuPath -OlderThanDays 0 -StepName "WindowsUpdateCache"

        Write-LogLine -Level "INFO" -Message "Starting Windows Update related services"
        foreach ($serviceName in $services) {
            try {
                $service = Get-Service -Name $serviceName -ErrorAction SilentlyContinue
                if ($service) {
                    if ($WhatIfMode) {
                        Write-LogLine -Level "INFO" -Message "WhatIf - Would start service: $serviceName"
                    }
                    else {
                        Start-Service -Name $serviceName -ErrorAction SilentlyContinue
                    }
                    Write-LogLine -Level "OK" -Message "Service start processed: $serviceName"
                }
            }
            catch {
                Write-LogLine -Level "WARN" -Message "Could not start service $serviceName - $($_.Exception.Message)"
            }
        }
    }

    function Run-DismCleanup {
        $dismPath = Join-Path $env:windir "System32\dism.exe"

        if (-not (Test-Path -LiteralPath $dismPath)) {
            Write-LogLine -Level "WARN" -Message "DISM executable not found at $dismPath"
            Add-ResultRow -Step "DISM" -Status "WARN" -Target $dismPath -Details "DISM executable not found"
            return
        }

        Write-LogLine -Level "INFO" -Message "Starting DISM component cleanup. This can take several minutes and may appear paused."

        try {
            if ($WhatIfMode) {
                Write-LogLine -Level "INFO" -Message "WhatIf - Would run: dism.exe /Online /Cleanup-Image /StartComponentCleanup"
                Add-ResultRow -Step "DISM" -Status "INFO" -Target $dismPath -Details "WhatIf only"
            }
            else {
                $startTime = Get-Date
                $process = Start-Process -FilePath $dismPath -ArgumentList "/Online /Cleanup-Image /StartComponentCleanup" -Wait -PassThru -NoNewWindow
                $elapsed = [math]::Round(((Get-Date) - $startTime).TotalMinutes, 2)

                Write-LogLine -Level "OK" -Message "DISM completed with exit code: $($process.ExitCode) | DurationMinutes=$elapsed"
                Add-ResultRow -Step "DISM" -Status "OK" -Target $dismPath -Details ("ExitCode={0};DurationMinutes={1}" -f $process.ExitCode, $elapsed)
            }
        }
        catch {
            Write-LogLine -Level "ERROR" -Message "DISM cleanup failed - $($_.Exception.Message)"
            Add-ResultRow -Step "DISM" -Status "ERROR" -Target $dismPath -Details $_.Exception.Message
        }
    }

    function Clear-IISLogs {
        $logPath = Join-Path $env:SystemDrive "inetpub\logs\LogFiles"
        Write-LogLine -Level "INFO" -Message "Cleaning IIS logs older than $IISLogRetentionDays days"

        if (-not (Test-Path -LiteralPath $logPath)) {
            Write-LogLine -Level "WARN" -Message "IIS log path not found: $logPath"
            Add-ResultRow -Step "IISLogs" -Status "WARN" -Target $logPath -Details "Path not found"
            return
        }

        try {
            $cutoff = (Get-Date).AddDays(-1 * $IISLogRetentionDays)
            $files = Get-ChildItem -Path $logPath -Recurse -File -ErrorAction SilentlyContinue | Where-Object { $_.LastWriteTime -lt $cutoff }

            $processed = 0
            $removed = 0
            $skipped = 0

            foreach ($file in $files) {
                $processed++

                try {
                    if ($WhatIfMode) {
                        Write-LogLine -Level "INFO" -Message "WhatIf - Would remove IIS log: $($file.FullName)"
                        $removed++
                    }
                    else {
                        Remove-Item -LiteralPath $file.FullName -Force -ErrorAction Stop
                        $removed++
                    }
                }
                catch {
                    $skipped++
                    Write-LogLine -Level "WARN" -Message "Could not remove IIS log: $($file.FullName) - $($_.Exception.Message)"
                }
            }

            $details = "Processed=$processed Removed=$removed Skipped=$skipped"
            Write-LogLine -Level "OK" -Message "IIS log cleanup complete | $details"
            Add-ResultRow -Step "IISLogs" -Status "OK" -Target $logPath -Details $details
        }
        catch {
            Write-LogLine -Level "ERROR" -Message "IIS log cleanup failed - $($_.Exception.Message)"
            Add-ResultRow -Step "IISLogs" -Status "ERROR" -Target $logPath -Details $_.Exception.Message
        }
    }

    function Clear-CustomExtraPaths {
        if (-not $CleanCustomPaths) {
            Write-LogLine -Level "INFO" -Message "Custom path cleanup disabled"
            Add-ResultRow -Step "CustomPaths" -Status "INFO" -Target "" -Details "Custom path cleanup disabled"
            return
        }

        if (-not $CustomExtraPaths -or $CustomExtraPaths.Count -eq 0) {
            Write-LogLine -Level "WARN" -Message "Custom path cleanup enabled but no paths were provided"
            Add-ResultRow -Step "CustomPaths" -Status "WARN" -Target "" -Details "Enabled but no paths provided"
            return
        }

        foreach ($path in $CustomExtraPaths) {
            Remove-ItemsSafe -Path $path -OlderThanDays $TempFileAgeDays -StepName "CustomPaths"
        }
    }

    $SelectedStepNames = @()
    if ($CleanUserTemp)      { $SelectedStepNames += "Clean user temp" }
    if ($CleanWindowsTemp)   { $SelectedStepNames += "Clean Windows temp" }
    if ($CleanRecycleBin)    { $SelectedStepNames += "Clean Recycle Bin" }
    if ($CleanSCCMCache)     { $SelectedStepNames += "Clean SCCM / ConfigMgr cache" }
    if ($CleanWindowsUpdate) { $SelectedStepNames += "Clean Windows Update download cache" }
    if ($RunDISM)            { $SelectedStepNames += "Run DISM component cleanup" }
    if ($CleanIISLogs)       { $SelectedStepNames += "Clean IIS logs" }
    if ($CleanCustomPaths)   { $SelectedStepNames += "Process custom extra paths" }

    $TotalSteps = $SelectedStepNames.Count + 1
    $CurrentStep = 0

    Write-LogLine -Level "INFO" -Message "=================================================================="
    Write-LogLine -Level "INFO" -Message "Starting REMOTE cleanup on $env:COMPUTERNAME"
    Write-LogLine -Level "INFO" -Message "Target drive: $DriveLetter"
    Write-LogLine -Level "INFO" -Message "WhatIf mode: $WhatIfMode"
    Write-LogLine -Level "INFO" -Message "Log file: $LogFile"
    Write-LogLine -Level "INFO" -Message "CSV file: $CsvFile"
    Write-LogLine -Level "INFO" -Message "=================================================================="

    $CurrentStep++
    Write-StepLine -StepNumber $CurrentStep -TotalSteps $TotalSteps -StatusText "Measure free space before cleanup"
    $BeforeFreeGB = Get-FreeSpaceGB -Drive $DriveLetter
    Write-LogLine -Level "INFO" -Message "Free space before cleanup on $DriveLetter : $BeforeFreeGB GB"

    if ($CleanUserTemp) {
        $CurrentStep++
        Write-StepLine -StepNumber $CurrentStep -TotalSteps $TotalSteps -StatusText "Clean user temp"
        Clear-UserTemp
    }

    if ($CleanWindowsTemp) {
        $CurrentStep++
        Write-StepLine -StepNumber $CurrentStep -TotalSteps $TotalSteps -StatusText "Clean Windows temp"
        Clear-WindowsTemp
    }

    if ($CleanRecycleBin) {
        $CurrentStep++
        Write-StepLine -StepNumber $CurrentStep -TotalSteps $TotalSteps -StatusText "Clean Recycle Bin"
        Clear-RecycleBinSafe
    }

    if ($CleanSCCMCache) {
        $CurrentStep++
        Write-StepLine -StepNumber $CurrentStep -TotalSteps $TotalSteps -StatusText "Clean SCCM / ConfigMgr cache"
        Clear-SCCMCache
    }

    if ($CleanWindowsUpdate) {
        $CurrentStep++
        Write-StepLine -StepNumber $CurrentStep -TotalSteps $TotalSteps -StatusText "Clean Windows Update download cache"
        Clear-WindowsUpdateDownload
    }

    if ($RunDISM) {
        $CurrentStep++
        Write-StepLine -StepNumber $CurrentStep -TotalSteps $TotalSteps -StatusText "Run DISM component cleanup"
        Run-DismCleanup
    }

    if ($CleanIISLogs) {
        $CurrentStep++
        Write-StepLine -StepNumber $CurrentStep -TotalSteps $TotalSteps -StatusText "Clean IIS logs"
        Clear-IISLogs
    }

    if ($CleanCustomPaths) {
        $CurrentStep++
        Write-StepLine -StepNumber $CurrentStep -TotalSteps $TotalSteps -StatusText "Process custom extra paths"
        Clear-CustomExtraPaths
    }

    Start-Sleep -Seconds 2

    $AfterFreeGB = Get-FreeSpaceGB -Drive $DriveLetter
    Write-LogLine -Level "INFO" -Message "Free space after cleanup on $DriveLetter : $AfterFreeGB GB"

    if ($BeforeFreeGB -ne $null -and $AfterFreeGB -ne $null) {
        $RecoveredGB = [math]::Round(($AfterFreeGB - $BeforeFreeGB), 2)
        Write-LogLine -Level "OK" -Message "Estimated space recovered: $RecoveredGB GB"
        Add-ResultRow -Step "Summary" -Status "OK" -Target $DriveLetter -Details ("RecoveredGB={0}" -f $RecoveredGB)
    }

    $script:ResultRows | Export-Csv -Path $CsvFile -NoTypeInformation -Force
    Write-LogLine -Level "DONE" -Message "Remote cleanup completed"
}

# =========================================
# MAIN FORM
# =========================================
$form = New-Object System.Windows.Forms.Form
$form.Text = "Remote Server Cleanup"
$form.Size = New-Object System.Drawing.Size(840,860)
$form.StartPosition = "CenterScreen"
$form.Topmost = $true
$form.FormBorderStyle = 'FixedDialog'
$form.MaximizeBox = $false
$form.KeyPreview = $true

$fontRegular = New-Object System.Drawing.Font("Segoe UI",9)
$fontBold = New-Object System.Drawing.Font("Segoe UI",9,[System.Drawing.FontStyle]::Bold)
$form.Font = $fontRegular

# =========================================
# MENU STRIP
# =========================================
$menuMain = New-Object System.Windows.Forms.MenuStrip
$menuMain.Dock = 'Top'
$form.MainMenuStrip = $menuMain
$form.Controls.Add($menuMain)

$menuTools = New-Object System.Windows.Forms.ToolStripMenuItem
$menuTools.Text = "&Tools"
[void]$menuMain.Items.Add($menuTools)

$menuToolsOpenViewer = New-Object System.Windows.Forms.ToolStripMenuItem
$menuToolsOpenViewer.Text = "Open Process Viewer"
[void]$menuTools.DropDownItems.Add($menuToolsOpenViewer)

$menuToolsClearViewer = New-Object System.Windows.Forms.ToolStripMenuItem
$menuToolsClearViewer.Text = "Clear Process Viewer"
[void]$menuTools.DropDownItems.Add($menuToolsClearViewer)

$menuToolsSaveViewer = New-Object System.Windows.Forms.ToolStripMenuItem
$menuToolsSaveViewer.Text = "Save Process Viewer Log"
[void]$menuTools.DropDownItems.Add($menuToolsSaveViewer)

[void]$menuTools.DropDownItems.Add((New-Object System.Windows.Forms.ToolStripSeparator))

$menuToolsTestWinRM = New-Object System.Windows.Forms.ToolStripMenuItem
$menuToolsTestWinRM.Text = "Test WinRM for Listed Servers"
[void]$menuTools.DropDownItems.Add($menuToolsTestWinRM)

$menuHelp = New-Object System.Windows.Forms.ToolStripMenuItem
$menuHelp.Text = "&Help"
[void]$menuMain.Items.Add($menuHelp)

$menuHelpContents = New-Object System.Windows.Forms.ToolStripMenuItem
$menuHelpContents.Text = "Help Contents"
$menuHelpContents.ShortcutKeys = [System.Windows.Forms.Keys]::F1
[void]$menuHelp.DropDownItems.Add($menuHelpContents)

$menuHelpFeatures = New-Object System.Windows.Forms.ToolStripMenuItem
$menuHelpFeatures.Text = "Features and Options"
[void]$menuHelp.DropDownItems.Add($menuHelpFeatures)

$menuHelpCommands = New-Object System.Windows.Forms.ToolStripMenuItem
$menuHelpCommands.Text = "Specific Commands"
[void]$menuHelp.DropDownItems.Add($menuHelpCommands)

$menuHelpViewer = New-Object System.Windows.Forms.ToolStripMenuItem
$menuHelpViewer.Text = "Open Process Viewer"
[void]$menuHelp.DropDownItems.Add($menuHelpViewer)

[void]$menuHelp.DropDownItems.Add((New-Object System.Windows.Forms.ToolStripSeparator))

$menuHelpAbout = New-Object System.Windows.Forms.ToolStripMenuItem
$menuHelpAbout.Text = "About"
[void]$menuHelp.DropDownItems.Add($menuHelpAbout)

# =========================================
# PROCESS VIEWER FORM
# =========================================
$script:AllowProcessViewerClose = $false

$processForm = New-Object System.Windows.Forms.Form
$processForm.Text = "Process Viewer"
$processForm.Size = New-Object System.Drawing.Size(980,580)
$processForm.StartPosition = "Manual"
$processForm.Location = New-Object System.Drawing.Point(860,90)
$processForm.Topmost = $false

$txtProcess = New-Object System.Windows.Forms.RichTextBox
$txtProcess.Location = New-Object System.Drawing.Point(10,10)
$txtProcess.Size = New-Object System.Drawing.Size(944,470)
$txtProcess.ReadOnly = $true
$txtProcess.Multiline = $true
$txtProcess.WordWrap = $false
$txtProcess.Font = New-Object System.Drawing.Font("Consolas",9)
$txtProcess.BackColor = [System.Drawing.Color]::White
$txtProcess.ScrollBars = "Both"
$processForm.Controls.Add($txtProcess)

$btnClearViewer = New-Object System.Windows.Forms.Button
$btnClearViewer.Location = New-Object System.Drawing.Point(10,490)
$btnClearViewer.Size = New-Object System.Drawing.Size(120,28)
$btnClearViewer.Text = "Clear Viewer"
$processForm.Controls.Add($btnClearViewer)

$btnHideViewer = New-Object System.Windows.Forms.Button
$btnHideViewer.Location = New-Object System.Drawing.Point(140,490)
$btnHideViewer.Size = New-Object System.Drawing.Size(120,28)
$btnHideViewer.Text = "Hide Viewer"
$processForm.Controls.Add($btnHideViewer)

$btnSaveViewer = New-Object System.Windows.Forms.Button
$btnSaveViewer.Location = New-Object System.Drawing.Point(270,490)
$btnSaveViewer.Size = New-Object System.Drawing.Size(140,28)
$btnSaveViewer.Text = "Save Viewer Log"
$processForm.Controls.Add($btnSaveViewer)

$processForm.Add_FormClosing({
    param($sender, $e)

    if (-not $script:AllowProcessViewerClose) {
        $e.Cancel = $true
        $sender.Hide()
    }
})

function Add-ViewerLine {
    param(
        [string]$Message,
        [string]$Color = "Black"
    )

    if (-not $processForm.Visible) {
        $processForm.Show()
    }

    $txtProcess.SelectionStart = $txtProcess.TextLength
    $txtProcess.SelectionLength = 0

    try {
        $txtProcess.SelectionColor = [System.Drawing.Color]::$Color
    }
    catch {
        $txtProcess.SelectionColor = [System.Drawing.Color]::Black
    }

    $txtProcess.AppendText($Message + "`r`n")
    $txtProcess.SelectionColor = $txtProcess.ForeColor
    $txtProcess.ScrollToCaret()
    [System.Windows.Forms.Application]::DoEvents()
}

# =========================================
# STATUS HELPERS
# =========================================
function Set-StatusBarText {
    param([string]$Message)

    if ($statusLabel) {
        $statusLabel.Text = $Message
    }

    [System.Windows.Forms.Application]::DoEvents()
}

function Add-MainStatusLine {
    param([string]$Message)

    if ($txtStatus) {
        $txtStatus.AppendText($Message + "`r`n")
        $txtStatus.SelectionStart = $txtStatus.Text.Length
        $txtStatus.ScrollToCaret()
    }

    Set-StatusBarText -Message $Message
    [System.Windows.Forms.Application]::DoEvents()
}

function Set-OverallProgress {
    param(
        [int]$Current,
        [int]$Total,
        [string]$LabelText = ""
    )

    if ($progressBar) {
        if ($Total -le 0) {
            $progressBar.Value = 0
        }
        else {
            $percent = [int][math]::Round((($Current / $Total) * 100), 0)

            if ($percent -lt 0)   { $percent = 0 }
            if ($percent -gt 100) { $percent = 100 }

            $progressBar.Value = $percent
        }
    }

    if (-not [string]::IsNullOrWhiteSpace($LabelText)) {
        Set-StatusBarText -Message $LabelText
    }

    [System.Windows.Forms.Application]::DoEvents()
}

function Show-AboutDialog {
    [System.Windows.Forms.MessageBox]::Show(
@"
Remote Server Cleanup

Purpose
- Run selected cleanup tasks against one or more remote servers

Key capabilities
- Multi-server GUI
- Process Viewer
- WhatIf support
- Help and command reference
- Remote log and CSV output

Output location on each remote server
- C:\ProgramData\ServerCleanup
"@,
        "About - Remote Server Cleanup",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Information
    ) | Out-Null
}

function Save-ViewerLog {
    $saveDialog = New-Object System.Windows.Forms.SaveFileDialog
    $saveDialog.Filter = "Log files (*.log)|*.log|Text files (*.txt)|*.txt|All files (*.*)|*.*"
    $saveDialog.FileName = "ProcessViewer_{0}.log" -f (Get-Date -Format "yyyyMMdd_HHmmss")

    if ($saveDialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        try {
            [System.IO.File]::WriteAllText($saveDialog.FileName, $txtProcess.Text)
            [System.Windows.Forms.MessageBox]::Show(
                "Viewer log saved successfully.",
                "Saved",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information
            ) | Out-Null
        }
        catch {
            [System.Windows.Forms.MessageBox]::Show(
                "Failed to save viewer log.`r`n$($_.Exception.Message)",
                "Error",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Error
            ) | Out-Null
        }
    }
}

function Open-ProcessViewer {
    if (-not $processForm.Visible) {
        $processForm.Show()
    }
    else {
        $processForm.Activate()
    }
}

function Get-ListedServers {
    return @(
        $txtServers.Lines |
        ForEach-Object { $_.Trim() } |
        Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
        Select-Object -Unique
    )
}

function Test-ListedServersWinRM {
    $servers = Get-ListedServers

    if (-not $servers -or $servers.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show(
            "Please enter at least one server before testing WinRM.",
            "Validation",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        ) | Out-Null
        return
    }

    Open-ProcessViewer
    Add-MainStatusLine "Starting WinRM test for listed servers..."
    Add-ViewerLine -Message ("[{0}] [INFO] [LOCAL] Starting WinRM test for listed servers" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss")) -Color "Blue"

    $total = $servers.Count
    $index = 0

    foreach ($server in $servers) {
        $index++
        Set-OverallProgress -Current $index -Total $total -LabelText ("Testing WinRM: {0} ({1}/{2})" -f $server, $index, $total)

        Add-MainStatusLine "Testing WinRM for $server..."
        Add-ViewerLine -Message ("[{0}] [STEP] [LOCAL] Testing WinRM for {1}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $server) -Color "Blue"

        try {
            Test-WSMan -ComputerName $server -ErrorAction Stop | Out-Null
            Add-MainStatusLine "WinRM OK for $server"
            Add-ViewerLine -Message ("[{0}] [OK] [LOCAL] WinRM OK for {1}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $server) -Color "Green"
        }
        catch {
            Add-MainStatusLine ("WinRM FAILED for {0} : {1}" -f $server, $_.Exception.Message)
            Add-ViewerLine -Message ("[{0}] [ERROR] [LOCAL] WinRM FAILED for {1} - {2}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $server, $_.Exception.Message) -Color "Red"
        }
    }

    Set-OverallProgress -Current $total -Total $total -LabelText "WinRM test completed."
    Add-MainStatusLine "WinRM test completed."
    Add-ViewerLine -Message ("[{0}] [DONE] [LOCAL] WinRM test completed" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss")) -Color "DarkGreen"
}

# =========================================
# HELP WINDOW CONTENT
# =========================================
$HelpText_Contents = @"
REMOTE SERVER CLEANUP - HELP CONTENTS
====================================

OVERVIEW
--------
This tool connects to one or more remote Windows servers through PowerShell remoting
and performs selected cleanup tasks.

REQUIREMENTS
------------
- Run this tool elevated
- Target servers must allow PowerShell remoting
- If needed, run on target:
  Enable-PSRemoting -Force

MAIN FORM FIELDS
----------------
Servers (one per line)
- Enter one server per line
- Duplicate names are removed automatically

Drive letter
- Used for free-space measurement
- Used for Recycle Bin cleanup
- Usually C:

Temp age days
- Older files are eligible for removal from temp-related locations

IIS retention days
- IIS logs older than this number of days are eligible for deletion

WhatIf mode
- Shows what would happen without actually deleting files

Cleanup options
---------------
User temp
- Cleans current user TEMP and TMP locations

Windows temp
- Cleans:
  * %windir%\Temp
  * %windir%\Logs\CBS
  * %windir%\Logs\DISM

Recycle Bin
- Cleans the Recycle Bin on the selected drive

SCCM / ConfigMgr cache
- Attempts COM-based cleanup
- Also checks %windir%\ccmcache

Windows Update download cache
- Stops update-related services
- Cleans %windir%\SoftwareDistribution\Download
- Starts services again

DISM component cleanup
- Runs:
  dism.exe /Online /Cleanup-Image /StartComponentCleanup

IIS logs
- Deletes IIS logs older than the IIS retention value

Custom paths
- Cleans operator-supplied directories

Use alternate credential
- Prompts for alternate credentials

PROCESS VIEWER
--------------
Displays detailed step-by-step runtime output similar to an ISE-like output pane.

OUTPUT LOCATION
---------------
On each remote server:
C:\ProgramData\ServerCleanup

Files written:
- Timestamped .log file
- Timestamped .csv file

BEST PRACTICES
--------------
- Start with one server
- Use WhatIf mode first
- Review logs after each run
- Use custom paths carefully
"@

$HelpText_Features = @"
FEATURES AND OPTIONS - VERBOSE EXPLANATION
==========================================

1. SERVERS
----------
Purpose:
- Defines the remote systems to process

Behavior:
- Blank lines are ignored
- Duplicate entries are removed
- DNS is tested before remote execution

2. DRIVE LETTER
---------------
Purpose:
- Used for free-space checks
- Used to locate the Recycle Bin path

3. TEMP AGE DAYS
----------------
Purpose:
- Prevents deletion of very recent temp files

Applies to:
- User temp cleanup
- Windows temp cleanup
- Custom path cleanup

4. IIS RETENTION DAYS
---------------------
Purpose:
- Retains newer IIS logs while deleting older ones

5. WHATIF MODE
--------------
Purpose:
- Preview mode

Behavior:
- Logs intended actions
- Does not remove files
- Does not truly run DISM

6. USER TEMP
------------
Targets:
- $env:TEMP
- $env:TMP

7. WINDOWS TEMP
---------------
Targets:
- C:\Windows\Temp
- C:\Windows\Logs\CBS
- C:\Windows\Logs\DISM

8. RECYCLE BIN
--------------
Target:
- C:\$Recycle.Bin or the selected drive equivalent

9. SCCM / CONFIGMGR CACHE
-------------------------
Methods:
- COM cleanup through UIResource.UIResourceMgr
- Direct cleanup of %windir%\ccmcache

10. WINDOWS UPDATE DOWNLOAD CACHE
---------------------------------
Target:
- %windir%\SoftwareDistribution\Download

Services:
- wuauserv
- bits
- cryptsvc

11. DISM COMPONENT CLEANUP
--------------------------
Command:
- dism.exe /Online /Cleanup-Image /StartComponentCleanup

Purpose:
- Reduces component store overhead

Notes:
- Can take several minutes
- Best used in a maintenance window for sensitive systems

12. IIS LOGS
------------
Target:
- C:\inetpub\logs\LogFiles

13. CUSTOM PATHS
----------------
Purpose:
- Cleans additional directories supplied by the operator

Examples:
- C:\Temp\App1
- D:\Staging\OldFiles
- C:\Logs\Archive

14. ALTERNATE CREDENTIALS
-------------------------
Purpose:
- Use another account for remote execution

15. PROCESS VIEWER
------------------
Purpose:
- Displays live detailed output

Color meaning:
- Blue = step
- Green = success
- Orange = warning
- Red = error
- Dark green = complete

16. LOGS AND CSV
----------------
Remote output folder:
- C:\ProgramData\ServerCleanup

CSV fields:
- TimeStamp
- ComputerName
- Step
- Status
- Target
- Details
"@

$HelpText_Commands = @"
SPECIFIC COMMANDS USED BY THIS TOOL
===================================

REMOTE COMMANDS
---------------
Test-WSMan -ComputerName SERVER
- Verifies WinRM connectivity

Invoke-Command -ComputerName SERVER -ScriptBlock { ... }
- Executes cleanup remotely

Invoke-Command -ComputerName SERVER -Credential CREDENTIAL -ScriptBlock { ... }
- Executes cleanup remotely with alternate credentials

FILE AND PATH COMMANDS
----------------------
Test-Path -LiteralPath PATH
- Checks whether a path exists

Get-ChildItem -LiteralPath PATH -Force
- Enumerates items including hidden items

Remove-Item -LiteralPath ITEM -Recurse -Force
- Removes files and folders recursively

Join-Path BASE CHILD
- Safely builds full paths

SERVICE COMMANDS
----------------
Get-Service -Name wuauserv
Get-Service -Name bits
Get-Service -Name cryptsvc

Stop-Service -Name wuauserv -Force
Stop-Service -Name bits -Force
Stop-Service -Name cryptsvc -Force

Start-Service -Name wuauserv
Start-Service -Name bits
Start-Service -Name cryptsvc

DISK SPACE CHECK
----------------
Get-WmiObject -Class Win32_LogicalDisk -Filter "DeviceID='C:'"

FREE SPACE CALCULATION
----------------------
[math]::Round(($disk.FreeSpace / 1GB), 2)

DISM COMMAND
------------
dism.exe /Online /Cleanup-Image /StartComponentCleanup

SCCM / CONFIGMGR COMMANDS
-------------------------
New-Object -ComObject UIResource.UIResourceMgr
$uiResource.GetCacheInfo()
$cacheInfo.GetCacheElements()
$cacheInfo.DeleteCacheElement(ID)

CSV EXPORT
----------
Export-Csv -Path FILE -NoTypeInformation -Force

LOGGING
-------
Add-Content -Path LOGFILE -Value LINE
Write-Output LINE

COMMON MANUAL TEST COMMANDS
---------------------------
1. Test remoting:
   Test-WSMan SERVER01

2. Test remote command:
   Invoke-Command -ComputerName SERVER01 -ScriptBlock { hostname }

3. Enable remoting on target:
   Enable-PSRemoting -Force

4. Check WinRM service:
   Get-Service WinRM

5. Run DISM manually:
   dism.exe /Online /Cleanup-Image /StartComponentCleanup

6. Check Windows Update cache:
   Test-Path C:\Windows\SoftwareDistribution\Download

7. Check IIS logs:
   Test-Path C:\inetpub\logs\LogFiles

8. Check SCCM cache:
   Test-Path C:\Windows\ccmcache
"@

$HelpText_About = @"
ABOUT REMOTE SERVER CLEANUP
===========================

Purpose
-------
This tool provides a GUI-driven way to perform common Windows server cleanup tasks
across one or more remote servers through PowerShell remoting.

Designed for
------------
- Windows PowerShell 5.1
- Windows Server 2016+
- Multi-server admin workflows
- Optional alternate credentials
- Step-by-step visibility through the Process Viewer

Key benefits
------------
- One GUI for many servers
- Detailed Process Viewer
- Help and command reference
- WhatIf support
- Remote logs and CSV export
"@

function Show-HelpWindow {
    param(
        [int]$StartTab = 0
    )

    if ($script:HelpForm -and -not $script:HelpForm.IsDisposed) {
        $script:HelpTabs.SelectedIndex = $StartTab
        $script:HelpForm.Show()
        $script:HelpForm.Activate()
        return
    }

    $script:HelpForm = New-Object System.Windows.Forms.Form
    $script:HelpForm.Text = "Help - Remote Server Cleanup"
    $script:HelpForm.Size = New-Object System.Drawing.Size(1000,730)
    $script:HelpForm.StartPosition = "CenterParent"
    $script:HelpForm.Topmost = $false

    $script:HelpTabs = New-Object System.Windows.Forms.TabControl
    $script:HelpTabs.Location = New-Object System.Drawing.Point(10,10)
    $script:HelpTabs.Size = New-Object System.Drawing.Size(964,600)
    $script:HelpForm.Controls.Add($script:HelpTabs)

    function New-HelpTab {
        param(
            [string]$Title,
            [string]$TextContent
        )

        $tab = New-Object System.Windows.Forms.TabPage
        $tab.Text = $Title

        $rtb = New-Object System.Windows.Forms.RichTextBox
        $rtb.Location = New-Object System.Drawing.Point(5,5)
        $rtb.Size = New-Object System.Drawing.Size(945,560)
        $rtb.ReadOnly = $true
        $rtb.Multiline = $true
        $rtb.WordWrap = $false
        $rtb.ScrollBars = "Both"
        $rtb.Font = New-Object System.Drawing.Font("Consolas",10)
        $rtb.Text = $TextContent
        $tab.Tag = $rtb
        $tab.Controls.Add($rtb)

        return $tab
    }

    [void]$script:HelpTabs.TabPages.Add((New-HelpTab -Title "Contents" -TextContent $HelpText_Contents))
    [void]$script:HelpTabs.TabPages.Add((New-HelpTab -Title "Features" -TextContent $HelpText_Features))
    [void]$script:HelpTabs.TabPages.Add((New-HelpTab -Title "Commands" -TextContent $HelpText_Commands))
    [void]$script:HelpTabs.TabPages.Add((New-HelpTab -Title "About" -TextContent $HelpText_About))
    $script:HelpTabs.SelectedIndex = $StartTab

    $lblHelpSearch = New-Object System.Windows.Forms.Label
    $lblHelpSearch.Location = New-Object System.Drawing.Point(10,620)
    $lblHelpSearch.Size = New-Object System.Drawing.Size(80,22)
    $lblHelpSearch.Text = "Search:"
    $script:HelpForm.Controls.Add($lblHelpSearch)

    $txtHelpSearch = New-Object System.Windows.Forms.TextBox
    $txtHelpSearch.Location = New-Object System.Drawing.Point(70,618)
    $txtHelpSearch.Size = New-Object System.Drawing.Size(300,24)
    $script:HelpForm.Controls.Add($txtHelpSearch)

    $btnHelpFindNext = New-Object System.Windows.Forms.Button
    $btnHelpFindNext.Location = New-Object System.Drawing.Point(380,616)
    $btnHelpFindNext.Size = New-Object System.Drawing.Size(100,28)
    $btnHelpFindNext.Text = "Find Next"
    $script:HelpForm.Controls.Add($btnHelpFindNext)

    $btnHelpClose = New-Object System.Windows.Forms.Button
    $btnHelpClose.Location = New-Object System.Drawing.Point(10,655)
    $btnHelpClose.Size = New-Object System.Drawing.Size(120,30)
    $btnHelpClose.Text = "Close"
    $script:HelpForm.Controls.Add($btnHelpClose)

    $btnHelpClose.Add_Click({
        $script:HelpForm.Close()
    })

    $script:HelpSearchState = [ordered]@{
        LastText     = ""
        LastTabIndex = -1
        LastIndex    = 0
    }

    $btnHelpFindNext.Add_Click({
        $findText = $txtHelpSearch.Text

        if ([string]::IsNullOrWhiteSpace($findText)) {
            return
        }

        $tab = $script:HelpTabs.SelectedTab
        if (-not $tab) {
            return
        }

        $rtb = $tab.Tag
        if (-not $rtb) {
            return
        }

        if (
            $script:HelpSearchState.LastText -ne $findText -or
            $script:HelpSearchState.LastTabIndex -ne $script:HelpTabs.SelectedIndex
        ) {
            $script:HelpSearchState.LastText = $findText
            $script:HelpSearchState.LastTabIndex = $script:HelpTabs.SelectedIndex
            $script:HelpSearchState.LastIndex = 0
        }

        $startIndex = $script:HelpSearchState.LastIndex
        $foundIndex = $rtb.Text.IndexOf($findText, $startIndex, [System.StringComparison]::OrdinalIgnoreCase)

        if ($foundIndex -lt 0 -and $startIndex -gt 0) {
            $foundIndex = $rtb.Text.IndexOf($findText, 0, [System.StringComparison]::OrdinalIgnoreCase)
        }

        if ($foundIndex -ge 0) {
            $rtb.Focus()
            $rtb.SelectionStart = $foundIndex
            $rtb.SelectionLength = $findText.Length
            $rtb.ScrollToCaret()
            $script:HelpSearchState.LastIndex = $foundIndex + $findText.Length
        }
        else {
            [System.Windows.Forms.MessageBox]::Show(
                "Text not found in the current tab.",
                "Search",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Information
            ) | Out-Null

            $script:HelpSearchState.LastIndex = 0
        }
    })

    [void]$script:HelpForm.Show()
}

# =========================================
# MAIN GUI CONTROLS
# =========================================
$lblServers = New-Object System.Windows.Forms.Label
$lblServers.Location = New-UiPoint 15 15
$lblServers.Size = New-Object System.Drawing.Size(300,20)
$lblServers.Text = "Servers (one per line):"
$lblServers.Font = $fontBold
$form.Controls.Add($lblServers)

$txtServers = New-Object System.Windows.Forms.TextBox
$txtServers.Location = New-UiPoint 15 40
$txtServers.Size = New-Object System.Drawing.Size(790,180)
$txtServers.Multiline = $true
$txtServers.ScrollBars = "Vertical"
$form.Controls.Add($txtServers)

$lblDrive = New-Object System.Windows.Forms.Label
$lblDrive.Location = New-UiPoint 15 235
$lblDrive.Size = New-Object System.Drawing.Size(100,20)
$lblDrive.Text = "Drive letter:"
$form.Controls.Add($lblDrive)

$txtDrive = New-Object System.Windows.Forms.TextBox
$txtDrive.Location = New-UiPoint 120 232
$txtDrive.Size = New-Object System.Drawing.Size(70,24)
$txtDrive.Text = $DefaultDriveLetter
$form.Controls.Add($txtDrive)

$lblTempAge = New-Object System.Windows.Forms.Label
$lblTempAge.Location = New-UiPoint 220 235
$lblTempAge.Size = New-Object System.Drawing.Size(110,20)
$lblTempAge.Text = "Temp age days:"
$form.Controls.Add($lblTempAge)

$txtTempAge = New-Object System.Windows.Forms.TextBox
$txtTempAge.Location = New-UiPoint 335 232
$txtTempAge.Size = New-Object System.Drawing.Size(60,24)
$txtTempAge.Text = [string]$DefaultTempFileAgeDays
$form.Controls.Add($txtTempAge)

$lblIISRetention = New-Object System.Windows.Forms.Label
$lblIISRetention.Location = New-UiPoint 420 235
$lblIISRetention.Size = New-Object System.Drawing.Size(125,20)
$lblIISRetention.Text = "IIS retention days:"
$form.Controls.Add($lblIISRetention)

$txtIISRetention = New-Object System.Windows.Forms.TextBox
$txtIISRetention.Location = New-UiPoint 550 232
$txtIISRetention.Size = New-Object System.Drawing.Size(60,24)
$txtIISRetention.Text = [string]$DefaultIISLogRetentionDays
$form.Controls.Add($txtIISRetention)

$chkWhatIf = New-Object System.Windows.Forms.CheckBox
$chkWhatIf.Location = New-UiPoint 640 232
$chkWhatIf.Size = New-Object System.Drawing.Size(120,24)
$chkWhatIf.Text = "WhatIf mode"
$chkWhatIf.Checked = $DefaultWhatIfMode
$form.Controls.Add($chkWhatIf)

$grpOptions = New-Object System.Windows.Forms.GroupBox
$grpOptions.Location = New-UiPoint 15 270
$grpOptions.Size = New-Object System.Drawing.Size(790,145)
$grpOptions.Text = "Cleanup options"
$grpOptions.Font = $fontBold
$form.Controls.Add($grpOptions)

$chkUserTemp = New-Object System.Windows.Forms.CheckBox
$chkUserTemp.Location = New-Object System.Drawing.Point(15,25)
$chkUserTemp.Size = New-Object System.Drawing.Size(180,22)
$chkUserTemp.Text = "User temp"
$chkUserTemp.Checked = $DefaultCleanUserTemp
$grpOptions.Controls.Add($chkUserTemp)

$chkWindowsTemp = New-Object System.Windows.Forms.CheckBox
$chkWindowsTemp.Location = New-Object System.Drawing.Point(200,25)
$chkWindowsTemp.Size = New-Object System.Drawing.Size(180,22)
$chkWindowsTemp.Text = "Windows temp"
$chkWindowsTemp.Checked = $DefaultCleanWindowsTemp
$grpOptions.Controls.Add($chkWindowsTemp)

$chkRecycleBin = New-Object System.Windows.Forms.CheckBox
$chkRecycleBin.Location = New-Object System.Drawing.Point(385,25)
$chkRecycleBin.Size = New-Object System.Drawing.Size(180,22)
$chkRecycleBin.Text = "Recycle Bin"
$chkRecycleBin.Checked = $DefaultCleanRecycleBin
$grpOptions.Controls.Add($chkRecycleBin)

$chkSCCM = New-Object System.Windows.Forms.CheckBox
$chkSCCM.Location = New-Object System.Drawing.Point(570,25)
$chkSCCM.Size = New-Object System.Drawing.Size(200,22)
$chkSCCM.Text = "SCCM / ConfigMgr cache"
$chkSCCM.Checked = $DefaultCleanSCCMCache
$grpOptions.Controls.Add($chkSCCM)

$chkWindowsUpdate = New-Object System.Windows.Forms.CheckBox
$chkWindowsUpdate.Location = New-Object System.Drawing.Point(15,55)
$chkWindowsUpdate.Size = New-Object System.Drawing.Size(220,22)
$chkWindowsUpdate.Text = "Windows Update download cache"
$chkWindowsUpdate.Checked = $DefaultCleanWindowsUpdate
$grpOptions.Controls.Add($chkWindowsUpdate)

$chkDISM = New-Object System.Windows.Forms.CheckBox
$chkDISM.Location = New-Object System.Drawing.Point(240,55)
$chkDISM.Size = New-Object System.Drawing.Size(180,22)
$chkDISM.Text = "DISM component cleanup"
$chkDISM.Checked = $DefaultRunDismCleanup
$grpOptions.Controls.Add($chkDISM)

$chkIIS = New-Object System.Windows.Forms.CheckBox
$chkIIS.Location = New-Object System.Drawing.Point(425,55)
$chkIIS.Size = New-Object System.Drawing.Size(180,22)
$chkIIS.Text = "IIS logs"
$chkIIS.Checked = $DefaultCleanIISLogs
$grpOptions.Controls.Add($chkIIS)

$chkCustomPaths = New-Object System.Windows.Forms.CheckBox
$chkCustomPaths.Location = New-Object System.Drawing.Point(610,55)
$chkCustomPaths.Size = New-Object System.Drawing.Size(140,22)
$chkCustomPaths.Text = "Custom paths"
$chkCustomPaths.Checked = $DefaultCleanCustomPaths
$grpOptions.Controls.Add($chkCustomPaths)

$chkAltCred = New-Object System.Windows.Forms.CheckBox
$chkAltCred.Location = New-Object System.Drawing.Point(15,85)
$chkAltCred.Size = New-Object System.Drawing.Size(180,22)
$chkAltCred.Text = "Use alternate credential"
$chkAltCred.Checked = $false
$grpOptions.Controls.Add($chkAltCred)

$btnSelectAll = New-Object System.Windows.Forms.Button
$btnSelectAll.Location = New-Object System.Drawing.Point(520,100)
$btnSelectAll.Size = New-Object System.Drawing.Size(100,28)
$btnSelectAll.Text = "Select All"
$grpOptions.Controls.Add($btnSelectAll)

$btnSelectNone = New-Object System.Windows.Forms.Button
$btnSelectNone.Location = New-Object System.Drawing.Point(630,100)
$btnSelectNone.Size = New-Object System.Drawing.Size(100,28)
$btnSelectNone.Text = "Select None"
$grpOptions.Controls.Add($btnSelectNone)

$lblCustomPaths = New-Object System.Windows.Forms.Label
$lblCustomPaths.Location = New-UiPoint 15 425
$lblCustomPaths.Size = New-Object System.Drawing.Size(330,20)
$lblCustomPaths.Text = "Custom extra paths (optional, one per line):"
$lblCustomPaths.Font = $fontBold
$form.Controls.Add($lblCustomPaths)

$txtCustomPaths = New-Object System.Windows.Forms.TextBox
$txtCustomPaths.Location = New-UiPoint 15 450
$txtCustomPaths.Size = New-Object System.Drawing.Size(790,110)
$txtCustomPaths.Multiline = $true
$txtCustomPaths.ScrollBars = "Vertical"
$form.Controls.Add($txtCustomPaths)

$btnRun = New-Object System.Windows.Forms.Button
$btnRun.Location = New-UiPoint 15 575
$btnRun.Size = New-Object System.Drawing.Size(150,35)
$btnRun.Text = "Run Cleanup"
$form.Controls.Add($btnRun)

$btnClose = New-Object System.Windows.Forms.Button
$btnClose.Location = New-UiPoint 175 575
$btnClose.Size = New-Object System.Drawing.Size(120,35)
$btnClose.Text = "Close"
$form.Controls.Add($btnClose)

$btnOpenViewer = New-Object System.Windows.Forms.Button
$btnOpenViewer.Location = New-UiPoint 310 575
$btnOpenViewer.Size = New-Object System.Drawing.Size(170,35)
$btnOpenViewer.Text = "Open Process Viewer"
$form.Controls.Add($btnOpenViewer)

$txtStatus = New-Object System.Windows.Forms.TextBox
$txtStatus.Location = New-UiPoint 15 620
$txtStatus.Size = New-Object System.Drawing.Size(790,90)
$txtStatus.Multiline = $true
$txtStatus.ScrollBars = "Vertical"
$txtStatus.ReadOnly = $true
$form.Controls.Add($txtStatus)

$progressBar = New-Object System.Windows.Forms.ProgressBar
$progressBar.Location = New-UiPoint 15 720
$progressBar.Size = New-Object System.Drawing.Size(790,18)
$progressBar.Minimum = 0
$progressBar.Maximum = 100
$progressBar.Value = 0
$form.Controls.Add($progressBar)

$statusStrip = New-Object System.Windows.Forms.StatusStrip
$statusLabel = New-Object System.Windows.Forms.ToolStripStatusLabel
$statusLabel.Spring = $true
$statusLabel.Text = "Ready"
[void]$statusStrip.Items.Add($statusLabel)
$form.Controls.Add($statusStrip)

# =========================================
# EVENTS
# =========================================
$FeatureCheckboxes = @(
    $chkUserTemp,
    $chkWindowsTemp,
    $chkRecycleBin,
    $chkSCCM,
    $chkWindowsUpdate,
    $chkDISM,
    $chkIIS,
    $chkCustomPaths
)

$btnSelectAll.Add_Click({
    foreach ($chk in $FeatureCheckboxes) {
        $chk.Checked = $true
    }
})

$btnSelectNone.Add_Click({
    foreach ($chk in $FeatureCheckboxes) {
        $chk.Checked = $false
    }
})

$btnOpenViewer.Add_Click({
    Open-ProcessViewer
})

$btnClearViewer.Add_Click({
    $txtProcess.Clear()
})

$btnHideViewer.Add_Click({
    $processForm.Hide()
})

$btnSaveViewer.Add_Click({
    Save-ViewerLog
})

$menuToolsOpenViewer.Add_Click({
    Open-ProcessViewer
})

$menuToolsClearViewer.Add_Click({
    $txtProcess.Clear()
})

$menuToolsSaveViewer.Add_Click({
    Save-ViewerLog
})

$menuToolsTestWinRM.Add_Click({
    Test-ListedServersWinRM
})

$menuHelpContents.Add_Click({
    Show-HelpWindow -StartTab 0
})

$menuHelpFeatures.Add_Click({
    Show-HelpWindow -StartTab 1
})

$menuHelpCommands.Add_Click({
    Show-HelpWindow -StartTab 2
})

$menuHelpViewer.Add_Click({
    Open-ProcessViewer
})

$menuHelpAbout.Add_Click({
    Show-AboutDialog
})

$btnClose.Add_Click({
    $script:AllowProcessViewerClose = $true
    if (-not $processForm.IsDisposed) {
        $processForm.Close()
    }
    $form.Close()
})

$form.Add_FormClosing({
    param($sender, $e)

    $script:AllowProcessViewerClose = $true
    if ($processForm -and -not $processForm.IsDisposed) {
        $processForm.Close()
    }
})

$form.Add_KeyDown({
    param($sender, $e)

    if ($e.KeyCode -eq [System.Windows.Forms.Keys]::F1) {
        Show-HelpWindow -StartTab 0
        $e.Handled = $true
    }
})

$btnRun.Add_Click({
    $btnRun.Enabled = $false
    $txtStatus.Clear()
    $progressBar.Style = 'Continuous'
    Set-OverallProgress -Current 0 -Total 1 -LabelText "Preparing run..."

    try {
        $servers = Get-ListedServers

        if (-not $servers -or $servers.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show(
                "Please enter at least one server.",
                "Validation",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Warning
            ) | Out-Null

            $btnRun.Enabled = $true
            return
        }

        $driveLetter = $txtDrive.Text.Trim()
        if ([string]::IsNullOrWhiteSpace($driveLetter)) {
            $driveLetter = $DefaultDriveLetter
        }

        $tempAge = 0
        if (-not [int]::TryParse($txtTempAge.Text.Trim(), [ref]$tempAge)) {
            $tempAge = $DefaultTempFileAgeDays
        }

        $iisRetention = 0
        if (-not [int]::TryParse($txtIISRetention.Text.Trim(), [ref]$iisRetention)) {
            $iisRetention = $DefaultIISLogRetentionDays
        }

        $whatIfMode = $chkWhatIf.Checked
        $useAltCred = $chkAltCred.Checked
        $cleanUserTemp = $chkUserTemp.Checked
        $cleanWindowsTemp = $chkWindowsTemp.Checked
        $cleanRecycleBin = $chkRecycleBin.Checked
        $cleanSCCMCache = $chkSCCM.Checked
        $cleanWindowsUpdate = $chkWindowsUpdate.Checked
        $runDISM = $chkDISM.Checked
        $cleanIISLogs = $chkIIS.Checked
        $cleanCustomPaths = $chkCustomPaths.Checked

        if (-not ($cleanUserTemp -or $cleanWindowsTemp -or $cleanRecycleBin -or $cleanSCCMCache -or $cleanWindowsUpdate -or $runDISM -or $cleanIISLogs -or $cleanCustomPaths)) {
            [System.Windows.Forms.MessageBox]::Show(
                "Select at least one cleanup option.",
                "Validation",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Warning
            ) | Out-Null

            $btnRun.Enabled = $true
            return
        }

        $customExtraPaths = @(
            $txtCustomPaths.Lines |
            ForEach-Object { $_.Trim() } |
            Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
        )

        if ($cleanCustomPaths -and (-not $customExtraPaths -or $customExtraPaths.Count -eq 0)) {
            [System.Windows.Forms.MessageBox]::Show(
                "Custom paths is selected, but no paths were provided.",
                "Validation",
                [System.Windows.Forms.MessageBoxButtons]::OK,
                [System.Windows.Forms.MessageBoxIcon]::Warning
            ) | Out-Null

            $btnRun.Enabled = $true
            return
        }

        $credential = $null
        if ($useAltCred) {
            $credential = Get-Credential -Message "Enter credentials for remote cleanup"
            if (-not $credential) {
                Add-MainStatusLine "Credential prompt canceled."
                $btnRun.Enabled = $true
                return
            }
        }

        Open-ProcessViewer

        $totalServers = $servers.Count
        $processedServers = 0

        foreach ($server in $servers) {
            $processedServers++
            Set-OverallProgress -Current ($processedServers - 1) -Total $totalServers -LabelText ("Preparing server {0} ({1}/{2})" -f $server, $processedServers, $totalServers)

            Add-MainStatusLine "============================================================"
            Add-MainStatusLine "Processing server: $server"
            Add-ViewerLine -Message ("=" * 110) -Color "DarkGray"
            Add-ViewerLine -Message ("[{0}] [INFO] [LOCAL] Starting server: {1}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $server) -Color "Blue"

            try {
                Add-MainStatusLine "Testing DNS / name resolution..."
                Add-ViewerLine -Message ("[{0}] [STEP] [LOCAL] Testing DNS / name resolution for {1}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $server) -Color "Blue"

                [void][System.Net.Dns]::GetHostEntry($server)

                Add-MainStatusLine "DNS resolution succeeded."
                Add-ViewerLine -Message ("[{0}] [OK] [LOCAL] DNS resolution succeeded for {1}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $server) -Color "Green"
            }
            catch {
                Add-MainStatusLine ("DNS resolution warning: {0}" -f $_.Exception.Message)
                Add-ViewerLine -Message ("[{0}] [WARN] [LOCAL] DNS resolution warning for {1} - {2}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $server, $_.Exception.Message) -Color "DarkOrange"
            }

            try {
                Add-MainStatusLine "Testing WinRM..."
                Add-ViewerLine -Message ("[{0}] [STEP] [LOCAL] Testing WinRM for {1}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $server) -Color "Blue"

                Test-WSMan -ComputerName $server -ErrorAction Stop | Out-Null

                Add-MainStatusLine "WinRM connectivity succeeded."
                Add-ViewerLine -Message ("[{0}] [OK] [LOCAL] WinRM connectivity succeeded for {1}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $server) -Color "Green"
            }
            catch {
                Add-MainStatusLine ("WinRM failed: {0}" -f $_.Exception.Message)
                Add-MainStatusLine "Run on target server if needed: Enable-PSRemoting -Force"

                Add-ViewerLine -Message ("[{0}] [ERROR] [LOCAL] WinRM failed for {1} - {2}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $server, $_.Exception.Message) -Color "Red"
                Add-ViewerLine -Message ("[{0}] [INFO] [LOCAL] Run on target server if needed: Enable-PSRemoting -Force" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss")) -Color "Black"

                Set-OverallProgress -Current $processedServers -Total $totalServers -LabelText ("Skipped server {0}" -f $server)
                continue
            }

            try {
                Add-MainStatusLine "Starting remote cleanup..."
                Add-ViewerLine -Message ("[{0}] [STEP] [LOCAL] Starting remote cleanup on {1}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $server) -Color "Blue"

                $argumentList = @(
                    $driveLetter,
                    $tempAge,
                    $iisRetention,
                    $cleanUserTemp,
                    $cleanWindowsTemp,
                    $cleanRecycleBin,
                    $cleanSCCMCache,
                    $cleanWindowsUpdate,
                    $runDISM,
                    $cleanIISLogs,
                    $cleanCustomPaths,
                    $customExtraPaths,
                    $whatIfMode,
                    $DefaultOutputRoot
                )

                $job = $null
                if ($useAltCred -and $credential) {
                    $job = Invoke-Command -ComputerName $server -Credential $credential -ScriptBlock $RemoteCleanupScript -ArgumentList $argumentList -AsJob -ErrorAction Stop
                }
                else {
                    $job = Invoke-Command -ComputerName $server -ScriptBlock $RemoteCleanupScript -ArgumentList $argumentList -AsJob -ErrorAction Stop
                }

                $seenCount = 0
                $progressBar.Style = 'Marquee'

                while ($true) {
                    $jobRef = Get-Job -Id $job.Id -ErrorAction SilentlyContinue
                    if (-not $jobRef) { break }

                    $jobState = $jobRef.State
                    $allOutput = @(Receive-Job -Job $jobRef -Keep -ErrorAction SilentlyContinue)

                    if ($allOutput.Count -gt $seenCount) {
                        for ($i = $seenCount; $i -lt $allOutput.Count; $i++) {
                            $line = [string]$allOutput[$i]

                            $color = "Black"
                            if ($line -match "\[ERROR\]")      { $color = "Red" }
                            elseif ($line -match "\[WARN\]")  { $color = "DarkOrange" }
                            elseif ($line -match "\[OK\]")    { $color = "Green" }
                            elseif ($line -match "\[STEP\]")  { $color = "Blue" }
                            elseif ($line -match "\[DONE\]")  { $color = "DarkGreen" }

                            Add-ViewerLine -Message $line -Color $color
                        }

                        $seenCount = $allOutput.Count
                    }

                    if ($jobState -in @("Completed","Failed","Stopped")) {
                        break
                    }

                    Set-StatusBarText -Message ("Running cleanup on {0}..." -f $server)
                    [System.Windows.Forms.Application]::DoEvents()
                    Start-Sleep -Milliseconds 400
                }

                $finalOutput = @(Receive-Job -Job $job -Keep -ErrorAction SilentlyContinue)
                if ($finalOutput.Count -gt $seenCount) {
                    for ($i = $seenCount; $i -lt $finalOutput.Count; $i++) {
                        $line = [string]$finalOutput[$i]

                        $color = "Black"
                        if ($line -match "\[ERROR\]")      { $color = "Red" }
                        elseif ($line -match "\[WARN\]")  { $color = "DarkOrange" }
                        elseif ($line -match "\[OK\]")    { $color = "Green" }
                        elseif ($line -match "\[STEP\]")  { $color = "Blue" }
                        elseif ($line -match "\[DONE\]")  { $color = "DarkGreen" }

                        Add-ViewerLine -Message $line -Color $color
                    }
                }

                $progressBar.Style = 'Continuous'
                $jobRefFinal = Get-Job -Id $job.Id -ErrorAction SilentlyContinue
                $jobStateFinal = $null
                if ($jobRefFinal) {
                    $jobStateFinal = $jobRefFinal.State
                }

                if ($jobStateFinal -eq "Completed") {
                    Add-MainStatusLine "Cleanup finished for $server"
                    Add-MainStatusLine "Review results on remote server under: $DefaultOutputRoot"
                    Add-ViewerLine -Message ("[{0}] [DONE] [LOCAL] Completed server: {1}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $server) -Color "DarkGreen"
                }
                else {
                    $jobErrors = @()
                    if ($job.ChildJobs) {
                        $jobErrors = $job.ChildJobs | ForEach-Object { $_.Error } | Where-Object { $_ }
                    }

                    $errorText = ($jobErrors | ForEach-Object { $_.ToString() }) -join " | "
                    if ([string]::IsNullOrWhiteSpace($errorText)) {
                        $errorText = "Remote job did not complete successfully."
                    }

                    Add-MainStatusLine ("Remote cleanup failed for {0} : {1}" -f $server, $errorText)
                    Add-ViewerLine -Message ("[{0}] [ERROR] [LOCAL] FAILED server: {1} | {2}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $server, $errorText) -Color "Red"
                }

                Set-OverallProgress -Current $processedServers -Total $totalServers -LabelText ("Processed {0}/{1} servers" -f $processedServers, $totalServers)
                Remove-Job -Job $job -Force -ErrorAction SilentlyContinue
            }
            catch {
                $progressBar.Style = 'Continuous'
                Add-MainStatusLine ("Remote cleanup failed for {0} : {1}" -f $server, $_.Exception.Message)
                Add-ViewerLine -Message ("[{0}] [ERROR] [LOCAL] Remote cleanup failed for {1} - {2}" -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss"), $server, $_.Exception.Message) -Color "Red"
                Set-OverallProgress -Current $processedServers -Total $totalServers -LabelText ("Processed {0}/{1} servers" -f $processedServers, $totalServers)
            }
        }

        Add-MainStatusLine "All requested servers processed."
        Add-ViewerLine -Message ("[{0}] [DONE] [LOCAL] All requested servers processed." -f (Get-Date -Format "yyyy-MM-dd HH:mm:ss")) -Color "DarkGreen"
        Set-OverallProgress -Current $totalServers -Total $totalServers -LabelText "Run completed."
    }
    finally {
        $btnRun.Enabled = $true
        $progressBar.Style = 'Continuous'
    }
})

[void]$form.ShowDialog()
