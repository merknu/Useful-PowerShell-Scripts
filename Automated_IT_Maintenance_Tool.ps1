##############################################################################
# Load necessary assemblies for Windows Forms
##############################################################################
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

##############################################################################
# Create the Form
##############################################################################
$form = New-Object System.Windows.Forms.Form
$form.Text = "Automated IT Maintenance Tool"
$form.Size = New-Object System.Drawing.Size(900, 750)
$form.StartPosition = 'CenterScreen'

##############################################################################
# Create a TextBox for displaying output logs
##############################################################################
$outputBox = New-Object System.Windows.Forms.TextBox
$outputBox.Multiline = $true
$outputBox.ScrollBars = "Vertical"
$outputBox.Size = New-Object System.Drawing.Size(850, 200)
$outputBox.Location = New-Object System.Drawing.Point(20, 520)
$outputBox.ReadOnly = $true
$form.Controls.Add($outputBox)

##############################################################################
# Function to update the output box safely on the UI thread
##############################################################################
function Update-Output {
    param ($message)
    $form.Invoke([Action]{
        $outputBox.AppendText("$message`r`n")
    })
}

##############################################################################
# ROW 1: Basic & Advanced System Tasks
##############################################################################
$btnY1 = 20
$btnWidth = 120
$btnHeight = 40

# Disk Cleanup Button
$diskCleanupBtn = New-Object System.Windows.Forms.Button
$diskCleanupBtn.Text = "Disk Cleanup"
$diskCleanupBtn.Size = New-Object System.Drawing.Size($btnWidth, $btnHeight)
$diskCleanupBtn.Location = New-Object System.Drawing.Point(20, $btnY1)
$form.Controls.Add($diskCleanupBtn)
$diskCleanupBtn.Add_Click({
    Update-Output "Starting Disk Cleanup..."
    try {
        Start-Process -FilePath "cleanmgr.exe" -ArgumentList "/sagerun:1" -NoNewWindow -Wait
        Update-Output "Disk Cleanup completed."
    } catch {
        Update-Output "Error during Disk Cleanup: $_"
    }
})

# System Info Button
$systemInfoBtn = New-Object System.Windows.Forms.Button
$systemInfoBtn.Text = "System Info"
$systemInfoBtn.Size = New-Object System.Drawing.Size($btnWidth, $btnHeight)
$systemInfoBtn.Location = New-Object System.Drawing.Point(150, $btnY1)
$form.Controls.Add($systemInfoBtn)
$systemInfoBtn.Add_Click({
    Update-Output "Gathering System Information..."
    try {
        $sysInfo = Get-ComputerInfo | Select-Object `
            WindowsProductName, WindowsVersion, OsArchitecture, CsName,
            CsManufacturer, CsModel, CsProcessors, OsLocale, CsTotalPhysicalMemory,
            BiosManufacturer, BiosVersion, BiosReleaseDate, OsInstallDate, OsLastBootUpTime, OsUptime,
            OsTotalVisibleMemorySize, OsFreePhysicalMemory, OsCountryCode, TimeZone

        $formattedInfo = @"
==========================
 System Information
==========================
- Computer Name       : $($sysInfo.CsName)
- Operating System    : $($sysInfo.WindowsProductName) $($sysInfo.OsArchitecture)
- OS Version          : $($sysInfo.WindowsVersion)
- Language/Locale     : $($sysInfo.OsLocale)
- Country Code        : $($sysInfo.OsCountryCode)

==========================
 Hardware Information
==========================
- Manufacturer        : $($sysInfo.CsManufacturer)
- Model               : $($sysInfo.CsModel)
- Processor           : $($sysInfo.CsProcessors)
- Total Memory (RAM)  : $([math]::round($sysInfo.CsTotalPhysicalMemory / 1GB, 2)) GB

==========================
 BIOS Information
==========================
- BIOS Manufacturer   : $($sysInfo.BiosManufacturer)
- BIOS Version        : $($sysInfo.BiosVersion)
- BIOS Release Date   : $($sysInfo.BiosReleaseDate)

==========================
 System Status
==========================
- OS Install Date     : $($sysInfo.OsInstallDate)
- Last Boot Up Time   : $($sysInfo.OsLastBootUpTime)
- System Uptime       : $($sysInfo.OsUptime.Days) days, $($sysInfo.OsUptime.Hours) hrs
- Total Physical Memory: $([math]::round($sysInfo.OsTotalVisibleMemorySize / 1MB, 2)) MB
- Free Physical Memory: $([math]::round($sysInfo.OsFreePhysicalMemory / 1MB, 2)) MB
- Time Zone           : $($sysInfo.TimeZone)
"@
        Update-Output $formattedInfo
    } catch {
        Update-Output "Error gathering system info: $_"
    }
})

# PC Hardware Info Button
$pcInfoBtn = New-Object System.Windows.Forms.Button
$pcInfoBtn.Text = "PC Hardware Info"
$pcInfoBtn.Size = New-Object System.Drawing.Size($btnWidth, $btnHeight)
$pcInfoBtn.Location = New-Object System.Drawing.Point(280, $btnY1)
$form.Controls.Add($pcInfoBtn)
$pcInfoBtn.Add_Click({
    Update-Output "Gathering PC Hardware Info..."
    try {
        $cpu = Get-WmiObject -Class Win32_Processor -ErrorAction Stop | 
               Select-Object Name, NumberOfCores, NumberOfLogicalProcessors, MaxClockSpeed
        $gpuList = Get-WmiObject -Class Win32_VideoController -ErrorAction Stop | 
                   Select-Object Name, AdapterRAM, DriverVersion
        $ram = (Get-WmiObject -Class Win32_ComputerSystem -ErrorAction Stop).TotalPhysicalMemory
        $windowsVersion = (Get-ComputerInfo -ErrorAction Stop).WindowsVersion

        $cpuInfo = "CPU Name: $($cpu.Name) (Cores: $($cpu.NumberOfCores), Threads: $($cpu.NumberOfLogicalProcessors), Clock: $($cpu.MaxClockSpeed) MHz)"
        $gpuInfo = "GPU(s):"
        foreach ($g in $gpuList) {
            $gpuInfo += "`r`n   - Name: $($g.Name), Driver: $($g.DriverVersion), Memory: $([math]::Round($g.AdapterRAM/1MB,2)) MB"
        }
        $ramInfo = "Total RAM: $([math]::Round($ram/1GB,2)) GB"
        $winVersionInfo = "Windows Version: $windowsVersion"
        $formattedHardwareInfo = @"
==========================
 PC Hardware Info
==========================
$cpuInfo
$gpuInfo
$ramInfo
$winVersionInfo
"@
        Update-Output $formattedHardwareInfo
    } catch {
        Update-Output "Error gathering PC Hardware Info: $_"
    }
})

# Optimize Disk Button
$optimizeDiskBtn = New-Object System.Windows.Forms.Button
$optimizeDiskBtn.Text = "Optimize Disk"
$optimizeDiskBtn.Size = New-Object System.Drawing.Size($btnWidth, $btnHeight)
$optimizeDiskBtn.Location = New-Object System.Drawing.Point(410, $btnY1)
$form.Controls.Add($optimizeDiskBtn)
$optimizeDiskBtn.Add_Click({
    Update-Output "Optimizing Disk C:..."
    try {
        Optimize-Volume -DriveLetter C -Defrag -Verbose
        Update-Output "Disk optimization completed."
    } catch {
        Update-Output "Error optimizing disk: $_"
    }
})

# Ping Test Button (for vg.no)
$pingTestBtn = New-Object System.Windows.Forms.Button
$pingTestBtn.Text = "Ping vg.no"
$pingTestBtn.Size = New-Object System.Drawing.Size($btnWidth, $btnHeight)
$pingTestBtn.Location = New-Object System.Drawing.Point(540, $btnY1)
$form.Controls.Add($pingTestBtn)
$pingTestBtn.Add_Click({
    Update-Output "Performing Ping Test on vg.no..."
    $target = "vg.no"
    try {
        $result = Test-Connection -ComputerName $target -Count 4 -ErrorAction SilentlyContinue
        if ($result) {
            $avg = [math]::Round(($result | Measure-Object -Property ResponseTime -Average).Average,2)
            Update-Output "Ping to ${target}: Average = $avg ms"
        } else {
            Update-Output "Ping to ${target} failed."
        }
    } catch {
        Update-Output "Error performing Ping Test: $_"
    }
})

# Registry Cleanup Button (example – customize as needed)
$registryCleanupBtn = New-Object System.Windows.Forms.Button
$registryCleanupBtn.Text = "Registry Cleanup"
$registryCleanupBtn.Size = New-Object System.Drawing.Size($btnWidth, $btnHeight)
$registryCleanupBtn.Location = New-Object System.Drawing.Point(670, $btnY1)
$form.Controls.Add($registryCleanupBtn)
$registryCleanupBtn.Add_Click({
    Update-Output "Performing Registry Cleanup..."
    try {
        $regPath = "HKCU:\Software\TempCleanup"
        if (Test-Path $regPath) {
            Remove-Item $regPath -Recurse -Force
            Update-Output "Obsolete registry entries removed from $regPath."
        } else {
            Update-Output "No obsolete registry entries found at $regPath."
        }
    } catch {
        Update-Output "Error during Registry Cleanup: $_"
    }
})

##############################################################################
# ROW 2: Windows Updates Section
##############################################################################
$updateListBox = New-Object System.Windows.Forms.ListBox
$updateListBox.SelectionMode = 'MultiExtended'
$updateListBox.Size = New-Object System.Drawing.Size(550, 100)
$updateListBox.Location = New-Object System.Drawing.Point(20, 70)
$form.Controls.Add($updateListBox)

$updateCheckBtn = New-Object System.Windows.Forms.Button
$updateCheckBtn.Text = "Check for Updates"
$updateCheckBtn.Size = New-Object System.Drawing.Size(150, 30)
$updateCheckBtn.Location = New-Object System.Drawing.Point(590, 70)
$form.Controls.Add($updateCheckBtn)

$installAllBtn = New-Object System.Windows.Forms.Button
$installAllBtn.Text = "Install All Updates"
$installAllBtn.Size = New-Object System.Drawing.Size(150, 30)
$installAllBtn.Location = New-Object System.Drawing.Point(590, 110)
$form.Controls.Add($installAllBtn)

$installSelectedBtn = New-Object System.Windows.Forms.Button
$installSelectedBtn.Text = "Install Selected Updates"
$installSelectedBtn.Size = New-Object System.Drawing.Size(150, 30)
$installSelectedBtn.Location = New-Object System.Drawing.Point(590, 150)
$form.Controls.Add($installSelectedBtn)

$updateCheckBtn.Add_Click({
    Update-Output "Checking for Windows Updates..."
    if (!(Get-Module -ListAvailable -Name PSWindowsUpdate)) {
        try {
            Update-Output "Installing PSWindowsUpdate module..."
            Install-Module -Name PSWindowsUpdate -Force -Scope CurrentUser
            Update-Output "PSWindowsUpdate module installed successfully."
        } catch {
            Update-Output "Error installing PSWindowsUpdate module: $_"
            return
        }
    }
    try {
        Import-Module PSWindowsUpdate -ErrorAction Stop
        $updateListBox.Items.Clear()
        $availableUpdates = Get-WindowsUpdate -ErrorAction Stop
        if ($availableUpdates) {
            foreach ($update in $availableUpdates) {
                $updateListBox.Items.Add("$($update.Title) - $($update.KBArticleIDs)")
            }
            Update-Output "Updates available."
        } else {
            Update-Output "No updates available."
        }
    } catch {
        Update-Output "Error checking for updates: $_"
    }
})

$installAllBtn.Add_Click({
    try {
        Import-Module PSWindowsUpdate -ErrorAction Stop
        Update-Output "Installing all updates..."
        Install-WindowsUpdate -AcceptAll -AutoReboot -ErrorAction Stop
        Update-Output "All updates installed successfully."
    } catch {
        Update-Output "Error installing updates: $_"
    }
})

$installSelectedBtn.Add_Click({
    try {
        Import-Module PSWindowsUpdate -ErrorAction Stop
        $selectedUpdates = $updateListBox.SelectedItems
        if ($selectedUpdates.Count -eq 0) {
            Update-Output "No updates selected."
            return
        }
        foreach ($updateItem in $selectedUpdates) {
            $updateTitle = $updateItem.Split(" - ")[0]
            $updateToInstall = Get-WindowsUpdate | Where-Object { $_.Title -eq $updateTitle }
            if ($updateToInstall) {
                Install-WindowsUpdate -KBArticleID $updateToInstall.KBArticleIDs -AcceptAll -AutoReboot -ErrorAction Stop
                Update-Output "Installed update: $updateTitle"
            } else {
                Update-Output "Update $updateTitle not found."
            }
        }
        Update-Output "Selected updates installed successfully."
    } catch {
        Update-Output "Error installing selected updates: $_"
    }
})

##############################################################################
# ROW 3: Maintenance Tools
##############################################################################
$btnY3 = 190

# Clear Temp Files
$clearTempBtn = New-Object System.Windows.Forms.Button
$clearTempBtn.Text = "Clear Temp Files"
$clearTempBtn.Size = New-Object System.Drawing.Size($btnWidth, $btnHeight)
$clearTempBtn.Location = New-Object System.Drawing.Point(20, $btnY3)
$form.Controls.Add($clearTempBtn)
$clearTempBtn.Add_Click({
    Update-Output "Clearing temporary files..."
    try {
        $tempPath = $env:TEMP
        if (-Not (Test-Path $tempPath)) {
            throw "Temporary path not found: $tempPath"
        }
        Get-ChildItem -Path $tempPath -Recurse -Force -ErrorAction SilentlyContinue | Remove-Item -Force -Recurse -ErrorAction SilentlyContinue
        Update-Output "Temporary files cleared successfully."
    } catch {
        Update-Output "Error clearing temporary files: $_"
    }
})

# Run SFC
$runSfcBtn = New-Object System.Windows.Forms.Button
$runSfcBtn.Text = "Run SFC"
$runSfcBtn.Size = New-Object System.Drawing.Size($btnWidth, $btnHeight)
$runSfcBtn.Location = New-Object System.Drawing.Point(150, $btnY3)
$form.Controls.Add($runSfcBtn)
$runSfcBtn.Add_Click({
    Update-Output "Running SFC..."
    try {
        Start-Process -FilePath "sfc.exe" -ArgumentList "/scannow" -NoNewWindow -Wait
        Update-Output "SFC completed."
    } catch {
        Update-Output "Error running SFC: $_"
    }
})

# Memory Diagnostic
$memoryDiagBtn = New-Object System.Windows.Forms.Button
$memoryDiagBtn.Text = "Memory Diagnostic"
$memoryDiagBtn.Size = New-Object System.Drawing.Size($btnWidth, $btnHeight)
$memoryDiagBtn.Location = New-Object System.Drawing.Point(280, $btnY3)
$form.Controls.Add($memoryDiagBtn)
$memoryDiagBtn.Add_Click({
    Update-Output "Starting Memory Diagnostic..."
    try {
        Start-Process -FilePath "mdsched.exe" -NoNewWindow -Wait
        Update-Output "Memory Diagnostic launched."
    } catch {
        Update-Output "Error launching Memory Diagnostic: $_"
    }
})

# Event Viewer
$eventViewerBtn = New-Object System.Windows.Forms.Button
$eventViewerBtn.Text = "Event Viewer"
$eventViewerBtn.Size = New-Object System.Drawing.Size($btnWidth, $btnHeight)
$eventViewerBtn.Location = New-Object System.Drawing.Point(410, $btnY3)
$form.Controls.Add($eventViewerBtn)
$eventViewerBtn.Add_Click({
    Update-Output "Opening Event Viewer..."
    try {
        Start-Process -FilePath "eventvwr.exe" -NoNewWindow -Wait
        Update-Output "Event Viewer opened."
    } catch {
        Update-Output "Error opening Event Viewer: $_"
    }
})

##############################################################################
# ROW 4: Disk Tasks
##############################################################################
$btnY4 = 240

# Defrag Disk
$defragBtn = New-Object System.Windows.Forms.Button
$defragBtn.Text = "Defrag Disk"
$defragBtn.Size = New-Object System.Drawing.Size($btnWidth, $btnHeight)
$defragBtn.Location = New-Object System.Drawing.Point(20, $btnY4)
$form.Controls.Add($defragBtn)
$defragBtn.Add_Click({
    Update-Output "Starting Disk Defragmentation..."
    try {
        Start-Process -FilePath "defrag.exe" -ArgumentList "C:" -NoNewWindow -Wait
        Update-Output "Disk defragmentation completed."
    } catch {
        Update-Output "Error defragmenting disk: $_"
    }
})

# Check Disk
$chkdskBtn = New-Object System.Windows.Forms.Button
$chkdskBtn.Text = "Check Disk"
$chkdskBtn.Size = New-Object System.Drawing.Size($btnWidth, $btnHeight)
$chkdskBtn.Location = New-Object System.Drawing.Point(150, $btnY4)
$form.Controls.Add($chkdskBtn)
$chkdskBtn.Add_Click({
    Update-Output "Running CHKDSK on C:..."
    try {
        Start-Process -FilePath "chkdsk.exe" -ArgumentList "C:" -NoNewWindow -Wait
        Update-Output "CHKDSK completed."
    } catch {
        Update-Output "Error running CHKDSK: $_"
    }
})

##############################################################################
# ROW 5: Security, Scheduling, Remote & Multi-threaded Tests
##############################################################################
$btnY5 = 290

# Security Check
$securityCheckBtn = New-Object System.Windows.Forms.Button
$securityCheckBtn.Text = "Security Check"
$securityCheckBtn.Size = New-Object System.Drawing.Size($btnWidth, $btnHeight)
$securityCheckBtn.Location = New-Object System.Drawing.Point(20, $btnY5)
$form.Controls.Add($securityCheckBtn)
$securityCheckBtn.Add_Click({
    Update-Output "Performing Security Check..."
    try {
        $firewallProfiles = Get-NetFirewallProfile | Select-Object Name, Enabled
        foreach ($profile in $firewallProfiles) {
            Update-Output "Firewall Profile $($profile.Name): Enabled = $($profile.Enabled)"
        }
    } catch {
        Update-Output "Error checking Firewall: $_"
    }
    try {
        if (Get-Command Get-MpComputerStatus -ErrorAction SilentlyContinue) {
            $defenderStatus = Get-MpComputerStatus
            Update-Output "Defender Enabled: $($defenderStatus.AntivirusEnabled)"
            Update-Output "Defender Signature Updated: $($defenderStatus.AVSignatureLastUpdated)"
        } else {
            Update-Output "Windows Defender cmdlet not available."
        }
    } catch {
        Update-Output "Error checking Windows Defender: $_"
    }
})

# Schedule Maintenance Button
$scheduleMaintBtn = New-Object System.Windows.Forms.Button
$scheduleMaintBtn.Text = "Schedule Maint."
$scheduleMaintBtn.Size = New-Object System.Drawing.Size($btnWidth, $btnHeight)
$scheduleMaintBtn.Location = New-Object System.Drawing.Point(150, $btnY5)
$form.Controls.Add($scheduleMaintBtn)
$scheduleMaintBtn.Add_Click({
    Update-Output "Scheduling maintenance task..."
    try {
        $trigger = New-ScheduledTaskTrigger -Once -At ((Get-Date).AddDays(1).Date.AddHours(2))
        $action = New-ScheduledTaskAction -Execute "PowerShell.exe" -Argument "-File `"`"C:\Path\To\YourScript.ps1`"`" -NoProfile -WindowStyle Hidden"
        Register-ScheduledTask -TaskName "AutomatedITMaintenance" -Trigger $trigger -Action $action -RunLevel Highest -Force
        Update-Output "Maintenance task scheduled for tomorrow at 2:00 AM."
    } catch {
        Update-Output "Error scheduling maintenance: $_"
    }
})

# Remote Check Button
$remoteCheckBtn = New-Object System.Windows.Forms.Button
$remoteCheckBtn.Text = "Remote Check"
$remoteCheckBtn.Size = New-Object System.Drawing.Size($btnWidth, $btnHeight)
$remoteCheckBtn.Location = New-Object System.Drawing.Point(280, $btnY5)
$form.Controls.Add($remoteCheckBtn)
$remoteCheckBtn.Add_Click({
    Update-Output "Starting remote check..."
    $remoteComputers = @("RemotePC1", "RemotePC2")  # Replace with actual names
    try {
        $results = Invoke-Command -ComputerName $remoteComputers -ScriptBlock {
            $uptime = (Get-CimInstance Win32_OperatingSystem).LastBootUpTime
            return "Uptime: " + (New-TimeSpan -Start $uptime).ToString("dd\.hh\:mm\:ss")
        } -ErrorAction SilentlyContinue
        foreach ($res in $results) {
            Update-Output "Remote: $res"
        }
    } catch {
        Update-Output "Error during remote check: $_"
    }
})

# Multi-threaded Test Button (demonstration)
$multiThreadBtn = New-Object System.Windows.Forms.Button
$multiThreadBtn.Text = "Multi-thread Test"
$multiThreadBtn.Size = New-Object System.Drawing.Size($btnWidth, $btnHeight)
$multiThreadBtn.Location = New-Object System.Drawing.Point(410, $btnY5)
$form.Controls.Add($multiThreadBtn)
$multiThreadBtn.Add_Click({
    Update-Output "Starting multi-threaded test..."
    $jobList = @()
    for ($i = 1; $i -le 3; $i++) {
        $jobList += Start-Job -ScriptBlock {
            param($num)
            Start-Sleep -Seconds (Get-Random -Minimum 1 -Maximum 3)
            return "Job $num completed after sleep."
        } -ArgumentList $i
    }
    $jobList | Wait-Job
    foreach ($job in $jobList) {
        $result = Receive-Job -Job $job
        Update-Output $result
    }
    $jobList | Remove-Job
})

##############################################################################
# ROW 6: ALL-IN-ONE MAINTENANCE TASKS
##############################################################################
$btnY6 = 350
$allInOneBtn = New-Object System.Windows.Forms.Button
$allInOneBtn.Text = "Run All Maintenance"
$allInOneBtn.Size = New-Object System.Drawing.Size(180, 40)
$allInOneBtn.Location = New-Object System.Drawing.Point(20, $btnY6)
$form.Controls.Add($allInOneBtn)
$allInOneBtn.Add_Click({
    Update-Output "Starting All Maintenance Tasks..."

    # Disk Cleanup
    Update-Output "Running Disk Cleanup..."
    try {
        Start-Process -FilePath "cleanmgr.exe" -ArgumentList "/sagerun:1" -NoNewWindow -Wait
        Update-Output "Disk Cleanup completed."
    } catch {
        Update-Output "Error during Disk Cleanup: $_"
    }
    
    # Optimize Disk
    Update-Output "Optimizing Disk C:..."
    try {
        Optimize-Volume -DriveLetter C -Defrag -Verbose
        Update-Output "Disk optimization completed."
    } catch {
        Update-Output "Error optimizing disk: $_"
    }
    
    # Clear Temp Files
    Update-Output "Clearing Temporary Files..."
    try {
        $tempPath = $env:TEMP
        if (-Not (Test-Path $tempPath)) {
            throw "Temporary path not found: $tempPath"
        }
        Get-ChildItem -Path $tempPath -Recurse -Force -ErrorAction SilentlyContinue | Remove-Item -Force -Recurse -ErrorAction SilentlyContinue
        Update-Output "Temporary files cleared."
    } catch {
        Update-Output "Error clearing temporary files: $_"
    }
    
    # Run SFC
    Update-Output "Running System File Checker (SFC)..."
    try {
        Start-Process -FilePath "sfc.exe" -ArgumentList "/scannow" -NoNewWindow -Wait
        Update-Output "SFC scan completed."
    } catch {
        Update-Output "Error running SFC: $_"
    }
    
    # Run CHKDSK
    Update-Output "Running CHKDSK on C:..."
    try {
        Start-Process -FilePath "chkdsk.exe" -ArgumentList "C:" -NoNewWindow -Wait
        Update-Output "CHKDSK completed."
    } catch {
        Update-Output "Error running CHKDSK: $_"
    }
    
    # Security Check
    Update-Output "Performing Security Check..."
    try {
        $firewallProfiles = Get-NetFirewallProfile | Select-Object Name, Enabled
        foreach ($profile in $firewallProfiles) {
            Update-Output "Firewall Profile $($profile.Name): Enabled = $($profile.Enabled)"
        }
    } catch {
        Update-Output "Error checking Firewall: $_"
    }
    try {
        if (Get-Command Get-MpComputerStatus -ErrorAction SilentlyContinue) {
            $defenderStatus = Get-MpComputerStatus
            Update-Output "Defender Enabled: $($defenderStatus.AntivirusEnabled)"
            Update-Output "Defender Signature Updated: $($defenderStatus.AVSignatureLastUpdated)"
        } else {
            Update-Output "Windows Defender cmdlet not available."
        }
    } catch {
        Update-Output "Error checking Windows Defender: $_"
    }
    
    # Ping Test
    Update-Output "Performing Ping Test on vg.no..."
    $target = "vg.no"
    try {
        $result = Test-Connection -ComputerName $target -Count 4 -ErrorAction SilentlyContinue
        if ($result) {
            $avg = [math]::Round(($result | Measure-Object -Property ResponseTime -Average).Average,2)
            Update-Output "Ping to ${target}: Average = $avg ms"
        } else {
            Update-Output "Ping to ${target} failed."
        }
    } catch {
        Update-Output "Error performing Ping Test: $_"
    }
    
    Update-Output "All maintenance tasks completed."
})

##############################################################################
# Show the Form
##############################################################################
$form.ShowDialog()
