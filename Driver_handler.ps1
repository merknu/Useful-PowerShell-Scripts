##############################################################################
# 1. Check for Administrator Rights
##############################################################################
if (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole(
    [Security.Principal.WindowsBuiltinRole] "Administrator")) {
    [System.Windows.Forms.MessageBox]::Show(
        "Please run this script as Administrator.",
        "Insufficient Privileges",
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Error
    )
    exit
}

##############################################################################
# 2. Load Assemblies and Enable Visual Styles
##############################################################################
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

##############################################################################
# 3. Global Logging Variable (for CSV/XML exports)
##############################################################################
$Global:LogEntries = @()

##############################################################################
# 4. Function: Write-Log
# Logs a message (with timestamp) to the GUI log textbox and appends it
# to CSV/XML logs. All messages are stored in $Global:LogEntries.
##############################################################################
function Write-Log {
    param(
        [string]$Message,
        [switch]$IsError
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $entry = [PSCustomObject]@{
        Timestamp = $timestamp
        Message   = $Message
        IsError   = [bool]$IsError
    }
    $Global:LogEntries += $entry

    # Safely update GUI log on the main thread
    $form.Invoke([Action]{
        $LogTextbox.AppendText("[$timestamp] $Message`r`n")
    })
    
    # Append to CSV log (creates the file if missing)
    $csvPath = ".\DriverRepairLog.csv"
    if (-not (Test-Path $csvPath)) {
        $entry | Export-Csv -Path $csvPath -NoTypeInformation
    } else {
        $entry | Export-Csv -Path $csvPath -NoTypeInformation -Append
    }
    
    # Overwrite an XML log snapshot each time
    $xmlPath = ".\DriverRepairLog.xml"
    $Global:LogEntries | Export-Clixml -Path $xmlPath
}

##############################################################################
# 5. Function: Get-ProblemDevices
# Returns devices that are not in an "OK" state. Optionally, use -PresentOnly.
##############################################################################
function Get-ProblemDevices {
    [CmdletBinding()]
    param(
        [switch]$PresentOnly
    )
    Write-Verbose "Scanning for problematic devices..."
    try {
        $devices = Get-PnpDevice -PresentOnly:$PresentOnly | Where-Object { $_.Status -ne "OK" }
        return $devices
    }
    catch {
        Write-Log "Error retrieving devices: $($_.Exception.Message)" -IsError
        return @()
    }
}

##############################################################################
# 6. Build the GUI
##############################################################################
$form = New-Object System.Windows.Forms.Form
$form.Text = "Driver Repair Tool"
$form.Size = New-Object System.Drawing.Size(800,600)
$form.StartPosition = 'CenterScreen'
$form.FormBorderStyle = 'FixedDialog'
$form.MaximizeBox = $false

# Log TextBox: Displays log messages (read-only)
$LogTextbox = New-Object System.Windows.Forms.TextBox
$LogTextbox.Multiline = $true
$LogTextbox.ScrollBars = "Vertical"
$LogTextbox.Size = New-Object System.Drawing.Size(760,200)
$LogTextbox.Location = New-Object System.Drawing.Point(20,20)
$LogTextbox.ReadOnly = $true
$LogTextbox.Font = New-Object System.Drawing.Font("Consolas",10)
$form.Controls.Add($LogTextbox)

# ListBox: Displays problematic devices
$DeviceListBox = New-Object System.Windows.Forms.ListBox
$DeviceListBox.Size = New-Object System.Drawing.Size(760,150)
$DeviceListBox.Location = New-Object System.Drawing.Point(20,230)
$form.Controls.Add($DeviceListBox)
$DeviceListBox.DisplayMember = "FriendlyName"

# Button: Scan for Problem Devices
$ScanButton = New-Object System.Windows.Forms.Button
$ScanButton.Text = "Scan Devices"
$ScanButton.Size = New-Object System.Drawing.Size(120,40)
$ScanButton.Location = New-Object System.Drawing.Point(20,400)
$form.Controls.Add($ScanButton)

# Button: Repair All Devices
$RepairAllButton = New-Object System.Windows.Forms.Button
$RepairAllButton.Text = "Repair All"
$RepairAllButton.Size = New-Object System.Drawing.Size(120,40)
$RepairAllButton.Location = New-Object System.Drawing.Point(160,400)
$form.Controls.Add($RepairAllButton)

# Button: Clear Log
$ClearLogButton = New-Object System.Windows.Forms.Button
$ClearLogButton.Text = "Clear Log"
$ClearLogButton.Size = New-Object System.Drawing.Size(120,40)
$ClearLogButton.Location = New-Object System.Drawing.Point(300,400)
$form.Controls.Add($ClearLogButton)

# Button: Export Log (to CSV and XML)
$ExportLogButton = New-Object System.Windows.Forms.Button
$ExportLogButton.Text = "Export Log"
$ExportLogButton.Size = New-Object System.Drawing.Size(120,40)
$ExportLogButton.Location = New-Object System.Drawing.Point(440,400)
$form.Controls.Add($ExportLogButton)

##############################################################################
# 7. Event Handlers for GUI Buttons
##############################################################################

# (A) Scan Button: Clears list, scans for devices, populates the ListBox
$ScanButton.Add_Click({
    $DeviceListBox.Items.Clear()
    Write-Log "Scanning for problematic devices..."
    $problemDevices = Get-ProblemDevices -PresentOnly
    if ($problemDevices.Count -eq 0) {
        Write-Log "No problematic devices found."
    }
    else {
        foreach ($dev in $problemDevices) {
            $DeviceListBox.Items.Add($dev)
        }
        Write-Log "Found $($problemDevices.Count) problematic device(s)."
    }
})

# (B) Repair All Button: Repairs all devices in the ListBox using a background job.
# After attempting repairs, re-scans each device and logs whether it is fixed.
$RepairAllButton.Add_Click({
    if ($DeviceListBox.Items.Count -eq 0) {
        Write-Log "No devices to repair. Please run a scan first." -IsError
        return
    }

    Write-Log "Starting repair on all devices..."
    $devicesToRepair = @()
    foreach ($item in $DeviceListBox.Items) {
        $devicesToRepair += $item
    }
    
    try {
        # Start a background job with inline definition of Repair-Device and import of PnPDevice module.
        $job = Start-Job -ScriptBlock {
            param($devices)
            
            # Import the PnPDevice module inside the job
            Import-Module PnpDevice -ErrorAction Stop
            
            # Define the Repair-Device function inside the job
            function Repair-Device {
                param([string]$InstanceId)
                try {
                    Write-Output "Disabling device ${InstanceId}..."
                    Disable-PnpDevice -InstanceId ${InstanceId} -Confirm:$false -ErrorAction Stop
                    Start-Sleep -Seconds 3

                    Write-Output "Enabling device ${InstanceId}..."
                    Enable-PnpDevice -InstanceId ${InstanceId} -Confirm:$false -ErrorAction Stop
                    Start-Sleep -Seconds 3

                    $currentStatus = (Get-PnpDevice -InstanceId ${InstanceId}).Status
                    Write-Output "Status after re-enable: $currentStatus"

                    if ($currentStatus -ne "OK") {
                        Write-Output "Device ${InstanceId} still problematic. Attempting driver update..."
                        Update-PnpDevice -InstanceId ${InstanceId} -ErrorAction Stop
                        Start-Sleep -Seconds 5
                        $finalStatus = (Get-PnpDevice -InstanceId ${InstanceId}).Status
                        Write-Output "Status after driver update: $finalStatus"
                        if ($finalStatus -eq "OK") {
                            Write-Output "Device ${InstanceId} repaired successfully."
                            return $true
                        }
                        else {
                            Write-Output "Failed to repair device ${InstanceId}."
                            return $false
                        }
                    }
                    else {
                        Write-Output "Device ${InstanceId} repaired successfully."
                        return $true
                    }
                }
                catch {
                    Write-Output "Error repairing device ${InstanceId}: $($_.Exception.Message)"
                    return $false
                }
            }
            
            $results = @()
            # Attempt repair for each device
            foreach ($dev in $devices) {
                $outcome = Repair-Device -InstanceId $dev.InstanceId
                if ($outcome) {
                    $results += "Device $($dev.FriendlyName) repaired successfully."
                }
                else {
                    $results += "Device $($dev.FriendlyName) repair attempt failed."
                }
            }
            
            # After repairs, re-scan each device to confirm final status
            foreach ($dev in $devices) {
                try {
                    $updated = Get-PnpDevice -InstanceId $dev.InstanceId
                    if ($updated.Status -eq "OK") {
                        $results += "Final check: Device $($dev.FriendlyName) is fixed."
                    }
                    else {
                        $results += "Final check: Could not fix device $($dev.FriendlyName). Status: $($updated.Status)"
                    }
                }
                catch {
                    $results += "Final check error for device $($dev.FriendlyName): $($_.Exception.Message)"
                }
            }
            return $results
        } -ArgumentList ($devicesToRepair)
    }
    catch {
        Write-Log "Failed to start background job: $($_.Exception.Message)" -IsError
        return
    }
    
    # Wait for the job to actually finish while keeping the GUI responsive
    while ($job.State -in 'Running','NotStarted','Blocked') {
        [System.Windows.Forms.Application]::DoEvents()
        Start-Sleep -Milliseconds 100
    }
    
    # Process the job result based on its state
    switch ($job.State) {
        'Completed' {
            $jobResults = Receive-Job -Job $job -ErrorAction SilentlyContinue
            foreach ($line in $jobResults) {
                Write-Log $line
            }
            Remove-Job -Job $job -Force
            Write-Log "Repair process completed."
        }
        'Failed' {
            Write-Log "Job failed. Checking errors..." -IsError
            $jobError = $job.ChildJobs[0].Error
            if ($jobError) {
                foreach ($err in $jobError) {
                    Write-Log "Error in job: $($err.Exception.Message)" -IsError
                }
            }
            Remove-Job -Job $job -Force
            Write-Log "Repair process ended with errors."
        }
        'Stopped' {
            Write-Log "Job was stopped manually." -IsError
            Remove-Job -Job $job -Force
        }
        default {
            Write-Log "Repair job ended unexpectedly with state: $($job.State)." -IsError
            Remove-Job -Job $job -Force
        }
    }
})

# (C) Clear Log Button: Clears the log TextBox (GUI only)
$ClearLogButton.Add_Click({
    $LogTextbox.Clear()
})

# (D) Export Log Button: Exports the global log entries to CSV and XML with timestamped filenames
$ExportLogButton.Add_Click({
    $csvPath = ".\DriverRepairLog_$(Get-Date -Format 'yyyyMMdd_HHmmss').csv"
    $xmlPath = ".\DriverRepairLog_$(Get-Date -Format 'yyyyMMdd_HHmmss').xml"
    $Global:LogEntries | Export-Csv -Path $csvPath -NoTypeInformation
    $Global:LogEntries | Export-Clixml -Path $xmlPath
    Write-Log "Log exported to $csvPath and $xmlPath"
})

##############################################################################
# 8. Run the GUI
##############################################################################
$form.Add_Shown({ $form.Activate() })
[System.Windows.Forms.Application]::Run($form)
