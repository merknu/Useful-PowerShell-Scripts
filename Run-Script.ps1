# Run-Script.ps1
# Interactive menu to run scripts from this collection

param(
    [string]$ScriptName
)

# ASCII Art Banner
$banner = @"
╔═══════════════════════════════════════════════════════════════╗
║        Useful PowerShell Scripts Collection                   ║
║        By: merknu                                             ║
╚═══════════════════════════════════════════════════════════════╝
"@

Write-Host $banner -ForegroundColor Cyan
Write-Host ""

# Define available scripts
$scripts = @{
    "1" = @{
        "Name" = "Dynamic Network Monitor"
        "File" = ".\Dynamic_ping.ps1"
        "Description" = "Monitor network connectivity with real-time charts"
        "RequiresAdmin" = $false
    }
    "2" = @{
        "Name" = "IT Maintenance Tool"
        "File" = ".\Automated_IT_Maintenance_Tool.ps1"
        "Description" = "Comprehensive system maintenance and optimization"
        "RequiresAdmin" = $true
    }
    "3" = @{
        "Name" = "Driver Handler"
        "File" = ".\Driver_handler.ps1"
        "Description" = "Manage and update system drivers"
        "RequiresAdmin" = $true
    }
    "4" = @{
        "Name" = "APA Style Installer"
        "File" = ".\APA_Style_installer.ps1"
        "Description" = "Install and configure APA citation style"
        "RequiresAdmin" = $false
    }
    "5" = @{
        "Name" = "CCNA Quiz System"
        "File" = ".\Quizer\Start_Quiz.ps1"
        "Description" = "Interactive CCNA exam preparation quiz"
        "RequiresAdmin" = $false
    }
}

# Function to check if running as administrator
function Test-Administrator {
    $currentUser = [Security.Principal.WindowsIdentity]::GetCurrent()
    $principal = New-Object Security.Principal.WindowsPrincipal($currentUser)
    return $principal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
}

# Function to display menu
function Show-Menu {
    Write-Host "Available Scripts:" -ForegroundColor Yellow
    Write-Host "═════════════════════════════════════════════════════════════" -ForegroundColor DarkGray
    
    foreach ($key in ($scripts.Keys | Sort-Object)) {
        $script = $scripts[$key]
        $adminTag = if ($script.RequiresAdmin) { "[ADMIN]" } else { "" }
        Write-Host "$key. $($script.Name) $adminTag" -ForegroundColor Green
        Write-Host "   $($script.Description)" -ForegroundColor Gray
    }
    
    Write-Host ""
    Write-Host "Q. Quit" -ForegroundColor Red
    Write-Host "═════════════════════════════════════════════════════════════" -ForegroundColor DarkGray
}

# Function to run selected script
function Run-SelectedScript {
    param($selection)
    
    if ($scripts.ContainsKey($selection)) {
        $script = $scripts[$selection]
        
        Write-Host ""
        Write-Host "Selected: $($script.Name)" -ForegroundColor Cyan
        
        # Check if admin required
        if ($script.RequiresAdmin -and -not (Test-Administrator)) {
            Write-Host "This script requires administrator privileges." -ForegroundColor Yellow
            $response = Read-Host "Would you like to restart with admin rights? (Y/N)"
            
            if ($response -eq 'Y' -or $response -eq 'y') {
                # Restart script as admin
                Start-Process powershell.exe -ArgumentList "-File", $script.File -Verb RunAs
                return
            } else {
                Write-Host "Script cancelled." -ForegroundColor Red
                return
            }
        }
        
        # Check if script file exists
        if (Test-Path $script.File) {
            Write-Host "Starting $($script.Name)..." -ForegroundColor Green
            Write-Host ""
            
            # Run the script
            & $script.File
            
            Write-Host ""
            Write-Host "Script completed." -ForegroundColor Green
        } else {
            Write-Host "Error: Script file not found at $($script.File)" -ForegroundColor Red
        }
    } else {
        Write-Host "Invalid selection." -ForegroundColor Red
    }
}

# Main program loop
if ($ScriptName) {
    # If script name provided as parameter, try to find and run it
    $found = $false
    foreach ($key in $scripts.Keys) {
        if ($scripts[$key].Name -like "*$ScriptName*" -or $scripts[$key].File -like "*$ScriptName*") {
            Run-SelectedScript $key
            $found = $true
            break
        }
    }
    if (-not $found) {
        Write-Host "Script '$ScriptName' not found." -ForegroundColor Red
    }
} else {
    # Interactive menu
    do {
        Show-Menu
        Write-Host ""
        $selection = Read-Host "Enter your selection"
        
        if ($selection -ne 'Q' -and $selection -ne 'q') {
            Run-SelectedScript $selection
            Write-Host ""
            Read-Host "Press Enter to continue..."
            Clear-Host
            Write-Host $banner -ForegroundColor Cyan
            Write-Host ""
        }
    } while ($selection -ne 'Q' -and $selection -ne 'q')
    
    Write-Host "Thank you for using Useful PowerShell Scripts!" -ForegroundColor Cyan
}