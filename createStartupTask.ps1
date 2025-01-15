# Define task parameters
$TaskName = "RemoveTeamsTask"
$ScriptPath = "C:\Windows\removeTeamsScript.ps1"  # Path to the startup script
$ExpireDate = "2025-03-30T23:59:59"

# Create the startup script
$ScriptContent = @'
# Get the Teams Machine-Wide Installer
$installer = Get-WmiObject -Class Win32_Product | Where-Object { $_.Name -like "Teams Machine-Wide Installer" }

# Uninstall the Teams Machine-Wide Installer if it exists
if ($installer) {
    $installer.Uninstall()
    Write-Output "Teams Machine-Wide Installer has been uninstalled."
} else {
    Write-Output "Teams Machine-Wide Installer is not installed."
}

# Function to load the user hive
function Load-UserHive {
    param (
        [string]$profilePath,
        [string]$hiveName
    )
    $ntUserDat = "$profilePath\NTUSER.DAT"
    if (-Not (Test-Path $ntUserDat)) {
        Write-Host "NTUSER.DAT file not found at $profilePath"
        return
    }
    reg load "HKU\$hiveName" $ntUserDat
}

# Function to delete the Teams registry key
function Remove-TeamsKey {
    param (
        [string]$hiveName
    )
    $teamsKey = "Registry::" + "HKU\$hiveName\Software\Microsoft\Windows\CurrentVersion\Uninstall\Teams"
    write-host "Searching for Key:" $teamsKey
    if (Test-Path $teamsKey) {
        Remove-Item -Path $teamsKey -Recurse -Force
        Write-Host "Teams key removed"
    } else {
        Write-Host "Teams key not found"
    }
}

# Function to unload the user hive
function Unload-UserHive {
    param (
        [string]$hiveName
    )
    write-host "Unload HIV:" $hiveName
    [gc]::Collect()
    Start-Sleep 5
    reg unload "HKU\$hiveName"
    start-sleep 2
    
    # Check if the hive is unloaded
    if (-Not (Test-Path ("Registry::" + "HKU\Test"))) 
    {
        Write-Host "Hive $hiveName successfully unloaded."
    } else 
    {
        Write-Host "Failed to unload hive $hiveName."
    }
}

# Main script

#list all users Profiles in c:\users
$ProfileNames = (Get-ChildItem -Path C:\Users -Directory | Where-Object { $_.Name -ne "Public" }).Name
foreach($Profile in $ProfileNames)
{
    #content to search and remove
    $TeamsRoot = "C:\Users\$Profile\AppData\Local\Microsoft\Teams"
    $TeamsEXE = "C:\Users\$Profile\AppData\Local\Microsoft\Teams\Update.exe" #current\Teams.exe"
    #$TeamsREG = "HKEY_USERS\$Profile\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\Teams"
    
    #Attempt Uninstall Teams
    if(Test-Path -Path $TeamsEXE)
    {
        
        Set-Location $TeamsRoot
        write-host "Found: " $TeamsRoot "Attempting Removal" -ForegroundColor Green
        Start-Process ".\Update.exe" -ArgumentList "--uninstall -s"
        $TeamsUpdateEXE = $TeamsRoot + "\Update.exe"
        if(Test-Path -Path $TeamsUpdateEXE){Write-host "    Failed to Uninstall" $TeamsUpdateEXE -ForegroundColor Red}
        if(!(Test-Path -Path $TeamsUpdateEXE))
        {
            Write-host "    Uninstalled Successfully" $TeamsUpdateEXE -ForegroundColor Green
            remove-item $TeamsRoot -recurse -Force
        }

    }


    $profilePath = "C:\Users\$Profile"
    $hiveName = "Test"

    write-host "Load User Hive for: " $profilePath
    Load-UserHive -profilePath $profilePath -hiveName $hiveName
    Remove-TeamsKey -hiveName $hiveName
    write-host "UnLoad User Hive"
    Unload-UserHive -hiveName $hiveName
    write-host "`r`n"
}
'@
Set-Content -Path $ScriptPath -Value $ScriptContent -Force

# Create the action for the task
$Action = New-ScheduledTaskAction -Execute "powershell.exe" -Argument "-NoProfile -ExecutionPolicy Bypass -File `"$ScriptPath`""

# Create the trigger to run the task at system startup
$Trigger = New-ScheduledTaskTrigger -AtStartup

# Set the expiration date for the trigger
$Trigger.EndBoundary = $ExpireDate

# Define the settings for the task
$Settings = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -StartWhenAvailable

# Define compatibility for Windows 10
$Principal = New-ScheduledTaskPrincipal -UserId "SYSTEM" -LogonType ServiceAccount -RunLevel Highest

# Register the scheduled task to run as SYSTEM
Register-ScheduledTask -TaskName $TaskName -Action $Action -Trigger $Trigger -Settings $Settings -Principal $Principal

Write-Output "Scheduled task '$TaskName' created successfully, and startup script saved to '$ScriptPath'."

