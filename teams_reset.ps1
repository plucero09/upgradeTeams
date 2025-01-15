$logfile = "C:\Temp\teamsupgrade1.log"
Start-Transcript -Path $logfile -Append

Write-Host "Defining Variables", (Get-Date)
# Download link straight from microsoft
$dl_link = "https://go.microsoft.com/fwlink/?linkid=2243204&clcid=0x409" #https://go.microsoft.com/fwlink/?linkid=2243204&clcid=0x409
$msix_dl_link = "https://go.microsoft.com/fwlink/?linkid=2196106"  #https://go.microsoft.com/fwlink/?linkid=2196106

$tmp_installer = "C:\Temp\teamsbootstrapper.exe"
$tmp_msix = "C:\Temp\MSTeams-x64.msix"

function Start-EnsureTmpDir() {
    Write-Host "Creating temp file", (Get-Date)
    if (-not (Test-Path "C:\Temp")) {
        New-Item -ItemType Directory -Path "C:\Temp"
    }
}
function Start-EnsureResources() {
    Write-Host "Ensuring installer", (Get-Date)
    if (-not (Test-Path $tmp_installer)) {
        Invoke-WebRequest -Uri $dl_link -OutFile $tmp_installer
    }
    Write-Host "Ensuring msix", (Get-Date)
    if (-not (Test-Path $tmp_msix)) {
        Invoke-WebRequest -Uri $msix_dl_link -OutFile $tmp_msix
    }
}


function Start-EnsureTeamsInactive {
    Write-Host "Checking if Teams is open"
    # Check if Teams processes are running
    if ((Get-Process "Teams","ms-Teams" -ErrorAction SilentlyContinue | Measure-Object).Count -ne 0) {
        Get-Process "Teams","ms-Teams" -ErrorAction SilentlyContinue | Stop-Process -Force
    }

    # Check if any Teams processes are still running
    if ((Get-Process "Teams","ms-Teams" -ErrorAction SilentlyContinue | Measure-Object).Count -eq 0) {
        Write-Host "Teams processes closed successfully."
        return 0
    }

    return 0
}

function Start-CleanupTeams() {
    Write-Host "Remove Machine-Wide", (Get-Date)
    Get-Package -Type All |
        Where-Object { $_.Name -like "*Teams Machine-Wide Installer*"} |
        Uninstall-Package -Force
    # ----------------------------------------------------------------------------
    Write-Host "Remove New", (Get-Date)
    Get-AppxPackage -AllUsers |
        Where-Object { $_.Name -like '*MSTeams*'} |
        Remove-AppPackage -AllUsers
    # ----------------------------------------------------------------------------
    Write-Host "Remove Old", (Get-Date)
    Get-AppxPackage -AllUsers |
        Where-Object { $_.Name -like '*MicrosoftTeams*'} |
        Remove-AppPackage -AllUsers
    # ----------------------------------------------------------------------------
    Write-Host "Remove Leftover AppData", (Get-Date)
    Get-Item C:\Users\*\AppData\Local\Microsoft\Teams |
        Remove-Item -Recurse -Force
    # ----------------------------------------------------------------------------
    Write-Host "Remove old start menu icon", (Get-Date)
    Get-Item C:\Users\*\AppData\Local\Microsoft\Teams |
        Remove-Item -Recurse -Force
    # ----------------------------------------------------------------------------
    Write-Host "Removing start menu icon", (Get-Date)
    Remove-Item "C:\Users\*\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Microsoft Teams classic*"
    # ----------------------------------------------------------------------------
    Write-Host "Cleanup New", (Get-Date)

    & $tmp_installer -x

    # ----------------------------------------------------------------------------
    Write-Host "Cleanup Old", (Get-Date)

    & $tmp_installer -u
}
function Start-InstallTeamsNew() {
    Write-Host "Install New", (Get-Date)
    & $tmp_installer -p -o $tmp_msix
    $installs = (Get-Item "C:\Users\*\AppData\Local\Microsoft\WindowsApps\ms-teams.exe" | Measure-Object).Count
    if ($installs -eq 0) {
        throw "Failed this run due to ms-teams.exe being missing"
    }
}

$key = Get-ItemProperty HKLM:\SOFTWARE\ClearChannel\ -ErrorAction SilentlyContinue -Name "NewTeamsUpgradeRan"
if ($key) {
    Write-Host "Delete HKLM:\SOFTWARE\ClearChannel\NewTeamsUpgradeRan to run again"
    exit 0
}
Start-EnsureTmpDir
Start-EnsureResources
Start-EnsureTeamsInactive
Start-CleanupTeams
Start-InstallTeamsNew

#Create detection registry key
New-Item -ErrorAction SilentlyContinue -Path "HKLM:\SOFTWARE" -Name "ClearChannel"
New-ItemProperty -ErrorAction SilentlyContinue -Path "HKLM:\SOFTWARE\ClearChannel" -Name "NewTeamsUpgradeRan" -Value 1

Write-Host "Re-Install New", (Get-Date)
    & $tmp_installer -p -o $tmp_msix
 
#stop logging
Stop-Transcript

# msiexec.exe /i "C:\Program Files\WindowsApps\MSTeams_24243.1309.3132.617_x64__8wekyb3d8bbwe\MicrosoftTeamsMeetingAddinInstaller.msi" ALLUSERS=1 /qn /norestart TARGETDIR="C:\Program Files (x86)\Microsoft\TeamsMeetingAdd-in\24243.1309.3132.617\"