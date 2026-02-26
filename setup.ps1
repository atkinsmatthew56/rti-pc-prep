# ============================================================
# RTI PC Prep Script
# Prepares a Windows PC for RTI control system integration
# ============================================================

param(
    [string]$ClientName = "",
    [string]$SlackWebhook = ""
)

# --- Load Slack webhook from local config if not passed as parameter ---
$configFile = "C:\RTIListener\rti_config.txt"
if (-not $SlackWebhook) {
    if (Test-Path $configFile) {
        $SlackWebhook = (Get-Content $configFile | Where-Object { $_ -match "^SLACK_WEBHOOK=" }) -replace "^SLACK_WEBHOOK=", ""
    }
}

# --- Prompt for client name if not provided ---
if (-not $ClientName) {
    $ClientName = Read-Host "Enter client/job name (e.g. Smith Residence)"
}

$LogPath = "C:\RTIListener\prep_log.txt"
$ConfigPath = "C:\RTIListener\config.txt"
$ListenerScript = "C:\RTIListener\listener.py"
$VBSScript = "C:\RTIListener\start_listener.vbs"
$PCName = $env:COMPUTERNAME
$Username = $env:USERNAME

function Log {
    param([string]$msg)
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $line = "[$timestamp] $msg"
    Write-Host $line
    Add-Content -Path $LogPath -Value $line
}

function Step {
    param([string]$msg)
    Write-Host ""
    Write-Host "========================================" -ForegroundColor Cyan
    Write-Host "  $msg" -ForegroundColor Cyan
    Write-Host "========================================" -ForegroundColor Cyan
}

# --- Ensure running as administrator ---
if (-NOT ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]"Administrator")) {
    Write-Host "ERROR: This script must be run as Administrator." -ForegroundColor Red
    Write-Host "Right-click PowerShell and select 'Run as Administrator'" -ForegroundColor Yellow
    pause
    exit 1
}

# --- Create RTIListener folder ---
Step "Creating RTIListener folder"
if (-not (Test-Path "C:\RTIListener")) {
    New-Item -ItemType Directory -Path "C:\RTIListener" | Out-Null
    Log "Created C:\RTIListener"
} else {
    Log "C:\RTIListener already exists"
}

# --- Install Python if missing ---
Step "Checking Python installation"
$pythonPath = $null
$possiblePaths = @(
    "$env:LOCALAPPDATA\Python\bin\python.exe",
    "$env:LOCALAPPDATA\Programs\Python\Python312\python.exe",
    "$env:LOCALAPPDATA\Programs\Python\Python311\python.exe",
    "$env:LOCALAPPDATA\Programs\Python\Python310\python.exe",
    "C:\Python312\python.exe",
    "C:\Python311\python.exe",
    "C:\Python310\python.exe"
)

foreach ($path in $possiblePaths) {
    if (Test-Path $path) {
        $pythonPath = $path
        break
    }
}

# Also try PATH
if (-not $pythonPath) {
    try {
        $found = (Get-Command python -ErrorAction SilentlyContinue).Source
        if ($found -and $found -notlike "*WindowsApps*") {
            $pythonPath = $found
        }
    } catch {}
}

if ($pythonPath) {
    Log "Python found at: $pythonPath"
} else {
    Log "Python not found. Installing via winget..."
    try {
        winget install Python.Python.3.12 --silent --accept-package-agreements --accept-source-agreements
        Start-Sleep -Seconds 15
        $env:Path = [System.Environment]::GetEnvironmentVariable("Path","Machine") + ";" + [System.Environment]::GetEnvironmentVariable("Path","User")
        foreach ($path in $possiblePaths) {
            if (Test-Path $path) {
                $pythonPath = $path
                break
            }
        }
        if ($pythonPath) {
            Log "Python installed successfully at: $pythonPath"
        } else {
            Log "WARNING: Python install may need manual verification"
            $pythonPath = "python"
        }
    } catch {
        Log "ERROR: Could not install Python automatically. Please install manually from python.org"
    }
}

# --- Write listener.py ---
Step "Writing RTI Listener script"
$listenerCode = @'
from http.server import BaseHTTPRequestHandler, HTTPServer
import subprocess

class ShutdownHandler(BaseHTTPRequestHandler):
    def do_GET(self):
        if self.path == '/shutdown':
            self.send_response(200)
            self.end_headers()
            self.wfile.write(b'Shutting down...')
            subprocess.Popen(['shutdown', '/s', '/t', '5'])
        elif self.path == '/status':
            self.send_response(200)
            self.end_headers()
            self.wfile.write(b'Online')
        else:
            self.send_response(404)
            self.end_headers()

    def log_message(self, format, *args):
        pass

if __name__ == '__main__':
    server = HTTPServer(('0.0.0.0', 9100), ShutdownHandler)
    print('RTI Listener running on port 9100...')
    server.serve_forever()
'@

Set-Content -Path $ListenerScript -Value $listenerCode -Encoding UTF8
Log "listener.py written to $ListenerScript"

# --- Write silent VBS launcher ---
Step "Writing silent VBS launcher"
$vbsCode = "Set WshShell = CreateObject(`"WScript.Shell`")`r`nWshShell.Run `"$pythonPath C:\RTIListener\listener.py`", 0, False"
Set-Content -Path $VBSScript -Value $vbsCode -Encoding ASCII
Log "start_listener.vbs written to $VBSScript"

# --- Create scheduled task ---
Step "Creating startup scheduled task"
schtasks /delete /tn "RTIListener" /f 2>$null
$result = schtasks /create /tn "RTIListener" /tr "wscript.exe C:\RTIListener\start_listener.vbs" /sc onlogon /ru $Username /rl HIGHEST /f 2>&1
Log "Scheduled task result: $result"

# --- Configure NIC Wake on LAN ---
# Filters out virtual/Hyper-V/Trackman adapters - only configures physical NICs
Step "Configuring network adapter Wake on LAN"
$adapters = Get-NetAdapter | Where-Object {
    $_.Status -eq "Up" -and
    $_.PhysicalMediaType -ne "Wireless LAN" -and
    $_.InterfaceDescription -notmatch "Hyper-V|Virtual|vEthernet|Loopback|Bluetooth|WAN Miniport|Trackman"
}
foreach ($adapter in $adapters) {
    try {
        Log "Configured WoL on adapter: $($adapter.Name)"
    } catch {
        Log "WARNING: Could not configure WoL on $($adapter.Name): $_"
    }
}

# Enable WoL via PowerShell Power Management cmdlet
# Also explicitly sets both Device Manager WoL checkboxes:
# - "Allow this device to wake the computer"
# - "Only allow a magic packet to wake the computer"
try {
    $adapters | ForEach-Object {
        Enable-NetAdapterPowerManagement -Name $_.Name -WakeOnMagicPacket -WakeOnPattern -ErrorAction SilentlyContinue
        Set-NetAdapterAdvancedProperty -Name $_.Name -RegistryKeyword "*WakeOnMagicPacket" -RegistryValue 1 -ErrorAction SilentlyContinue
        Set-NetAdapterAdvancedProperty -Name $_.Name -RegistryKeyword "*WakeOnPattern" -RegistryValue 1 -ErrorAction SilentlyContinue
        Set-NetAdapterAdvancedProperty -Name $_.Name -RegistryKeyword "*WakeOnLink" -RegistryValue 1 -ErrorAction SilentlyContinue
        Log "Power management WoL enabled on: $($_.Name)"
    }
} catch {
    Log "NOTE: Enable-NetAdapterPowerManagement not available on this system, WoL set via BIOS"
}

# --- Force-enable 'Allow this device to wake the computer' via WMI ---
Step "Enabling WoL in Device Manager Power Management (critical)"
try {
    $wmiNics = Get-WmiObject -Namespace root\wmi -Class MSPower_DeviceWakeEnable
    $armed = $false
    foreach ($wmiNic in $wmiNics) {
        if ($wmiNic.InstanceName -match "PCI") {
            $wmiNic.Enable = $true
            $wmiNic.Put() | Out-Null
            Log "WMI wake enable set on: $($wmiNic.InstanceName)"
            $armed = $true
        }
    }
    if ($armed) {
        Log "SUCCESS: WMI wake enable applied to PCI network adapters"
    } else {
        Log "WARNING: No PCI network adapters found via WMI"
    }
} catch {
    Log "WARNING: Could not set WMI wake enable: $_"
}

# --- Verify NIC is armed for WoL ---
Step "Verifying Wake on LAN is armed"
Start-Sleep -Seconds 2
$wakeArmed = & powercfg /devicequery wake_armed
$wakeArmedStr = $wakeArmed -join ", "
Log "Wake-armed devices: $wakeArmedStr"

if ($wakeArmedStr -match "Ethernet|LAN|Intel|Realtek|I225|I226") {
    Write-Host ""
    Write-Host "  [OK] Ethernet adapter is armed for Wake on LAN" -ForegroundColor Green
    Write-Host ""
    Log "SUCCESS: WoL verification passed - Ethernet adapter is armed"
} else {
    Write-Host ""
    Write-Host "  [WARNING] Ethernet adapter not found in wake_armed list" -ForegroundColor Yellow
    Write-Host ""
    Write-Host "  ACTION REQUIRED after reboot:" -ForegroundColor Yellow
    Write-Host "  1. Open Device Manager" -ForegroundColor Yellow
    Write-Host "  2. Expand Network Adapters" -ForegroundColor Yellow
    Write-Host "  3. Right-click your Ethernet adapter > Properties" -ForegroundColor Yellow
    Write-Host "  4. Click Power Management tab" -ForegroundColor Yellow
    Write-Host "  5. Check 'Allow this device to wake the computer'" -ForegroundColor Yellow
    Write-Host "  6. Check 'Only allow a magic packet to wake the computer'" -ForegroundColor Yellow
    Write-Host "  7. Click OK" -ForegroundColor Yellow
    Write-Host ""
    Log "WARNING: WoL verification failed - manual Device Manager step required"
}

# --- Disable Sleep and Hibernate ---
Step "Configuring power settings"
powercfg /change standby-timeout-ac 0
powercfg /change standby-timeout-dc 0
powercfg /change hibernate-timeout-ac 0
powercfg /change hibernate-timeout-dc 0
powercfg /change monitor-timeout-ac 0
powercfg /change monitor-timeout-dc 0
powercfg /hibernate off

# Set High Performance power plan
$highPerf = powercfg /list | Select-String "High performance"
if ($highPerf) {
    $guid = ($highPerf -split "\s+")[3]
    powercfg /setactive $guid
    Log "Power plan set to High Performance: $guid"
} else {
    Log "High Performance plan not found, creating it"
    powercfg /duplicatescheme 8c5e7fda-e8bf-4a96-9a85-a6e23a8c635c
}
Log "Sleep and hibernate disabled, High Performance power plan activated"

# --- Open Firewall Port 9100 ---
Step "Configuring Windows Firewall"
netsh advfirewall firewall delete rule name="RTI Listener Port 9100" 2>$null
netsh advfirewall firewall add rule name="RTI Listener Port 9100" dir=in action=allow protocol=TCP localport=9100
Log "Firewall rule added for port 9100"

# --- Disable UAC ---
Step "Disabling UAC"
Set-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Policies\System" -Name "EnableLUA" -Value 0
Log "UAC disabled"

# --- Disable Windows Update automatic restarts ---
Step "Disabling Windows Update auto-restart"
$wuPath = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate\AU"
if (-not (Test-Path $wuPath)) {
    New-Item -Path $wuPath -Force | Out-Null
}
Set-ItemProperty -Path $wuPath -Name "NoAutoRebootWithLoggedOnUsers" -Value 1 -Type DWord
Set-ItemProperty -Path $wuPath -Name "AUOptions" -Value 3 -Type DWord
Log "Windows Update auto-restart disabled"

# --- Get MAC and IP (physical adapters only - excludes Hyper-V, virtual, Trackman) ---
Step "Collecting network information"
$nic = Get-NetAdapter | Where-Object {
    $_.Status -eq "Up" -and
    $_.PhysicalMediaType -ne "Wireless LAN" -and
    $_.InterfaceDescription -notmatch "Hyper-V|Virtual|vEthernet|Loopback|Bluetooth|WAN Miniport|Trackman"
} | Sort-Object Speed -Descending | Select-Object -First 1

$macAddress = $nic.MacAddress
$ipInfo = Get-NetIPAddress -InterfaceIndex $nic.ifIndex -AddressFamily IPv4 | Select-Object -First 1
$ipAddress = $ipInfo.IPAddress
$broadcastAddress = ($ipAddress -replace "\.\d+$", ".255")

Log "MAC Address: $macAddress"
Log "IP Address: $ipAddress"
Log "Broadcast Address: $broadcastAddress"

# --- Write config file ---
Step "Writing config summary"
$configContent = @"
==========================================
RTI PC CONFIGURATION SUMMARY
==========================================
Client:           $ClientName
PC Name:          $PCName
Date Configured:  $(Get-Date -Format "yyyy-MM-dd HH:mm:ss")
------------------------------------------
NETWORK
MAC Address:      $macAddress
IP Address:       $ipAddress
Broadcast:        $broadcastAddress
Listener Port:    9100
------------------------------------------
RTI COMMANDS
Status Check:     http://$ipAddress:9100/status
Shutdown:         http://$ipAddress:9100/shutdown
WoL Broadcast:    $broadcastAddress (UDP port 9)
------------------------------------------
FILES
Listener Script:  C:\RTIListener\listener.py
VBS Launcher:     C:\RTIListener\start_listener.vbs
Log File:         C:\RTIListener\prep_log.txt
------------------------------------------
SETTINGS APPLIED
- PCIe WoL:       Enabled (verify in BIOS)
- NIC WoL:        Configured (magic packet + pattern)
- NIC Wake Armed: Configured via WMI
- Sleep:          Disabled
- Hibernate:      Disabled
- Power Plan:     High Performance
- Firewall 9100:  Open
- UAC:            Disabled
- WU Auto-reboot: Disabled
- Listener Task:  Scheduled at logon
==========================================
"@

Set-Content -Path $ConfigPath -Value $configContent -Encoding UTF8
Log "Config written to $ConfigPath"

# --- Send Slack notification ---
Step "Sending Slack notification"
if ($SlackWebhook) {
    $slackMessage = @{
        text = "*RTI PC Configuration Complete*"
        attachments = @(
            @{
                color = "good"
                fields = @(
                    @{ title = "Client"; value = $ClientName; short = $true }
                    @{ title = "PC Name"; value = $PCName; short = $true }
                    @{ title = "MAC Address"; value = $macAddress; short = $true }
                    @{ title = "IP Address"; value = $ipAddress; short = $true }
                    @{ title = "Listener Port"; value = "9100"; short = $true }
                    @{ title = "Status Endpoint"; value = "http://$ipAddress:9100/status"; short = $false }
                    @{ title = "Configured"; value = (Get-Date -Format "yyyy-MM-dd HH:mm:ss"); short = $true }
                )
            }
        )
    } | ConvertTo-Json -Depth 5

    try {
        Invoke-RestMethod -Uri $SlackWebhook -Method Post -Body $slackMessage -ContentType "application/json"
        Log "Slack notification sent successfully"
    } catch {
        Log "WARNING: Could not send Slack notification: $_"
    }
} else {
    Log "No Slack webhook provided - skipping notification"
}

# --- Done ---
Write-Host ""
Write-Host "========================================" -ForegroundColor Green
Write-Host "  RTI PC PREP COMPLETE" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Green
Write-Host ""
Write-Host "Client:      $ClientName" -ForegroundColor White
Write-Host "MAC Address: $macAddress" -ForegroundColor White
Write-Host "IP Address:  $ipAddress" -ForegroundColor White
Write-Host "Port:        9100" -ForegroundColor White
Write-Host ""
Write-Host "Config saved to: $ConfigPath" -ForegroundColor Yellow
Write-Host "Log saved to:    $LogPath" -ForegroundColor Yellow
Write-Host ""
Write-Host "IMPORTANT: Please reboot the PC to activate all settings." -ForegroundColor Cyan
Write-Host ""
pause
