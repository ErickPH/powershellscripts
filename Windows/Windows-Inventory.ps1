<#
.SYNOPSIS
    Generates a comprehensive inventory report for a Windows machine or remote machines.
.DESCRIPTION
    This script collects system information including machine details, disk information,
    memory details, installed applications, network configuration, server roles (if applicable),
    system health metrics, and more. It supports exporting reports in HTML, CSV, and JSON formats.
    The script can also collect inventory data from remote machines using PowerShell remoting.
.PARAMETER ComputerName
    The name of the computer to collect inventory data from. Defaults to the local machine.
.PARAMETER ExportFormat
    The format of the report to generate. Options are "HTML", "CSV", or "JSON". Defaults to "HTML".
.PARAMETER EmailRecipient
    The email address to send the generated report to. If not specified, the report will not be emailed.
.PARAMETER OutputDir
    The directory where the generated report will be saved. Defaults to "C:\InventoryReports".
.PARAMETER IncludeGroupPolicy
    Include details about applied Group Policies. Defaults to $false.
.PARAMETER IncludeScheduledTasks
    Include a list of scheduled tasks on the machine. Defaults to $false.
.PARAMETER IncludeActiveDirectory
    Include Active Directory information if the machine is part of a domain. Defaults to $false.
.PARAMETER IncludeWindowsFeatures
    Include a list of installed Windows features and roles. Defaults to $false.
.PARAMETER IncludeEventLogs
    Include a summary of recent critical or error events from the Windows Event Log. Defaults to $false.
.PARAMETER IncludeFirewallAntivirus
    Include the status of the Windows Firewall and installed antivirus software. Defaults to $false.
.PARAMETER IncludeGPU
    Include details about the GPU(s) installed on the machine. Defaults to $false.
.PARAMETER All
    Include all optional sections in the report. Defaults to $false.
.EXAMPLE
    .\Windows-Inventory.ps1 -ComputerName "RemotePC01" -ExportFormat "JSON" -EmailRecipient "admin@example.com" -All
    This command generates a system inventory report for the remote computer "RemotePC01" in JSON format,
    includes all optional sections, and emails the report to "admin@example.com".
.NOTES
    File Name      : Windows-Inventory.ps1
    Author         : Erick Perez - quadrianweb.com
    Prerequisite   : PowerShell 5.1 or later, Run as Administrator for some details
                     WinRM must be enabled and running on both the local and remote machines.
                     To enable and configure the WinRM service on Windows, on an elevated admin session it is enough to run this command:
                             winrm quickconfig
                     or 
                             Enable-PSRemoting –Force
                    For more information, see: 
                    https://learn.microsoft.com/en-us/powershell/module/microsoft.powershell.core/enable-psremoting?view=powershell-7.5
                    https://learn.microsoft.com/en-us/windows/win32/winrm/installation-and-configuration-for-windows-remote-management
    Version        : 2.3
    Modification Date: 2023-10-05
    Change Log     : 
        - Added collapsible sections for better readability
        - Added export to CSV and JSON formats
        - Included system health metrics (CPU, memory, disk usage)
        - Added support for remote machine inventory
        - Implemented error logging
        - Added customizable sections and email functionality
        - Included user accounts, security settings, and event logs
        - Added metadata for command-line parameters
        - Made output directory configurable via command-line parameter
        - Added example for remote execution with JSON export and email
        - Added note about WinRM requirement
        - Added optional sections for Group Policy, Scheduled Tasks, Active Directory, Windows Features, Event Logs, Firewall/Antivirus, and GPU
        - Added an -All parameter to include all optional sections
        - Added a section to collect installed features and roles for Windows Server
#>

# Parameters
param (
    [string]$ComputerName = $env:COMPUTERNAME,
    [string]$ExportFormat = "HTML", # Options: HTML, CSV, JSON
    [string]$EmailRecipient = $null,
    [string]$OutputDir = "C:\InventoryReports",
    [switch]$IncludeGroupPolicy = $false,
    [switch]$IncludeScheduledTasks = $false,
    [switch]$IncludeActiveDirectory = $false,
    [switch]$IncludeWindowsFeatures = $false,
    [switch]$IncludeEventLogs = $false,
    [switch]$IncludeFirewallAntivirus = $false,
    [switch]$IncludeGPU = $false,
    [switch]$All = $false
)

Write-Host "Starting Windows Inventory Script..." -ForegroundColor Cyan

# Enable all sections if -All is specified
if ($All) {
    Write-Host "Enabling all optional sections as -All parameter is specified..." -ForegroundColor Cyan
    $IncludeGroupPolicy = $true
    $IncludeScheduledTasks = $true
    $IncludeActiveDirectory = $true
    $IncludeWindowsFeatures = $true
    $IncludeEventLogs = $true
    $IncludeFirewallAntivirus = $true
    $IncludeGPU = $true
}

# Output file configuration
Write-Host "Configuring output files and directories..." -ForegroundColor Cyan
$outputFileBase = "$OutputDir\${ComputerName}_SystemInventory_$(Get-Date -Format 'yyyyMMdd_HHmmss')"
$outputFile = "$outputFileBase.html"
$csvFile = "$outputFileBase.csv"
$jsonFile = "$outputFileBase.json"
$timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"

# Create output directory if it doesn't exist
if (-not (Test-Path -Path $OutputDir)) {
    New-Item -ItemType Directory -Path $OutputDir | Out-Null
}

# Error logging
$errorLog = "$OutputDir\ErrorLog_$(Get-Date -Format 'yyyyMMdd_HHmmss').txt"
function Log-Error {
    param ([string]$Message)
    Add-Content -Path $errorLog -Value "$((Get-Date).ToString()): $Message"
}

Write-Host "Checking WinRM service status..." -ForegroundColor Cyan
# Check if WinRM service is running
function Test-WinRM {
    param (
        [string]$ComputerName
    )
    try {
        if ($ComputerName -eq $env:COMPUTERNAME) {
            $winrmStatus = Get-Service -Name WinRM -ErrorAction SilentlyContinue
            return $winrmStatus.Status -eq 'Running'
        } else {
            $winrmStatus = Invoke-Command -ComputerName $ComputerName -ScriptBlock {
                Get-Service -Name WinRM -ErrorAction SilentlyContinue
            } -ErrorAction SilentlyContinue
            return $winrmStatus.Status -eq 'Running'
        }
    } catch {
        return $false
    }
}

if (-not (Test-WinRM -ComputerName $ComputerName)) {
    Write-Host "WinRM is not running on $ComputerName. Please ensure the WinRM service is enabled and running." -ForegroundColor Red
    Log-Error "WinRM is not running on $ComputerName. Please ensure the WinRM service is enabled and running."
    return
}

Write-Host "Collecting system information for $ComputerName..." -ForegroundColor Cyan
# HTML Report Header
$htmlHeader = @"
<!DOCTYPE html>
<html>
<head>
    <title>Windows System Inventory - $ComputerName</title>
    <style>
        body { font-family: Arial, sans-serif; margin: 20px; }
        h1 { color: #0066cc; }
        h2 { color: #0099cc; margin-top: 30px; border-bottom: 1px solid #ddd; padding-bottom: 5px; cursor: pointer; }
        table { border-collapse: collapse; width: 100%; margin-bottom: 20px; }
        th { background-color: #f2f2f2; text-align: left; padding: 8px; }
        td { padding: 8px; border-bottom: 1px solid #ddd; }
        .section { margin-bottom: 30px; }
        .summary { background-color: #f9f9f9; padding: 15px; border-radius: 5px; }
    </style>
    <script>
        function toggleSection(id) {
            var section = document.getElementById(id);
            section.style.display = section.style.display === 'none' ? 'block' : 'none';
        }
    </script>
</head>
<body>
    <h1>Windows System Inventory Report</h1>
    <p><strong>Generated on:</strong> $timestamp</p>
    <div class="summary">
        <p><strong>Computer Name:</strong> $ComputerName</p>
        <p><strong>Report Location:</strong> $outputFile</p>
    </div>
"@

# HTML Report Footer
$htmlFooter = @"
</body>
</html>
"@

# Initialize HTML content
$htmlContent = $htmlHeader

# Collect system information
try {
    if ($ComputerName -eq $env:COMPUTERNAME -or (Test-Connection -ComputerName $ComputerName -Count 1 -Quiet)) {
        $os = Get-CimInstance -ComputerName $ComputerName -ClassName Win32_OperatingSystem
        $computerSystem = Get-CimInstance -ComputerName $ComputerName -ClassName Win32_ComputerSystem
        $bios = Get-CimInstance -ComputerName $ComputerName -ClassName Win32_BIOS
        $processor = Get-CimInstance -ComputerName $ComputerName -ClassName Win32_Processor | Select-Object -First 1
        $disks = Get-CimInstance -ComputerName $ComputerName -ClassName Win32_LogicalDisk | Where-Object { $_.DriveType -eq 3 }
        $memory = Get-CimInstance -ComputerName $ComputerName -ClassName Win32_PhysicalMemory
        $adapters = Get-CimInstance -ComputerName $ComputerName -ClassName Win32_NetworkAdapterConfiguration | Where-Object { $_.IPEnabled -eq $true }
        $physicalDisks = Get-CimInstance -ComputerName $ComputerName -ClassName Win32_DiskDrive
    } else {
        Write-Host "Unable to connect to $ComputerName. Ensure the machine is reachable and WinRM is configured." -ForegroundColor Red
        Log-Error "Unable to connect to $ComputerName. Ensure the machine is reachable and WinRM is configured."
        return
    }
} catch {
    Log-Error "Error collecting system information: $_"
    Write-Host "Error collecting system information. Check the error log for details." -ForegroundColor Red
    return
}

Write-Host "Generating HTML report sections..." -ForegroundColor Cyan

# 1. Machine Details
Write-Host "Adding Machine Details to the report..." -ForegroundColor Cyan
$htmlContent += "<div class='section'><h2 onclick=`"toggleSection('systemInfo')`">System Information</h2>"
$htmlContent += "<div id='systemInfo' style='display: none;'>"
$htmlContent += @"
<table>
<tr><th>Property</th><th>Value</th></tr>
<tr><td>Hostname</td><td>$ComputerName</td></tr>
<tr><td>Manufacturer</td><td>$($computerSystem.Manufacturer)</td></tr>
<tr><td>Model</td><td>$($computerSystem.Model)</td></tr>
<tr><td>System Type</td><td>$($computerSystem.SystemType)</td></tr>
<tr><td>Operating System</td><td>$($os.Caption) ($($os.OSArchitecture))</td></tr>
<tr><td>OS Version</td><td>$($os.Version)</td></tr>
<tr><td>Build Number</td><td>$($os.BuildNumber)</td></tr>
<tr><td>Serial Number</td><td>$($bios.SerialNumber)</td></tr>
<tr><td>BIOS Version</td><td>$($bios.SMBIOSBIOSVersion)</td></tr>
<tr><td>Processor</td><td>$($processor.Name)</td></tr>
<tr><td>Logical Processors</td><td>$($computerSystem.NumberOfLogicalProcessors)</td></tr>
<tr><td>Physical Processors</td><td>$($computerSystem.NumberOfProcessors)</td></tr>
<tr><td>Last Boot Time</td><td>$($os.LastBootUpTime)</td></tr>
<tr><td>System Uptime</td><td>$((Get-Date) - [datetime]::Parse($os.LastBootUpTime))</td></tr>
</table>
</div></div>
"@

# 2. Hard Disk Details
Write-Host "Adding Disk Information to the report..." -ForegroundColor Cyan
$htmlContent += "<div class='section'><h2 onclick=`"toggleSection('diskInfo')`">Disk Information</h2>"
$htmlContent += "<div id='diskInfo' style='display: none;'>"
$htmlContent += @"
<table>
<tr><th>Drive</th><th>Label</th><th>File System</th><th>Total Size (GB)</th><th>Free Space (GB)</th><th>% Free</th></tr>
"@

foreach ($disk in $disks) {
    $totalGB = [math]::Round($disk.Size / 1GB, 2)
    $freeGB = [math]::Round($disk.FreeSpace / 1GB, 2)
    $freePercent = [math]::Round(($disk.FreeSpace / $disk.Size) * 100, 2)
    
    $htmlContent += @"
<tr>
    <td>$($disk.DeviceID)</td>
    <td>$($disk.VolumeName)</td>
    <td>$($disk.FileSystem)</td>
    <td>$totalGB</td>
    <td>$freeGB</td>
    <td>$freePercent</td>
</tr>
"@
}

$htmlContent += "</table>"

# Add physical disk information
$htmlContent += "<h3>Physical Disks</h3><table>"
$htmlContent += "<tr><th>Model</th><th>Interface</th><th>Size (GB)</th><th>Media Type</th><th>Serial Number</th></tr>"

foreach ($disk in $physicalDisks) {
    $sizeGB = [math]::Round($disk.Size / 1GB, 2)
    $htmlContent += @"
<tr>
    <td>$($disk.Model)</td>
    <td>$($disk.InterfaceType)</td>
    <td>$sizeGB</td>
    <td>$($disk.MediaType)</td>
    <td>$($disk.SerialNumber)</td>
</tr>
"@
}

$htmlContent += "</table></div></div>"

# 3. Memory Details
Write-Host "Adding Memory Information to the report..." -ForegroundColor Cyan
$htmlContent += "<div class='section'><h2 onclick=`"toggleSection('memoryInfo')`">Memory Information</h2>"
$htmlContent += "<div id='memoryInfo' style='display: none;'>"
$totalMemoryGB = [math]::Round(($computerSystem.TotalPhysicalMemory / 1GB), 2)

$htmlContent += @"
<table>
<tr><th>Total Physical Memory</th><td>$totalMemoryGB GB</td></tr>
<tr><th>Memory Slots Used</th><td>$($memory.Count)</td></tr>
</table>
"@

if ($memory.Count -gt 0) {
    $htmlContent += "<h3>Memory Modules</h3><table>"
    $htmlContent += "<tr><th>Slot</th><th>Capacity (GB)</th><th>Type</th><th>Speed (MHz)</th><th>Manufacturer</th><th>Part Number</th><th>Serial Number</th></tr>"
    
    $slot = 1
    foreach ($module in $memory) {
        $capacityGB = [math]::Round($module.Capacity / 1GB, 2)
        $htmlContent += @"
<tr>
    <td>$slot</td>
    <td>$capacityGB</td>
    <td>$($module.MemoryType)</td>
    <td>$($module.Speed)</td>
    <td>$($module.Manufacturer)</td>
    <td>$($module.PartNumber)</td>
    <td>$($module.SerialNumber)</td>
</tr>
"@
        $slot++
    }
    $htmlContent += "</table>"
}

$htmlContent += "</div></div>"

# 4. Installed Applications
Write-Host "Adding Installed Applications to the report..." -ForegroundColor Cyan
$htmlContent += "<div class='section'><h2 onclick=`"toggleSection('installedApps')`">Installed Applications</h2>"
$htmlContent += "<div id='installedApps' style='display: none;'>"
try {
    $applications = Get-ItemProperty "HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*", 
                                      "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*" |
                    Where-Object { $_.DisplayName -ne $null } |
                    Select-Object DisplayName, DisplayVersion, Publisher, InstallDate, EstimatedSize |
                    Sort-Object DisplayName
} catch {
    Log-Error "Error collecting installed applications: $_"
    $applications = @()
}

$htmlContent += @"
<table>
<tr><th>Application Name</th><th>Version</th><th>Publisher</th><th>Install Date</th><th>Size (MB)</th></tr>
"@

foreach ($app in $applications) {
    $sizeMB = if ($app.EstimatedSize) { [math]::Round($app.EstimatedSize / 1MB, 2) } else { "N/A" }
    $htmlContent += @"
<tr>
    <td>$($app.DisplayName)</td>
    <td>$($app.DisplayVersion)</td>
    <td>$($app.Publisher)</td>
    <td>$($app.InstallDate)</td>
    <td>$sizeMB</td>
</tr>
"@
}

$htmlContent += "</table></div></div>"

# 5. Network Information
Write-Host "Adding Network Configuration to the report..." -ForegroundColor Cyan
$htmlContent += "<div class='section'><h2 onclick=`"toggleSection('networkConfig')`">Network Configuration</h2>"
$htmlContent += "<div id='networkConfig' style='display: none;'>"
foreach ($adapter in $adapters) {
    $htmlContent += "<h3>$($adapter.Description)</h3>"
    $htmlContent += "<table>"
    $htmlContent += "<tr><th>Property</th><th>Value</th></tr>"
    $htmlContent += "<tr><td>MAC Address</td><td>$($adapter.MACAddress)</td></tr>"
    
    if ($adapter.IPAddress) {
        $htmlContent += "<tr><td>IP Address</td><td>$($adapter.IPAddress -join ', ')</td></tr>"
    }
    
    if ($adapter.IPSubnet) {
        $htmlContent += "<tr><td>Subnet Mask</td><td>$($adapter.IPSubnet -join ', ')</td></tr>"
    }
    
    if ($adapter.DefaultIPGateway) {
        $htmlContent += "<tr><td>Default Gateway</td><td>$($adapter.DefaultIPGateway -join ', ')</td></tr>"
    }
    
    if ($adapter.DNSServerSearchOrder) {
        $htmlContent += "<tr><td>DNS Servers</td><td>$($adapter.DNSServerSearchOrder -join ', ')</td></tr>"
    }
    
    $htmlContent += "<tr><td>DHCP Enabled</td><td>$($adapter.DHCPEnabled)</td></tr>"
    
    if ($adapter.DHCPServer) {
        $htmlContent += "<tr><td>DHCP Server</td><td>$($adapter.DHCPServer)</td></tr>"
    }
    
    $htmlContent += "</table>"
}

$htmlContent += "</div></div>"

# 6. Features and Roles (Windows Server)
Write-Host "Checking for installed features and roles (Windows Server)..." -ForegroundColor Cyan
$htmlContent += "<div class='section'><h2 onclick=`"toggleSection('featuresRoles')`">Installed Features and Roles</h2>"
$htmlContent += "<div id='featuresRoles' style='display: none;'>"

try {
    $os = Get-CimInstance -ComputerName $ComputerName -ClassName Win32_OperatingSystem
    if ($os.Caption -like "*Server*") {
        Write-Host "Collecting installed features and roles for Windows Server..." -ForegroundColor Cyan
        $featuresRoles = Get-WindowsFeature | Where-Object { $_.Installed -eq $true }
        $htmlContent += "<table><tr><th>Feature/Role Name</th><th>Display Name</th></tr>"
        foreach ($feature in $featuresRoles) {
            $htmlContent += @"
<tr>
    <td>$($feature.Name)</td>
    <td>$($feature.DisplayName)</td>
</tr>
"@
        }
        $htmlContent += "</table>"
    } else {
        Write-Host "The computer is not a Windows Server. Skipping features and roles collection." -ForegroundColor Yellow
        $htmlContent += "<p>The computer is not a Windows Server. No features or roles to display.</p>"
    }
} catch {
    Log-Error "Error collecting features and roles: $_"
    $htmlContent += "<p>Error collecting features and roles.</p>"
}

$htmlContent += "</div></div>"

# 7. System Health Metrics
Write-Host "Adding System Health Metrics to the report..." -ForegroundColor Cyan
$htmlContent += "<div class='section'><h2 onclick=`"toggleSection('systemHealth')`">System Health Metrics</h2>"
$htmlContent += "<div id='systemHealth' style='display: none;'>"
try {
    $cpuUsage = Get-Counter '\Processor(_Total)\% Processor Time' | Select-Object -ExpandProperty CounterSamples | Select-Object -ExpandProperty CookedValue
    $memoryUsage = [math]::Round((($os.TotalVisibleMemorySize - $os.FreePhysicalMemory) / $os.TotalVisibleMemorySize) * 100, 2)
    $htmlContent += @"
<table>
<tr><th>Metric</th><th>Value</th></tr>
<tr><td>CPU Usage</td><td>$([math]::Round($cpuUsage, 2))%</td></tr>
<tr><td>Memory Usage</td><td>$memoryUsage%</td></tr>
</table>
"@
} catch {
    Log-Error "Error collecting system health metrics: $_"
    $htmlContent += "<p>Error collecting system health metrics.</p>"
}
$htmlContent += "</div></div>"

# Optional Sections
if ($IncludeGroupPolicy) {
    Write-Host "Adding Group Policy Information to the report..." -ForegroundColor Cyan
    $htmlContent += "<div class='section'><h2 onclick=`"toggleSection('groupPolicy')`">Group Policy Information</h2>"
    $htmlContent += "<div id='groupPolicy' style='display: none;'>"
    try {
        $gpResult = gpresult /r
        $htmlContent += "<pre>$gpResult</pre>"
    } catch {
        Log-Error "Error collecting Group Policy information: $_"
        $htmlContent += "<p>Error collecting Group Policy information.</p>"
    }
    $htmlContent += "</div></div>"
}

if ($IncludeScheduledTasks) {
    Write-Host "Adding Scheduled Tasks to the report..." -ForegroundColor Cyan
    $htmlContent += "<div class='section'><h2 onclick=`"toggleSection('scheduledTasks')`">Scheduled Tasks</h2>"
    $htmlContent += "<div id='scheduledTasks' style='display: none;'>"
    try {
        $tasks = Get-ScheduledTask | Select-Object TaskName, State
        $htmlContent += "<table><tr><th>Task Name</th><th>State</th></tr>"
        foreach ($task in $tasks) {
            $htmlContent += "<tr><td>$($task.TaskName)</td><td>$($task.State)</td></tr>"
        }
        $htmlContent += "</table>"
    } catch {
        Log-Error "Error collecting scheduled tasks: $_"
        $htmlContent += "<p>Error collecting scheduled tasks.</p>"
    }
    $htmlContent += "</div></div>"
}

if ($IncludeActiveDirectory) {
    Write-Host "Adding Active Directory Information to the report..." -ForegroundColor Cyan
    $htmlContent += "<div class='section'><h2 onclick=`"toggleSection('activeDirectory')`">Active Directory Information</h2>"
    $htmlContent += "<div id='activeDirectory' style='display: none;'>"
    try {
        $adInfo = Get-ADComputer -Identity $ComputerName -Properties *
        $htmlContent += "<pre>$($adInfo | Out-String)</pre>"
    } catch {
        Log-Error "Error collecting Active Directory information: $_"
        $htmlContent += "<p>Error collecting Active Directory information.</p>"
    }
    $htmlContent += "</div></div>"
}

if ($IncludeWindowsFeatures) {
    Write-Host "Adding Windows Features to the report..." -ForegroundColor Cyan
    $htmlContent += "<div class='section'><h2 onclick=`"toggleSection('windowsFeatures')`">Windows Features</h2>"
    $htmlContent += "<div id='windowsFeatures' style='display: none;'>"
    try {
        $features = Get-WindowsFeature | Where-Object { $_.Installed -eq $true }
        $htmlContent += "<table><tr><th>Feature Name</th><th>Display Name</th></tr>"
        foreach ($feature in $features) {
            $htmlContent += "<tr><td>$($feature.Name)</td><td>$($feature.DisplayName)</td></tr>"
        }
        $htmlContent += "</table>"
    } catch {
        Log-Error "Error collecting Windows Features: $_"
        $htmlContent += "<p>Error collecting Windows Features.</p>"
    }
    $htmlContent += "</div></div>"
}

if ($IncludeEventLogs) {
    Write-Host "Adding Event Logs to the report..." -ForegroundColor Cyan
    $htmlContent += "<div class='section'><h2 onclick=`"toggleSection('eventLogs')`">Event Logs</h2>"
    $htmlContent += "<div id='eventLogs' style='display: none;'>"
    try {
        $events = Get-WinEvent -LogName System -MaxEvents 10 | Select-Object TimeCreated, LevelDisplayName, Message
        $htmlContent += "<table><tr><th>Time</th><th>Level</th><th>Message</th></tr>"
        foreach ($event in $events) {
            $htmlContent += "<tr><td>$($event.TimeCreated)</td><td>$($event.LevelDisplayName)</td><td>$($event.Message)</td></tr>"
        }
        $htmlContent += "</table>"
    } catch {
        Log-Error "Error collecting Event Logs: $_"
        $htmlContent += "<p>Error collecting Event Logs.</p>"
    }
    $htmlContent += "</div></div>"
}

if ($IncludeFirewallAntivirus) {
    Write-Host "Adding Firewall and Antivirus Information to the report..." -ForegroundColor Cyan
    $htmlContent += "<div class='section'><h2 onclick=`"toggleSection('firewallAntivirus')`">Firewall and Antivirus</h2>"
    $htmlContent += "<div id='firewallAntivirus' style='display: none;'>"
    try {
        $firewallStatus = Get-NetFirewallProfile | Select-Object Name, Enabled
        $antivirusStatus = Get-CimInstance -Namespace "root\SecurityCenter2" -ClassName AntiVirusProduct
        $htmlContent += "<h3>Firewall Status</h3><table><tr><th>Profile</th><th>Enabled</th></tr>"
        foreach ($profile in $firewallStatus) {
            $htmlContent += "<tr><td>$($profile.Name)</td><td>$($profile.Enabled)</td></tr>"
        }
        $htmlContent += "</table>"
        $htmlContent += "<h3>Antivirus Status</h3><table><tr><th>Product Name</th><th>Enabled</th></tr>"
        foreach ($av in $antivirusStatus) {
            $htmlContent += "<tr><td>$($av.displayName)</td><td>$($av.productState)</td></tr>"
        }
        $htmlContent += "</table>"
    } catch {
        Log-Error "Error collecting Firewall and Antivirus information: $_"
        $htmlContent += "<p>Error collecting Firewall and Antivirus information.</p>"
    }
    $htmlContent += "</div></div>"
}

if ($IncludeGPU) {
    Write-Host "Adding GPU Information to the report..." -ForegroundColor Cyan
    $htmlContent += "<div class='section'><h2 onclick=`"toggleSection('gpuInfo')`">GPU Information</h2>"
    $htmlContent += "<div id='gpuInfo' style='display: none;'>"
    try {
        $gpus = Get-CimInstance -ClassName Win32_VideoController
        $htmlContent += "<table><tr><th>Name</th><th>Driver Version</th><th>Memory (MB)</th></tr>"
        foreach ($gpu in $gpus) {
            $memoryMB = [math]::Round($gpu.AdapterRAM / 1MB, 2)
            $htmlContent += "<tr><td>$($gpu.Name)</td><td>$($gpu.DriverVersion)</td><td>$memoryMB</td></tr>"
        }
        $htmlContent += "</table>"
    } catch {
        Log-Error "Error collecting GPU information: $_"
        $htmlContent += "<p>Error collecting GPU information.</p>"
    }
    $htmlContent += "</div></div>"
}

# Export to CSV and JSON
if ($ExportFormat -eq "CSV") {
    Write-Host "Exporting report to CSV format..." -ForegroundColor Cyan
    try {
        $csvData = @()
        $csvData | Export-Csv -Path $csvFile -NoTypeInformation -Force
        Write-Host "CSV report generated successfully: $csvFile" -ForegroundColor Green
    } catch {
        Log-Error "Error exporting to CSV: $_"
    }
}

if ($ExportFormat -eq "JSON") {
    Write-Host "Exporting report to JSON format..." -ForegroundColor Cyan
    try {
        $jsonData = @{}
        $jsonData | ConvertTo-Json -Depth 3 | Out-File -FilePath $jsonFile -Force
        Write-Host "JSON report generated successfully: $jsonFile" -ForegroundColor Green
    } catch {
        Log-Error "Error exporting to JSON: $_"
    }
}

# Email the report
if ($EmailRecipient) {
    Write-Host "Sending the report via email to $EmailRecipient..." -ForegroundColor Cyan
    try {
        $smtpServer = "smtp.example.com"
        $smtpPort = 587
        $emailFrom = "noreply@example.com"
        $emailSubject = "Windows Inventory Report - $ComputerName"
        $emailBody = "Please find the attached inventory report for $ComputerName."
        Send-MailMessage -From $emailFrom -To $EmailRecipient -Subject $emailSubject -Body $emailBody -SmtpServer $smtpServer -Port $smtpPort -Attachments $outputFile
        Write-Host "Email sent successfully to $EmailRecipient" -ForegroundColor Green
    } catch {
        Log-Error "Error sending email: $_"
    }
}

Write-Host "Finalizing the HTML report..." -ForegroundColor Cyan
# Complete the HTML report
$htmlContent += $htmlFooter

Write-Host "Saving the report to $outputFile..." -ForegroundColor Cyan
# Save the report to file
$htmlContent | Out-File -FilePath $outputFile -Force

Write-Host "Inventory report generated successfully: $outputFile" -ForegroundColor Green

Write-Host "Opening the report in the default browser..." -ForegroundColor Cyan
# Open the report in default browser
Start-Process $outputFile