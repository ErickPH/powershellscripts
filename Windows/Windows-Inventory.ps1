<#
.SYNOPSIS
    Generates a comprehensive inventory report for a Windows machine.
.DESCRIPTION
    This script collects system information including machine details, disk information,
    memory details, installed applications, network configuration, and server roles (if applicable).
.NOTES
    File Name      : Windows-Inventory.ps1
    Author         : Erick Perez - quadrianweb.com
    Prerequisite   : PowerShell 5.1 or later, Run as Administrator for some details
    Version        : 1.1
    Change         : Added machine name to output filename
#>

# Output file configuration
$outputDir = "C:\InventoryReports"
$hostname = $env:COMPUTERNAME
$outputFile = "$outputDir\${hostname}_SystemInventory_$(Get-Date -Format 'yyyyMMdd_HHmmss').html"
$timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"

# Create output directory if it doesn't exist
if (-not (Test-Path -Path $outputDir)) {
    New-Item -ItemType Directory -Path $outputDir | Out-Null
}

# HTML Report Header
$htmlHeader = @"
<!DOCTYPE html>
<html>
<head>
    <title>Windows System Inventory - $hostname</title>
    <style>
        body { font-family: Arial, sans-serif; margin: 20px; }
        h1 { color: #0066cc; }
        h2 { color: #0099cc; margin-top: 30px; border-bottom: 1px solid #ddd; padding-bottom: 5px; }
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
        <p><strong>Computer Name:</strong> $hostname</p>
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

# 1. Machine Details
$htmlContent += "<div class='section'><h2>System Information</h2>"
$os = Get-CimInstance Win32_OperatingSystem
$computerSystem = Get-CimInstance Win32_ComputerSystem
$bios = Get-CimInstance Win32_BIOS
$processor = Get-CimInstance Win32_Processor | Select-Object -First 1

$htmlContent += @"
<table>
<tr><th>Property</th><th>Value</th></tr>
<tr><td>Hostname</td><td>$hostname</td></tr>
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
<tr><td>System Uptime</td><td>$((Get-Date) - $os.LastBootUpTime)</td></tr>
</table>
</div>
"@

# 2. Hard Disk Details
$htmlContent += "<div class='section'><h2>Disk Information</h2>"
$disks = Get-CimInstance Win32_LogicalDisk | Where-Object { $_.DriveType -eq 3 }

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
$physicalDisks = Get-CimInstance Win32_DiskDrive
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

$htmlContent += "</table></div>"

# 3. Memory Details
$htmlContent += "<div class='section'><h2>Memory Information</h2>"
$memory = Get-CimInstance Win32_PhysicalMemory
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

$htmlContent += "</div>"

# 4. Installed Applications
$htmlContent += "<div class='section'><h2>Installed Applications</h2>"
$applications = Get-ItemProperty "HKLM:\Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall\*", 
                                  "HKLM:\Software\Microsoft\Windows\CurrentVersion\Uninstall\*" |
                Where-Object { $_.DisplayName -ne $null } |
                Select-Object DisplayName, DisplayVersion, Publisher, InstallDate, EstimatedSize |
                Sort-Object DisplayName

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

$htmlContent += "</table></div>"

# 5. Network Information
$htmlContent += "<div class='section'><h2>Network Configuration</h2>"
$adapters = Get-CimInstance Win32_NetworkAdapterConfiguration | Where-Object { $_.IPEnabled -eq $true }

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

$htmlContent += "</div>"

# 6. Windows Server Roles (if applicable)
if ($os.ProductType -eq 2 -or $os.ProductType -eq 3) {
    $htmlContent += "<div class='section'><h2>Server Roles and Features</h2>"
    
    try {
        if (Get-Command -Name Get-WindowsFeature -ErrorAction SilentlyContinue) {
            $roles = Get-WindowsFeature | Where-Object { $_.Installed -eq $true }
            
            $htmlContent += "<table>"
            $htmlContent += "<tr><th>Role/Feature</th><th>Display Name</th><th>Install State</th></tr>"
            
            foreach ($role in $roles) {
                $htmlContent += @"
<tr>
    <td>$($role.Name)</td>
    <td>$($role.DisplayName)</td>
    <td>$($role.InstallState)</td>
</tr>
"@
            }
            
            $htmlContent += "</table>"
        } else {
            $htmlContent += "<p>Server Manager module not available to query roles and features.</p>"
        }
    } catch {
        $htmlContent += "<p>Error retrieving server roles: $($_.Exception.Message)</p>"
    }
    
    $htmlContent += "</div>"
}

# Complete the HTML report
$htmlContent += $htmlFooter

# Save the report to file
$htmlContent | Out-File -FilePath $outputFile -Force

Write-Host "Inventory report generated successfully: $outputFile" -ForegroundColor Green

# Open the report in default browser
Start-Process $outputFile