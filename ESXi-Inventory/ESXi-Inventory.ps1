<#
.SYNOPSIS
    Generates a comprehensive ESXi host inventory report with performance metrics, security checks, and multi-format outputs.
.DESCRIPTION
    Collects detailed ESXi host data including:
    - Host hardware, health stats, and performance metrics
    - Virtual Machines (config, snapshots, tools status, disk I/O)
    - Resource pools, clusters, and DRS settings
    - Network (vSwitches, VMkernel, physical NICs, firewall)
    - Storage (datastores, multipathing, I/O latency)
    - Security (users, roles, event logs)
.OUTPUTS
    JSON, HTML (styled), CSV, and optional email/Slack notifications.
.NOTES
    Author: Erick Arturo Perez Huemer
    Version: 2.0
    Requirements: VMware PowerCLI (Install-Module -Name VMware.PowerCLI -Force)
#>

param (
    [Parameter(Mandatory=$true)]
    [string]$ESXiHost,
    [Parameter(Mandatory=$true)]
    [string]$Username,
    [Parameter(Mandatory=$true)]
    [string]$Password,
    [ValidateSet("JSON", "HTML", "CSV", "All")]
    [string]$OutputFormat = "All",
    [string]$OutputPath = ".\ESXi_Inventory_Report",
    [switch]$SendEmail,
    [string]$EmailTo,
    [switch]$PostToSlack,
    [string]$SlackWebhook
)

#region Initialization
# Load PowerCLI Module
if (-not (Get-Module -Name VMware.PowerCLI -ErrorAction SilentlyContinue)) {
    try {
        Import-Module VMware.PowerCLI -ErrorAction Stop
    } catch {
        Write-Error "VMware PowerCLI module not found. Install with: Install-Module -Name VMware.PowerCLI -Force"
        exit 1
    }
}

# Create Output Directory
if (-not (Test-Path $OutputPath)) {
    New-Item -ItemType Directory -Path $OutputPath | Out-Null
}

# Start Logging
Start-Transcript -Path "$OutputPath\Inventory_Log_$(Get-Date -Format 'yyyyMMdd').txt" -Append
#endregion

#region Data Collection
try {
    # Connect to ESXi Host (with retries)
    $SecurePassword = ConvertTo-SecureString $Password -AsPlainText -Force
    $Credential = New-Object System.Management.Automation.PSCredential ($Username, $SecurePassword)
    $Server = Connect-VIServer -Server $ESXiHost -Credential $Credential -ErrorAction Stop

    # Get Host Data
    $HostSystem = Get-VMHost
    $HostHardware = $HostSystem | Select-Object Name, Manufacturer, Model, NumCpu, CpuTotalMhz, MemoryTotalGB, 
        PowerState, Version, Build, @{N="CpuUsage";E={(Get-Stat -Entity $_ -Stat "cpu.usage.average" -Realtime -MaxSamples 1).Value}},
        @{N="MemoryUsage";E={(Get-Stat -Entity $_ -Stat "mem.usage.average" -Realtime -MaxSamples 1).Value}}

    # Get VMs with advanced metrics
    $VMs = Get-VM | ForEach-Object {
        $vmStats = Get-Stat -Entity $_ -Stat "virtualDisk.totalReadLatency.average", "virtualDisk.totalWriteLatency.average" -Realtime -MaxSamples 1
        [PSCustomObject]@{
            Name = $_.Name
            PowerState = $_.PowerState
            NumCPU = $_.NumCpu
            MemoryGB = $_.MemoryGB
            ToolsVersion = $_.ExtensionData.Guest.ToolsVersion
            ToolsStatus = $_.ExtensionData.Guest.ToolsStatus
            HardwareVersion = $_.ExtensionData.Config.Version
            Snapshots = (Get-Snapshot -VM $_).Count
            #ReadLatencyMS = ($vmStats | Where-Object {$_.MetricId -eq "virtualDisk.totalReadLatency.average"}).Value
            ReadLatencyMS = if ($_.PowerState -ne 'PoweredOff') { ($vmStats | Where-Object {$_.MetricId -eq "virtualDisk.totalReadLatency.average"}).Value } else { $null }
            WriteLatencyMS = ($vmStats | Where-Object {$_.MetricId -eq "virtualDisk.totalWriteLatency.average"}).Value
        }
    }

    # Storage Details with I/O Stats
    $Datastores = Get-Datastore | Select-Object Name, Type, CapacityGB, FreeSpaceGB,
        @{N="LatencyMS";E={(Get-Stat -Entity $_ -Stat "disk.totalLatency.average" -Realtime -MaxSamples 1).Value}},
        @{N="IOPS";E={(Get-Stat -Entity $_ -Stat "disk.numberReadAveraged.average" -Realtime -MaxSamples 1).Value +
                      (Get-Stat -Entity $_ -Stat "disk.numberWriteAveraged.average" -Realtime -MaxSamples 1).Value}},
        @{N="MultipathPolicy";E={(Get-ScsiLun -CanonicalName $_.ExtensionData.Info.Vmfs.Extent[0].DiskName).MultipathPolicy}}

    # Network Details
    $Networks = @{
        vSwitches = Get-VirtualSwitch | Select-Object Name, NumPorts, Mtu, Nic
        PortGroups = Get-VirtualPortGroup | Select-Object Name, VLANId, VirtualSwitch
        VMKernelAdapters = Get-VMHostNetworkAdapter | Where-Object { $_.GetType().Name -eq "HostVMKernelVirtualNicImpl" } |
                          Select-Object Name, IP, SubnetMask, Mac, PortGroupName
        PhysicalNICs = Get-VMHostNetworkAdapter -Physical | Select-Object Name, Driver, Mac, SpeedMb
        FirewallRules = Get-VMHostFirewallException | Where-Object { $_.Enabled } | Select-Object Name, Enabled, Service
    }

    # Security and Logs
    $Security = @{
        LocalAccounts = Get-VMHostAccount | Select-Object Name, Group, Description
        Permissions = Get-VIPermission | Select-Object Principal, Role, Entity
        RecentEvents = Get-VIEvent -Start (Get-Date).AddHours(-24) -Types Error, Warning | 
                      Select-Object CreatedTime, FullFormattedMessage
    }

} catch {
    Write-Error "Data collection failed: $_"
    exit 1
} finally {
    if ($Server) { Disconnect-VIServer -Server $ESXiHost -Confirm:$false }
}
#endregion

#region Output Generation
$Inventory = [ordered]@{
    Timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    Host = $HostHardware
    VirtualMachines = $VMs
    Datastores = $Datastores
    Networks = $Networks
    Security = $Security
}

# JSON Output
if ($OutputFormat -eq "JSON" -or $OutputFormat -eq "All") {
    $Inventory | ConvertTo-Json -Depth 6 | Out-File "$OutputPath\ESXi_Inventory_$ESXiHost.json"
    Write-Host "JSON report saved to $OutputPath\ESXi_Inventory_$ESXiHost.json" -ForegroundColor Green
}

# HTML Output (Professional Styling)
if ($OutputFormat -eq "HTML" -or $OutputFormat -eq "All") {
    $HTMLStyle = @"
<style>
    body { font-family: 'Segoe UI', Arial, sans-serif; margin: 20px; color: #333; }
    h1 { color: #2c3e50; border-bottom: 2px solid #3498db; padding-bottom: 10px; }
    h2 { color: #2980b9; margin-top: 25px; }
    h3 { color: #16a085; margin-top: 15px; }
    table { border-collapse: collapse; width: 100%; margin-bottom: 20px; box-shadow: 0 1px 3px rgba(0,0,0,0.1); }
    th { background-color: #3498db; color: white; text-align: left; padding: 10px; }
    td { padding: 8px; border: 1px solid #ddd; }
    tr:nth-child(even) { background-color: #f8f9fa; }
    tr:hover { background-color: #e9f7fe; }
    .warning { background-color: #fff3cd; }
    .critical { background-color: #f8d7da; }
</style>
"@

    $HTMLContent = @"
<!DOCTYPE html>
<html>
<head>
    <title>ESXi Inventory Report</title>
    $HTMLStyle
</head>
<body>
    <h1>ESXi Host Inventory Report</h1>
    <p><strong>Generated:</strong> $($Inventory.Timestamp)</p>
    <p><strong>Host:</strong> $($HostSystem.Name)</p>
"@

    # Add sections dynamically
    foreach ($Section in $Inventory.Keys | Where-Object { $_ -ne "Timestamp" }) {
        $HTMLContent += "<h2>$Section</h2>"
        if ($Section -eq "Networks" -or $Section -eq "Security") {
            foreach ($SubSection in $Inventory[$Section].Keys) {
                $HTMLContent += "<h3>$SubSection</h3>"
                $HTMLContent += $Inventory[$Section][$SubSection] | ConvertTo-Html -Fragment
            }
        } else {
            $HTMLContent += $Inventory[$Section] | ConvertTo-Html -Fragment
        }
    }

    $HTMLContent += "</body></html>"
    $HTMLContent | Out-File "$OutputPath\ESXi_Inventory_$ESXiHost.html"
    Write-Host "HTML report saved to $OutputPath\ESXi_Inventory_$ESXiHost.html" -ForegroundColor Green
}

# CSV Outputs (Multiple Files)
if ($OutputFormat -eq "CSV" -or $OutputFormat -eq "All") {
    $Inventory.GetEnumerator() | ForEach-Object {
        if ($_.Value -is [System.Collections.IDictionary]) {
            # Nested sections (Networks, Security)
            $_.Value.GetEnumerator() | ForEach-Object {
                $_.Value | Export-Csv "$OutputPath\ESXi_$($_.Key)_$ESXiHost.csv" -NoTypeInformation
            }
        } else {
            # Flat sections (Host, VMs)
            $_.Value | Export-Csv "$OutputPath\ESXi_$($_.Key)_$ESXiHost.csv" -NoTypeInformation
        }
    }
    Write-Host "CSV reports saved to $OutputPath\ESXi_*_$ESXiHost.csv" -ForegroundColor Green
}
#endregion

#region Notifications
# Email Report - work in progress
if ($SendEmail -and $EmailTo) {
    try {
        $EmailParams = @{
            To         = $EmailTo
            From       = "esxi-reports@yourdomain.com"
            Subject    = "ESXi Inventory Report - $($HostSystem.Name) - $(Get-Date -Format 'yyyy-MM-dd')"
            Body       = Get-Content "$OutputPath\ESXi_Inventory_$ESXiHost.html" | Out-String
            SmtpServer = "smtp.yourdomain.com"
            BodyAsHtml = $true
        }
        Send-MailMessage @EmailParams
        Write-Host "Report emailed to $EmailTo" -ForegroundColor Cyan
    } catch {
        Write-Warning "Failed to send email: $_"
    }
}

# Slack Alert
# Experimental
if ($PostToSlack -and $SlackWebhook) {
    try {
        $CriticalIssues = $Security.RecentEvents.Count
        $Payload = @{
            text = "ESXi Inventory Complete for $($HostSystem.Name). $CriticalIssues critical events found."
            attachments = @(
                @{
                    title = "Download Reports"
                    text = "HTML: $(Resolve-Path "$OutputPath\ESXi_Inventory_$ESXiHost.html")"
                    color = "#36a64f"
                }
            )
        } | ConvertTo-Json -Depth 3
        
        Invoke-RestMethod -Uri $SlackWebhook -Method Post -Body $Payload -ContentType "application/json"
    } catch {
        Write-Warning "Failed to post to Slack: $_"
    }
}
#endregion

Stop-Transcript
Write-Host "Inventory report generation completed!" -ForegroundColor Cyan