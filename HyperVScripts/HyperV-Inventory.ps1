<#
.SYNOPSIS
    Creates an HTML inventory of Hyper-V server with improved navigation features.
.DESCRIPTION
    This script gathers detailed information about the Hyper-V host and all VMs,
    including a navigation section and summary table at the top, plus "Go to Top" links.
.NOTES
    File Name      : HyperV-Inventory.ps1
    Author         : Erick Perez
    Date           : 2025-05-01
    GitHub         : https://github.com/erickph
    Prerequisite   : PowerShell 5.1 or later, Hyper-V module
    Version        : 1.12
    Changelog      :
        - Fixed guest OS detection to work for all VMs, not just the first one
        - Added Windows Server 2016 compatibility for guest OS detection
        - Added guest OS information for each VM (name, version, architecture, state, uptime)
        - Made -NonInteractive the default behavior
        - Added -Interactive parameter to enable user interaction
        - Added error logging to a log file for troubleshooting
        - Added validation for prerequisites like Hyper-V module and administrative privileges
        - Improved error handling with try-catch blocks for critical sections
        - Added progress indicators to inform the user about script execution stages
        - Allowed user to specify a custom output path for the HTML report
        - Provided a summary of errors encountered during execution

    Examples:
        # Run in non-interactive mode (default)
        .\HyperV-Inventory.ps1
        
        # Run in interactive mode
        .\HyperV-Inventory.ps1 -Interactive
        
        # Specify custom output path
        .\HyperV-Inventory.ps1 -OutputPath "C:\Reports\HyperV-Inventory.html"
        
        # Run in interactive mode with custom output path
        .\HyperV-Inventory.ps1 -Interactive -OutputPath "C:\Reports\HyperV-Inventory.html"

    About the Execution Policy:
        The script may require the execution policy to be set to Bypass, RemoteSigned or Unrestricted.
        You can set the execution policy using the following command:
        Set-ExecutionPolicy Bypass -Scope Process -Force
        This sets the execution policy for the current PowerShell session only.
#>

# Default parameters
$OutputPath = "$env:TEMP\$($env:COMPUTERNAME)-HyperV-Inventory-$(Get-Date -Format 'yyyyMMdd-HHmmss').html"
$Interactive = $false

# Check for command line arguments
if ($args.Count -gt 0) {
    for ($i = 0; $i -lt $args.Count; $i++) {
        if ($args[$i] -eq "-Interactive") {
            $Interactive = $true
        }
        elseif ($args[$i] -eq "-OutputPath" -and $i -lt ($args.Count - 1)) {
            $OutputPath = $args[$i + 1]
            $i++ # Skip the next argument as it's the path value
        }
    }
}

# HTML styling for the report
$htmlStyle = @"
<style>
    body {
        font-family: Arial, sans-serif;
        margin: 20px;
        color: #333;
    }
    h1, h2, h3, h4, h5 {
        color: #2a5885;
    }
    table {
        border-collapse: collapse;
        width: 100%;
        margin-bottom: 20px;
    }
    th {
        background-color: #2a5885;
        color: white;
        text-align: left;
        padding: 8px;
    }
    td {
        border: 1px solid #ddd;
        padding: 8px;
    }
    tr:nth-child(even) {
        background-color: #f2f2f2;
    }
    .host-info {
        background-color: #e6f2ff;
        padding: 15px;
        border-radius: 5px;
        margin-bottom: 20px;
    }
    .vm-section {
        margin-bottom: 30px;
        border: 1px solid #ddd;
        border-radius: 5px;
        padding: 15px;
    }
    .timestamp {
        font-size: 0.9em;
        color: #666;
        text-align: right;
        margin-bottom: 20px;
    }
    .snapshot-info {
        background-color: #fff8e6;
        padding: 10px;
        border-radius: 5px;
        margin-top: 10px;
    }
    .warning {
        color: #d9534f;
        font-weight: bold;
    }
    .sub-section {
        background-color: #f9f9f9;
        padding: 10px;
        border-radius: 5px;
        margin-top: 15px;
    }
    .cluster-info {
        background-color: #e6ffe6;
        padding: 10px;
        border-radius: 5px;
        margin-top: 15px;
    }
    .nav-section {
        background-color: #f5f5f5;
        padding: 15px;
        border-radius: 5px;
        margin-bottom: 20px;
        border: 1px solid #ddd;
    }
    .summary-section {
        background-color: #f0f8ff;
        padding: 15px;
        border-radius: 5px;
        margin-bottom: 20px;
        border: 1px solid #ddd;
    }
    .vm-link {
        color: #2a5885;
        text-decoration: none;
    }
    .vm-link:hover {
        text-decoration: underline;
    }
    .top-link {
        display: block;
        text-align: right;
        margin-top: 15px;
        color: #2a5885;
        text-decoration: none;
    }
    .top-link:hover {
        text-decoration: underline;
    }
</style>
"@

# Define a log file for error logging
$logFilePath = "$env:TEMP\$($env:COMPUTERNAME)-HyperV-Inventory-$(Get-Date -Format 'yyyyMMdd-HHmmss')-ErrorLog.log"

# Function to log errors
function Log-Error {
    param (
        [string]$Message
    )
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logEntry = "[$timestamp] ERROR: $Message"
    Add-Content -Path $logFilePath -Value $logEntry
    Write-Host $logEntry -ForegroundColor Red
}

# Validate prerequisites
if (-not (Get-Command -Name Get-VM -ErrorAction SilentlyContinue)) {
    Log-Error "Hyper-V module is not available. Please ensure it is installed and loaded."
    exit
}

# Check if the user has administrative privileges
if (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Log-Error "This script must be run as an administrator."
    exit
}

# Check if FailoverClusters module is available
$clusterModuleAvailable = $false
try {
    if (Get-Module -ListAvailable -Name FailoverClusters) {
        $clusterModuleAvailable = $true
        Import-Module FailoverClusters -ErrorAction SilentlyContinue
    }
}
catch {
    $clusterModuleAvailable = $false
}

# Get current date and time for the report
$reportDate = Get-Date -Format "yyyy-MM-dd HH:mm:ss"

# Wrap critical sections in try-catch blocks
try {
    # Get Hyper-V host information
    $hostComputer = Get-CimInstance -ClassName Win32_ComputerSystem
    $hostOS = Get-CimInstance -ClassName Win32_OperatingSystem
    $hostProcessor = Get-CimInstance -ClassName Win32_Processor | Select-Object -First 1
    $hostMemory = Get-CimInstance -ClassName Win32_PhysicalMemory | Measure-Object -Property Capacity -Sum | Select-Object Sum
    $hostNetwork = Get-NetAdapter | Where-Object { $_.Status -eq 'Up' }
    $hostVirtualSwitch = Get-VMSwitch
    $hostStorage = Get-PhysicalDisk | Select-Object FriendlyName, MediaType, Size, HealthStatus, OperationalStatus
} catch {
    Log-Error "Failed to retrieve host information: $_"
}

# Get cluster information if available
$clusterInfo = $null
$clusterNodes = @()
$clusterName = "Not part of a cluster"
if ($clusterModuleAvailable) {
    try {
        $clusterInfo = Get-Cluster -ErrorAction SilentlyContinue
        if ($clusterInfo) {
            $clusterName = $clusterInfo.Name
            $clusterNodes = Get-ClusterNode -Cluster $clusterInfo.Name | Select-Object Name, State
        }
    }
    catch {
        Log-Error "Failed to retrieve cluster information: $_"
    }
}

# Wrap critical sections in try-catch blocks
try {
    # Get all VMs
    $virtualMachines = Get-VM | Sort-Object Name
} catch {
    Log-Error "Failed to retrieve virtual machines: $_"
}

# Add progress indicators
Write-Host "Generating Hyper-V inventory report..." -ForegroundColor Cyan

# Allow user to specify a custom output path if in interactive mode
if ($Interactive) {
    $userPath = Read-Host "Enter the desired output path for the report (press Enter to use default: $OutputPath)"
    if ($userPath) {
        $OutputPath = $userPath
    }
}

# Get all VMs and collect summary data
$vmSummaryData = @()

# Create HTML content
$htmlContent = @"
<!DOCTYPE html>
<html>
<head>
    <title>Hyper-V Inventory Report - $($hostComputer.Name)</title>
    $htmlStyle
</head>
<body>
    <h1>Hyper-V Inventory Report</h1>
    <div class="timestamp">Report generated on: $reportDate</div>
    
    <!-- Navigation Section -->
    <div class="nav-section">
        <h2>Virtual Machine Quick Navigation</h2>
        <p>
"@

# Add navigation links for each VM
foreach ($vm in $virtualMachines) {
    $vmId = $vm.Name -replace '[^a-zA-Z0-9]',''
    $htmlContent += @"
            <a href="#$vmId" class="vm-link">$($vm.Name)</a> | 
"@
}

$htmlContent += @"
        </p>
    </div>

    <!-- Summary Section -->
    <div class="summary-section">
        <h2>Virtual Machine Summary</h2>
        <table>
            <tr>
                <th>VM Name</th>
                <th>IP Addresses</th>
                <th>Assigned Memory (GB)</th>
                <th>State</th>
                <th>ID</th>
            </tr>
"@

# Process each VM to collect summary data
foreach ($vm in $virtualMachines) {
    Write-Host "Generating Hyper-V inventory report for VM $($vm.Name)...." -ForegroundColor Cyan
    $vmMemory = $vm | Get-VMMemory
    $vmNetwork = $vm | Get-VMNetworkAdapter
    
    # Collect IP addresses for summary
    $ipAddresses = @()
    foreach ($adapter in $vmNetwork) {
        if ($adapter.IPAddresses) {
            $ipAddresses += $adapter.IPAddresses
        }
    }
    $ipList = if ($ipAddresses.Count -gt 0) { $ipAddresses -join ", " } else { "No IP assigned" }
    
    # Add VM to summary data
    $vmSummaryData += [PSCustomObject]@{
        Name = $vm.Name
        IPAddresses = $ipList
        MemoryGB = [math]::Round($vmMemory.Startup / 1GB, 2)
        State = $vm.State
        ID = $vm.Id
    }
}

# Add summary rows
foreach ($vm in $vmSummaryData) {
    $htmlContent += @"
            <tr>
                <td><a href="#$($vm.Name -replace '[^a-zA-Z0-9]','')" class="vm-link">$($vm.Name)</a></td>
                <td>$($vm.IPAddresses)</td>
                <td>$($vm.MemoryGB)</td>
                <td>$($vm.State)</td>
                <td>$($vm.ID)</td>
            </tr>
"@
}

$htmlContent += @"
        </table>
    </div>
    
    <div class="host-info">
        <h2>Host System: $($hostComputer.Name)</h2>
        
        <h3>Host Information</h3>
        <table>
            <tr><th>Property</th><th>Value</th></tr>
            <tr><td>Host Name</td><td>$($hostComputer.Name)</td></tr>
            <tr><td>Manufacturer</td><td>$($hostComputer.Manufacturer)</td></tr>
            <tr><td>Model</td><td>$($hostComputer.Model)</td></tr>
            <tr><td>Total Physical Memory</td><td>$([math]::Round($hostMemory.Sum / 1GB, 2)) GB</td></tr>
            <tr><td>OS</td><td>$($hostOS.Caption) ($($hostOS.OSArchitecture))</td></tr>
            <tr><td>OS Version</td><td>$($hostOS.Version)</td></tr>
            <tr><td>Processor</td><td>$($hostProcessor.Name)</td></tr>
            <tr><td>Logical Processors</td><td>$($hostProcessor.NumberOfLogicalCores)</td></tr>
            <tr><td>Hyper-V Version</td><td>$((Get-Command Get-VM).Module.Version)</td></tr>
            <tr><td>Cluster Name</td><td>$clusterName</td></tr>
        </table>
        
        <h3>Network Adapters</h3>
        <table>
            <tr><th>Name</th><th>InterfaceDescription</th><th>Speed</th><th>MACAddress</th><th>IPAddress</th></tr>
"@

foreach ($adapter in $hostNetwork) {
    try {
        $ipAddresses = (Get-NetIPAddress -InterfaceIndex $adapter.ifIndex -ErrorAction Stop).IPAddress -join ", "
    }
    catch {
        $ipAddresses = "No IP address assigned"
    }
    $htmlContent += @"
            <tr>
                <td>$($adapter.Name)</td>
                <td>$($adapter.InterfaceDescription)</td>
                <td>$($adapter.LinkSpeed)</td>
                <td>$($adapter.MacAddress)</td>
                <td>$ipAddresses</td>
            </tr>
"@
}

$htmlContent += @"
        </table>
        
        <h3>Virtual Switches</h3>
        <table>
            <tr><th>Name</th><th>Type</th><th>NetAdapterInterfaceDescription</th><th>Notes</th></tr>
"@

foreach ($vswitch in $hostVirtualSwitch) {
    $htmlContent += @"
            <tr>
                <td>$($vswitch.Name)</td>
                <td>$($vswitch.SwitchType)</td>
                <td>$($vswitch.NetAdapterInterfaceDescription)</td>
                <td>$($vswitch.Notes)</td>
            </tr>
"@
}

$htmlContent += @"
        </table>
        
        <h3>Physical Disks</h3>
        <table>
            <tr><th>Name</th><th>Type</th><th>Size</th><th>Health</th><th>Status</th></tr>
"@

foreach ($disk in $hostStorage) {
    $sizeGB = [math]::Round($disk.Size / 1GB, 2)
    $htmlContent += @"
            <tr>
                <td>$($disk.FriendlyName)</td>
                <td>$($disk.MediaType)</td>
                <td>$sizeGB GB</td>
                <td>$($disk.HealthStatus)</td>
                <td>$($disk.OperationalStatus)</td>
            </tr>
"@
}

# Show cluster nodes if available
if ($clusterModuleAvailable -and $clusterInfo) {
    $htmlContent += @"
        </table>
        
        <h3>Cluster Nodes</h3>
        <table>
            <tr><th>Node Name</th><th>State</th></tr>
"@
    foreach ($node in $clusterNodes) {
        $htmlContent += @"
            <tr>
                <td>$($node.Name)</td>
                <td>$($node.State)</td>
            </tr>
"@
    }
}

$htmlContent += @"
        </table>
    </div>
    
    <h2>Virtual Machine Details</h2>
"@

# Process each VM for detailed sections
foreach ($vm in $virtualMachines) {
    $vmProcessor = $vm | Get-VMProcessor
    $vmMemory = $vm | Get-VMMemory
    $vmNetwork = $vm | Get-VMNetworkAdapter
    $vmHardDrives = $vm | Get-VMHardDiskDrive
    $vmDVDDrives = $vm | Get-VMDvdDrive
    $vmIntegrationServices = $vm | Get-VMIntegrationService
    $vmSnapshots = $vm | Get-VMSnapshot
    $vmSecurity = $vm | Get-VMSecurity
    $vmFirmware = $vm | Get-VMFirmware
    
    # Get guest OS information if available
    $vmGuestOS = $null
    try {
        if (Get-Command -Name Get-VMGuest -ErrorAction SilentlyContinue) {
            # Modern method (Windows Server 2019+)
            $vmGuestOS = $vm | Get-VMGuest
            Write-Host "Retrieved guest OS info for $($vm.Name) using Get-VMGuest: $($vmGuestOS.OSName)" -ForegroundColor Cyan
        } else {
            # Legacy method (Windows Server 2016)
            $vmElementName = $vm.ElementName
            if (-not $vmElementName) {
                $vmElementName = $vm.Name
            }
            
            $kvp = Get-CimInstance -Namespace root\virtualization\v2 -ClassName Msvm_ComputerSystem -Filter "ElementName='$vmElementName'" | 
                   Get-CimAssociatedInstance -ResultClassName Msvm_KvpExchangeComponent
            
            if ($kvp -and $kvp.GuestIntrinsicExchangeItems) {
                $kvpData = @{}
                foreach ($item in $kvp.GuestIntrinsicExchangeItems) {
                    if ($item.Data -ne $null) {
                        $kvpData[$item.Name] = [System.Text.Encoding]::Unicode.GetString($item.Data)
                    }
                }
                
                if ($kvpData.Count -gt 0) {
                    $vmGuestOS = [PSCustomObject]@{
                        OSName = $kvpData['OSName']
                        OSVersion = $kvpData['OSVersion']
                        OSArchitecture = $kvpData['OSArchitecture']
                        State = $vm.State
                        Uptime = $vm.Uptime
                    }
                    Write-Host "Retrieved guest OS info for $($vm.Name) using KVP: $($vmGuestOS.OSName)" -ForegroundColor Cyan
                } else {
                    Write-Host "No valid guest OS data found for $($vm.Name)" -ForegroundColor Yellow
                }
            } else {
                Write-Host "No KVP data available for $($vm.Name)" -ForegroundColor Yellow
            }
        }
    } catch {
        Write-Host "Error retrieving guest OS info for $($vm.Name): $_" -ForegroundColor Red
        $vmGuestOS = $null
    }
    
    # Get NUMA information if available
    $numaInfo = $null
    try {
        if (Get-Command -Name Get-VMNumaNode -ErrorAction SilentlyContinue) {
            $numaInfo = $vm | Get-VMNumaNode
        }
    } catch {
        $numaInfo = $null
    }
    
    # Get cluster information for VM if available
    $vmClusterInfo = $null
    $vmClusterOwner = "Not clustered"
    $vmClusterGroup = "N/A"
    $vmClusterState = "N/A"
    $vmClusterPreferredOwners = "N/A"
    
    if ($clusterModuleAvailable -and $clusterInfo) {
        try {
            $vmClusterGroup = Get-ClusterGroup -Cluster $clusterInfo.Name | Where-Object { $_.Name -eq $vm.Name }
            if ($vmClusterGroup) {
                $vmClusterOwner = $vmClusterGroup.OwnerNode
                $vmClusterState = $vmClusterGroup.State
                $vmClusterResource = Get-ClusterResource -Cluster $clusterInfo.Name | Where-Object { $_.OwnerGroup -eq $vm.Name }
                $vmClusterPreferredOwners = ($vmClusterResource | Get-ClusterOwnerNode).Ownernodes -join ", "
            }
        }
        catch {
            Log-Error "Failed to retrieve cluster information for VM $($vm.Name): $_"
        }
    }
    
    # Create anchor for navigation
    $vmId = $vm.Name -replace '[^a-zA-Z0-9]',''
    $htmlContent += @"
    <div class="vm-section" id="$vmId">
        <h3>$($vm.Name)</h3>
        
        <table>
            <tr><th>Property</th><th>Value</th></tr>
            <tr><td>Name</td><td>$($vm.Name)</td></tr>
            <tr><td>ID</td><td>$($vm.Id)</td></tr>
            <tr><td>State</td><td>$($vm.State)</td></tr>
            <tr><td>Status</td><td>$($vm.Status)</td></tr>
            <tr><td>Generation</td><td>$($vm.Generation)</td></tr>
            <tr><td>Version</td><td>$($vm.Version)</td></tr>
            <tr><td>Uptime</td><td>$($vm.Uptime)</td></tr>
            <tr><td>CPU Usage</td><td>$($vm.CPUUsage)%</td></tr>
            <tr><td>Assigned Memory</td><td>$([math]::Round($vmMemory.Startup / 1GB, 2)) GB</td></tr>
            <tr><td>Dynamic Memory</td><td>$($vmMemory.DynamicMemoryEnabled)</td></tr>
            <tr><td>Minimum Memory</td><td>$([math]::Round($vmMemory.Minimum / 1GB, 2)) GB</td></tr>
            <tr><td>Maximum Memory</td><td>$([math]::Round($vmMemory.Maximum / 1GB, 2)) GB</td></tr>
            <tr><td>Snapshot Count</td><td>$($vmSnapshots.Count)</td></tr>
"@

    # Add guest OS information if available
    if ($vmGuestOS) {
        $htmlContent += @"
            <tr><td>Guest OS</td><td>$($vmGuestOS.OSName)</td></tr>
            <tr><td>Guest OS Version</td><td>$($vmGuestOS.OSVersion)</td></tr>
            <tr><td>Guest OS Architecture</td><td>$($vmGuestOS.OSArchitecture)</td></tr>
            <tr><td>Guest OS State</td><td>$($vmGuestOS.State)</td></tr>
            <tr><td>Guest OS Uptime</td><td>$($vmGuestOS.Uptime)</td></tr>
"@
    }

    $htmlContent += @"
        </table>

        <div class="cluster-info">
            <h4>Cluster Information</h4>
            <table>
                <tr><th>Property</th><th>Value</th></tr>
                <tr><td>Cluster Name</td><td>$clusterName</td></tr>
                <tr><td>Current Owner Node</td><td>$vmClusterOwner</td></tr>
                <tr><td>Cluster Group State</td><td>$vmClusterState</td></tr>
                <tr><td>Preferred Owners</td><td>$vmClusterPreferredOwners</td></tr>
            </table>
        </div>

        <div class="sub-section">
            <h4>Firmware Settings</h4>
            <table>
                <tr><th>Property</th><th>Value</th></tr>
                <tr><td>Secure Boot</td><td>$($vmFirmware.SecureBoot)</td></tr>
                <tr><td>Secure Boot Template</td><td>$($vmFirmware.SecureBootTemplate)</td></tr>
                <tr><td>Boot Order</td><td>$($vmFirmware.BootOrder -join ", ")</td></tr>
                <tr><td>Preferred Network Boot Protocol</td><td>$($vmFirmware.PreferredNetworkBootProtocol)</td></tr>
                <tr><td>Console Mode</td><td>$($vmFirmware.ConsoleMode)</td></tr>
                <tr><td>Pause After Boot Failure</td><td>$($vmFirmware.PauseAfterBootFailure)</td></tr>
            </table>
        </div>

        <div class="sub-section">
            <h4>Security Settings</h4>
            <table>
                <tr><th>Property</th><th>Value</th></tr>
                <tr><td>Shielded</td><td>$($vmSecurity.Shielded)</td></tr>
                <tr><td>Encrypt State and VM Migration Traffic</td><td>$($vmSecurity.EncryptStateAndVMMigrationTraffic)</td></tr>
                <tr><td>Security PAUSE</td><td>$($vmSecurity.SecurityPAUSE)</td></tr>
                <tr><td>TPM Enabled</td><td>$($vmSecurity.TpmEnabled)</td></tr>
                <tr><td>Key Storage Drive</td><td>$($vmSecurity.KeyStorageDrive)</td></tr>
            </table>
        </div>

        <div class="sub-section">
            <h4>Processor Settings</h4>
            <table>
                <tr><th>Property</th><th>Value</th></tr>
                <tr><td>Count</td><td>$($vmProcessor.Count)</td></tr>
                <tr><td>Compatibility for Migration</td><td>$($vmProcessor.CompatibilityForMigrationEnabled)</td></tr>
                <tr><td>Compatibility for Older Operating Systems</td><td>$($vmProcessor.CompatibilityForOlderOperatingSystemsEnabled)</td></tr>
                <tr><td>HwThreadCountPerCore</td><td>$($vmProcessor.HwThreadCountPerCore)</td></tr>
                <tr><td>Maximum</td><td>$($vmProcessor.Maximum)</td></tr>
                <tr><td>MaximumCountPerNumaNode</td><td>$($vmProcessor.MaximumCountPerNumaNode)</td></tr>
                <tr><td>MaximumCountPerNumaSocket</td><td>$($vmProcessor.MaximumCountPerNumaSocket)</td></tr>
                <tr><td>Relative Weight</td><td>$($vmProcessor.RelativeWeight)</td></tr>
                <tr><td>Reserve</td><td>$($vmProcessor.Reserve)</td></tr>
            </table>
        </div>
"@

    # Add NUMA section if available
    if ($numaInfo) {
        $htmlContent += @"
        <div class="sub-section">
            <h4>NUMA Settings</h4>
            <table>
                <tr><th>Property</th><th>Value</th></tr>
                <tr><td>Maximum Processors</td><td>$($numaInfo.MaximumProcessors)</td></tr>
                <tr><td>Maximum Memory (MB)</td><td>$($numaInfo.MaximumMemory)</td></tr>
                <tr><td>Nodes Count</td><td>$($numaInfo.NodesCount)</td></tr>
                <tr><td>Processors Count</td><td>$($numaInfo.ProcessorsCount)</td></tr>
                <tr><td>Memory (MB)</td><td>$($numaInfo.Memory)</td></tr>
                <tr><td>Memory Interleave</td><td>$($numaInfo.MemoryInterleave)</td></tr>
            </table>
        </div>
"@
    }
    
    $htmlContent += @"
        <h4>Network Adapters ($($vmNetwork.Count))</h4>
"@

    foreach ($adapter in $vmNetwork) {
        $ipAddresses = if ($adapter.IPAddresses) { $adapter.IPAddresses -join ", " } else { "No IP address assigned" }
        $htmlContent += @"
        <div class="sub-section">
            <table>
                <tr><th>Property</th><th>Value</th></tr>
                <tr><td>Name</td><td>$($adapter.Name)</td></tr>
                <tr><td>SwitchName</td><td>$($adapter.SwitchName)</td></tr>
                <tr><td>MACAddress</td><td>$($adapter.MacAddress)</td></tr>
                <tr><td>IPAddresses</td><td>$ipAddresses</td></tr>
                <tr><td>Status</td><td>$($adapter.Status)</td></tr>
                <tr><td>IsLegacy</td><td>$($adapter.IsLegacy)</td></tr>
                <tr><td>DynamicMacAddress</td><td>$($adapter.DynamicMacAddress)</td></tr>
                <tr><td>MacAddressSpoofing</td><td>$($adapter.MacAddressSpoofing)</td></tr>
                <tr><td>DhcpGuard</td><td>$($adapter.DhcpGuard)</td></tr>
                <tr><td>RouterGuard</td><td>$($adapter.RouterGuard)</td></tr>
                <tr><td>PortMirroring</td><td>$($adapter.PortMirroring)</td></tr>
                <tr><td>IeeePriorityTag</td><td>$($adapter.IeeePriorityTag)</td></tr>
                <tr><td>VmqWeight</td><td>$($adapter.VmqWeight)</td></tr>
                <tr><td>IovQueuePairsRequested</td><td>$($adapter.IovQueuePairsRequested)</td></tr>
                <tr><td>IovInterruptModeration</td><td>$($adapter.IovInterruptModeration)</td></tr>
                <tr><td>IovWeight</td><td>$($adapter.IovWeight)</td></tr>
                <tr><td>IPSecOffloadMaximumSecurityAssociation</td><td>$($adapter.IPSecOffloadMaximumSecurityAssociation)</td></tr>
                <tr><td>MaximumBandwidth</td><td>$($adapter.MaximumBandwidth)</td></tr>
                <tr><td>MinimumBandwidthAbsolute</td><td>$($adapter.MinimumBandwidthAbsolute)</td></tr>
                <tr><td>MinimumBandwidthWeight</td><td>$($adapter.MinimumBandwidthWeight)</td></tr>
                <tr><td>VlanId</td><td>$($adapter.VlanId)</td></tr>
            </table>
        </div>
"@
    }

    $htmlContent += @"
        <h4>Storage Devices</h4>
        <div class="sub-section">
            <h5>Hard Drives ($($vmHardDrives.Count))</h5>
            <table>
                <tr><th>ControllerType</th><th>ControllerNumber</th><th>Path</th><th>Size</th></tr>
"@

    foreach ($drive in $vmHardDrives) {
        $driveFile = Get-Item -Path $drive.Path -ErrorAction SilentlyContinue
        $sizeGB = if ($driveFile) { [math]::Round($driveFile.Length / 1GB, 2) } else { "N/A" }
        $htmlContent += @"
                <tr>
                    <td>$($drive.ControllerType)</td>
                    <td>$($drive.ControllerNumber)</td>
                    <td>$($drive.Path)</td>
                    <td>$sizeGB GB</td>
                </tr>
"@
    }

    $htmlContent += @"
            </table>
            
            <h5>DVD Drives ($($vmDVDDrives.Count))</h5>
            <table>
                <tr><th>ControllerType</th><th>ControllerNumber</th><th>Path</th></tr>
"@

    foreach ($drive in $vmDVDDrives) {
        $htmlContent += @"
                <tr>
                    <td>$($drive.ControllerType)</td>
                    <td>$($drive.ControllerNumber)</td>
                    <td>$($drive.Path)</td>
                </tr>
"@
    }

    $htmlContent += @"
            </table>
        </div>
        
        <div class="sub-section">
            <h4>Integration Services</h4>
            <table>
                <tr><th>Name</th><th>Enabled</th><th>OperationalStatus</th></tr>
"@

    foreach ($service in $vmIntegrationServices) {
        $htmlContent += @"
                <tr>
                    <td>$($service.Name)</td>
                    <td>$($service.Enabled)</td>
                    <td>$($service.OperationalStatus)</td>
                </tr>
"@
    }

    $htmlContent += @"
            </table>
        </div>
        
        <h4>Snapshots ($($vmSnapshots.Count))</h4>
"@

    if ($vmSnapshots.Count -gt 0) {
        $htmlContent += @"
        <div class="snapshot-info">
            <table>
                <tr><th>Name</th><th>Creation Time</th><th>Parent Snapshot</th><th>Snapshot Type</th><th>Notes</th></tr>
"@

        foreach ($snapshot in $vmSnapshots) {
            $parentSnapshot = if ($snapshot.ParentSnapshotName) { $snapshot.ParentSnapshotName } else { "None" }
            $htmlContent += @"
                <tr>
                    <td>$($snapshot.Name)</td>
                    <td>$($snapshot.CreationTime)</td>
                    <td>$parentSnapshot</td>
                    <td>$($snapshot.SnapshotType)</td>
                    <td>$($snapshot.Notes)</td>
                </tr>
"@
        }

        $htmlContent += @"
            </table>
            
            <h5>Snapshot File Information</h5>
            <table>
                <tr><th>Configuration File</th><th>Size</th></tr>
"@

        $snapshotFiles = Get-ChildItem -Path (Split-Path $vm.ConfigurationLocation) -Filter "*.avhdx" -Recurse -ErrorAction SilentlyContinue
        $totalSnapshotSize = 0
        
        foreach ($file in $snapshotFiles) {
            $sizeGB = [math]::Round($file.Length / 1GB, 2)
            $totalSnapshotSize += $file.Length
            $htmlContent += @"
                <tr>
                    <td>$($file.FullName)</td>
                    <td>$sizeGB GB</td>
                </tr>
"@
        }

        $totalSizeGB = [math]::Round($totalSnapshotSize / 1GB, 2)
        $htmlContent += @"
                <tr class="warning">
                    <td><strong>Total Snapshot Storage Used</strong></td>
                    <td><strong>$totalSizeGB GB</strong></td>
                </tr>
            </table>
        </div>
"@
    } else {
        $htmlContent += @"
        <p>No snapshots exist for this virtual machine.</p>
"@
    }

    # Add "Go to Top" link
    $htmlContent += @"
        <a href="#" class="top-link">Go to Top</a>
    </div>
"@
}

$htmlContent += @"
</body>
</html>
"@

# Save the HTML report with error handling
try {
    $htmlContent | Out-File -FilePath $OutputPath -Force
    Write-Host "Inventory report generated: $((Get-Item $OutputPath).FullName)" -ForegroundColor Green
} catch {
    Log-Error "Failed to save the report: $_"
}

# Provide a summary of errors if any
if (Test-Path $logFilePath) {
    Write-Host "Errors were encountered during execution. See the log file for details: $logFilePath" -ForegroundColor Yellow
} else {
    Write-Host "Script completed successfully without errors." -ForegroundColor Green
}