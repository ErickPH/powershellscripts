<#
.SYNOPSIS
    Creates an HTML inventory of Hyper-V server with navigation and summary features.
.DESCRIPTION
    This script gathers detailed information about the Hyper-V host and all VMs,
    including a navigation section and summary table for quick reference.
.NOTES
    File Name      : HyperV-Inventory.ps1
    Author         : Your Name
    Prerequisite   : PowerShell 5.1 or later, Hyper-V module
    Version        : 1.6
#>

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
        margin-top: 20px;
        border: 1px solid #ddd;
    }
    .vm-link {
        color: #2a5885;
        text-decoration: none;
    }
    .vm-link:hover {
        text-decoration: underline;
    }
</style>
"@

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

# Get Hyper-V host information
$hostComputer = Get-CimInstance -ClassName Win32_ComputerSystem
$hostOS = Get-CimInstance -ClassName Win32_OperatingSystem
$hostProcessor = Get-CimInstance -ClassName Win32_Processor | Select-Object -First 1
$hostMemory = Get-CimInstance -ClassName Win32_PhysicalMemory | Measure-Object -Property Capacity -Sum | Select-Object Sum
$hostNetwork = Get-NetAdapter | Where-Object { $_.Status -eq 'Up' }
$hostVirtualSwitch = Get-VMSwitch
$hostStorage = Get-PhysicalDisk | Select-Object FriendlyName, MediaType, Size, HealthStatus, OperationalStatus

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
        # Cluster commands failed
    }
}

# Get all VMs and collect summary data
$virtualMachines = Get-VM | Sort-Object Name
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
    
    <h2>Virtual Machines ($($virtualMachines.Count))</h2>
"@

# Process each VM
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
            # Cluster commands failed for this VM
        }
    }
    
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

    $htmlContent += @"
    </div>
"@
}

# Add summary section at the end
$htmlContent += @"
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

foreach ($vm in $vmSummaryData) {
    $htmlContent += @"
        <tr>
            <td>$($vm.Name)</td>
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
</body>
</html>
"@

# Save the HTML report
$reportPath = "HyperV-Inventory-$($hostComputer.Name)-$(Get-Date -Format 'yyyyMMdd-HHmmss').html"
$htmlContent | Out-File -FilePath $reportPath -Force

Write-Host "Inventory report generated: $((Get-Item $reportPath).FullName)"