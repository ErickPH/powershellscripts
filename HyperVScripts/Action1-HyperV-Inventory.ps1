<#
.SYNOPSIS
    Creates an inventory of Hyper-V server and VMs in Action1-compatible format.
.DESCRIPTION
    This script gathers detailed information about the Hyper-V host and all VMs,
    organizing the data into an Action1-compatible format for data source integration.
.NOTES
    File Name      : HyperV-Inventory.ps1
    Author         : Erick Perez
    Date           : 2025-05-01
    GitHub         : https://github.com/erickph
    Prerequisite   : PowerShell 5.1 or later, Hyper-V module
    Version        : 2.0
    Changelog      :
        - Rewritten for Action1 data source compatibility
        - Outputs data as an array of objects with consistent properties
        - Added A1_Key property for Action1 integration
        - Removed HTML output in favor of structured data output
#>

# Initialize result array
$result = @()

# Validate prerequisites
if (-not (Get-Command -Name Get-VM -ErrorAction SilentlyContinue)) {
    Write-Error "Hyper-V module is not available. Please ensure it is installed and loaded."
    exit
}

# Check if the user has administrative privileges
if (-not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
    Write-Error "This script must be run as an administrator."
    exit
}

# Get Hyper-V host information
try {
    $hostComputer = Get-CimInstance -ClassName Win32_ComputerSystem
    $hostOS = Get-CimInstance -ClassName Win32_OperatingSystem
    $hostProcessor = Get-CimInstance -ClassName Win32_Processor | Select-Object -First 1
    $hostMemory = Get-CimInstance -ClassName Win32_PhysicalMemory | Measure-Object -Property Capacity -Sum | Select-Object Sum
    $hostNetwork = Get-NetAdapter | Where-Object { $_.Status -eq 'Up' }
    $hostVirtualSwitch = Get-VMSwitch
    $hostStorage = Get-PhysicalDisk | Select-Object FriendlyName, MediaType, Size, HealthStatus, OperationalStatus

    # Add host information to result
    $result += [PSCustomObject]@{
        Type = "Host"
        Name = $hostComputer.Name
        Manufacturer = $hostComputer.Manufacturer
        Model = $hostComputer.Model
        TotalMemoryGB = [math]::Round($hostMemory.Sum / 1GB, 2)
        OS = $hostOS.Caption
        OSVersion = $hostOS.Version
        OSArchitecture = $hostOS.OSArchitecture
        Processor = $hostProcessor.Name
        LogicalProcessors = $hostProcessor.NumberOfLogicalCores
        HyperVVersion = (Get-Command Get-VM).Module.Version
        NetworkAdapters = ($hostNetwork | ForEach-Object { "$($_.Name) ($($_.LinkSpeed))" }) -join "; "
        VirtualSwitches = ($hostVirtualSwitch | ForEach-Object { $_.Name }) -join "; "
        Storage = ($hostStorage | ForEach-Object { "$($_.FriendlyName) ($([math]::Round($_.Size / 1GB, 2))GB)" }) -join "; "
        A1_Key = "Host_$($hostComputer.Name)"
    }
} catch {
    Write-Error "Failed to retrieve host information: $_"
}

# Get all VMs
try {
    $virtualMachines = Get-VM | Sort-Object Name
} catch {
    Write-Error "Failed to retrieve virtual machines: $_"
    exit
}

# Process each VM
foreach ($vm in $virtualMachines) {
    try {
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
        $integrationServicesStatus = $null
        try {
            $integrationServices = $vm | Get-VMIntegrationService
            $integrationServicesStatus = $integrationServices | Where-Object { $_.Name -eq "Guest Service Interface" }
            
            if ($integrationServicesStatus -and $integrationServicesStatus.Enabled -and $integrationServicesStatus.OperationalStatus -eq "OK") {
                if (Get-Command -Name Get-VMGuest -ErrorAction SilentlyContinue) {
                    $vmGuestOS = $vm | Get-VMGuest
                } else {
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
                        }
                    }
                }
            }
        } catch {
            Write-Warning "Error retrieving guest OS info for $($vm.Name): $_"
        }

        # Add VM information to result
        $result += [PSCustomObject]@{
            Type = "VM"
            Name = $vm.Name
            ID = $vm.Id
            State = $vm.State
            Status = $vm.Status
            Generation = $vm.Generation
            Version = $vm.Version
            Uptime = $vm.Uptime
            CPUUsage = $vm.CPUUsage
            AssignedMemoryGB = [math]::Round($vmMemory.Startup / 1GB, 2)
            DynamicMemory = $vmMemory.DynamicMemoryEnabled
            MinimumMemoryGB = [math]::Round($vmMemory.Minimum / 1GB, 2)
            MaximumMemoryGB = [math]::Round($vmMemory.Maximum / 1GB, 2)
            ProcessorCount = $vmProcessor.Count
            NetworkAdapters = ($vmNetwork | ForEach-Object { "$($_.Name) ($($_.SwitchName))" }) -join "; "
            HardDrives = ($vmHardDrives | ForEach-Object { 
                $driveFile = Get-Item -Path $_.Path -ErrorAction SilentlyContinue
                $sizeGB = if ($driveFile) { [math]::Round($driveFile.Length / 1GB, 2) } else { "N/A" }
                "$($_.Path) ($sizeGB GB)"
            }) -join "; "
            DVDDrives = ($vmDVDDrives | ForEach-Object { $_.Path }) -join "; "
            IntegrationServices = ($vmIntegrationServices | Where-Object { $_.Enabled } | ForEach-Object { $_.Name }) -join "; "
            SnapshotCount = $vmSnapshots.Count
            GuestOS = if ($vmGuestOS) { $vmGuestOS.OSName } else { "Not available" }
            GuestOSVersion = if ($vmGuestOS) { $vmGuestOS.OSVersion } else { "Not available" }
            GuestOSArchitecture = if ($vmGuestOS) { $vmGuestOS.OSArchitecture } else { "Not available" }
            GuestOSState = if ($vmGuestOS) { $vmGuestOS.State } else { "Not available" }
            GuestOSUptime = if ($vmGuestOS) { $vmGuestOS.Uptime } else { "Not available" }
            SecureBoot = $vmFirmware.SecureBoot
            Shielded = $vmSecurity.Shielded
            A1_Key = "VM_$($vm.Name)_$($vm.Id)"
        }
    } catch {
        Write-Error "Failed to process VM $($vm.Name): $_"
    }
}

# Output the result
$result | ConvertTo-Json -Depth 10 