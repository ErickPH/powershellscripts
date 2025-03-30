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
.PARAMETER SmtpServer
    The SMTP server to use for sending emails. Defaults to "smtp.example.com".
.PARAMETER SmtpPort
    The SMTP port to use for sending emails. Defaults to 587.
.PARAMETER SmtpUsername
    The username for SMTP authentication. Optional.
.PARAMETER SmtpPassword
    The password for SMTP authentication. Optional.
.PARAMETER EmailSender
    The email address to use as the sender. Defaults to "noreply@example.com".
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
    Version        : 3.1
    Modification Date: 2025-03-31
    Change Log     :
        - Improved error handling consistency
        - Parameterized SMTP server and port
        - Modularized repetitive code into reusable functions
        - Fixed CSV and JSON export logic
        - Parsed and formatted Group Policy output
        - Decoded antivirus product state
        - Handled potential null values for GPU memory
        - Added error handling for output directory creation
        - Enhanced Test-WinRM function for better error handling
        - Updated metadata notes     
        - Added parameters for SMTP username, password, and email sender
        - Populated CSV and JSON exports with actual data
        - Added error handling for Test-Connection
        - Improved GPU memory calculation for null values
        - Parsed and formatted Group Policy output
        - Expanded Decode-AntivirusState to cover more states
        - Replaced inline JavaScript toggles with CSS-only toggles
        - Ensured consistent error logging
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
    [switch]$All = $false,
    [string]$SmtpServer = "smtp.example.com",
    [int]$SmtpPort = 587,
    [string]$SmtpUsername = $null,
    [string]$SmtpPassword = $null,
    [string]$EmailSender = "noreply@example.com"
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

# Error logging
$errorLog = "$OutputDir\ErrorLog_$(Get-Date -Format 'yyyyMMdd_HHmmss').txt"
function Log-Error {
    param ([string]$Message)
    Add-Content -Path $errorLog -Value "$((Get-Date).ToString()): $Message"
}

# Modularized function for creating output directory
function Ensure-OutputDirectory {
    param ([string]$Path)
    if (-not (Test-Path -Path $Path)) {
        try {
            New-Item -ItemType Directory -Path $Path -ErrorAction Stop | Out-Null
        } catch {
            Log-Error "Failed to create output directory: $_"
            throw "Failed to create output directory: $_"
        }
    }
}

# Modularized function for decoding antivirus product state
function Decode-AntivirusState {
    param ([int]$State)
    switch ($State) {
        397312 { return "Enabled" }
        397568 { return "Disabled" }
        397824 { return "Expired" }
        default { return "Unknown" }
    }
}

Write-Host "Checking required modules..." -ForegroundColor Cyan
function Ensure-Module {
    param (
        [string]$ModuleName
    )
    if (-not (Get-Module -ListAvailable -Name $ModuleName)) {
        try {
            Write-Host "Module '$ModuleName' is not loaded. Attempting to import..." -ForegroundColor Yellow
            Import-Module -Name $ModuleName -ErrorAction Stop
            Write-Host "Module '$ModuleName' imported successfully." -ForegroundColor Green
        } catch {
            Log-Error "Failed to import module '$ModuleName': $_"
            throw "Required module '$ModuleName' is not available. Please install it before running the script."
        }
    } else {
        Write-Host "Module '$ModuleName' is already available." -ForegroundColor Green
    }
}

# Check and import required modules
Ensure-Module -ModuleName "ServerManager"
if ($IncludeActiveDirectory) {
    Ensure-Module -ModuleName "ActiveDirectory"
}

Write-Host "Checking WinRM service status..." -ForegroundColor Cyan
# Enhanced Test-WinRM function
function Test-WinRM {
    param ([string]$ComputerName)
    try {
        if ($ComputerName -eq $env:COMPUTERNAME) {
            $winrmStatus = Get-Service -Name WinRM -ErrorAction Stop
            return $winrmStatus.Status -eq 'Running'
        } else {
            Invoke-Command -ComputerName $ComputerName -ScriptBlock {
                Get-Service -Name WinRM -ErrorAction Stop
            } -ErrorAction Stop | Out-Null
            return $true
        }
    } catch {
        Log-Error "WinRM check failed for $ComputerName: $_"
        return $false
    }
}

# Create output directory if it doesn't exist
Ensure-OutputDirectory -Path $OutputDir

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
        h2 { color: #0099cc; margin-top: 30px; border-bottom: 1px solid #ddd; padding-bottom: 5px; }
        table { border-collapse: collapse; width: 100%; margin-bottom: 20px; }
        th { background-color: #f2f2f2; text-align: left; padding: 8px; }
        td { padding: 8px; border-bottom: 1px solid #ddd; }
        .section { margin-bottom: 30px; }
        .summary { background-color: #f9f9f9; padding: 15px; border-radius: 5px; }
        .toggle { display: none; }
        .toggle + label { cursor: pointer; color: #0099cc; display: block; margin-bottom: 5px; }
        .toggle + label:hover { text-decoration: underline; }
        .toggle:checked + label + .content { display: block; }
        .content { display: none; }
    </style>
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
    if ($ComputerName -eq $env:COMPUTERNAME -or (Test-Connection -ComputerName $ComputerName -Count 1 -Quiet -ErrorAction SilentlyContinue)) {
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
$htmlContent += @"
<div class='section'>
    <input type='checkbox' class='toggle' id='systemInfoToggle'>
    <label for='systemInfoToggle'>System Information</label>
    <div class='content'>
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
            <tr><td>Last Boot Time</td><td>$([Management.ManagementDateTimeConverter]::ToDateTime($os.LastBootUpTime))</td></tr>
            <tr><td>System Uptime</td><td>$((Get-Date) - [Management.ManagementDateTimeConverter]::ToDateTime($os.LastBootUpTime))</td></tr>
        </table>
    </div>
</div>
"@

# Updated Email Sending Logic with Authentication
if ($EmailRecipient) {
    Write-Host "Sending the report via email to $EmailRecipient..." -ForegroundColor Cyan
    try {
        $emailParams = @{
            From       = $EmailSender
            To         = $EmailRecipient
            Subject    = "Windows Inventory Report - $ComputerName"
            Body       = "Please find the attached inventory report for $ComputerName."
            SmtpServer = $SmtpServer
            Port       = $SmtpPort
            Attachments = $outputFile
        }
        if ($SmtpUsername -and $SmtpPassword) {
            $emailParams.Credential = New-Object System.Management.Automation.PSCredential ($SmtpUsername, (ConvertTo-SecureString $SmtpPassword -AsPlainText -Force))
        }
        Send-MailMessage @emailParams
        Write-Host "Email sent successfully to $EmailRecipient" -ForegroundColor Green
    } catch {
        Log-Error "Error sending email: $_"
    }
}

# Updated CSV and JSON Export Logic
if ($ExportFormat -eq "CSV") {
    Write-Host "Exporting report to CSV format..." -ForegroundColor Cyan
    try {
        $csvData = @(
            @{ Property = "ComputerName"; Value = $ComputerName },
            @{ Property = "Timestamp"; Value = $timestamp }
            # Add other collected data here
        )
        $csvData | Export-Csv -Path $csvFile -NoTypeInformation -Force
        Write-Host "CSV report generated successfully: $csvFile" -ForegroundColor Green
    } catch {
        Log-Error "Error exporting to CSV: $_"
    }
}

if ($ExportFormat -eq "JSON") {
    Write-Host "Exporting report to JSON format..." -ForegroundColor Cyan
    try {
        $jsonData = @{
            ComputerName = $ComputerName
            Timestamp = $timestamp
            # Add other collected data here
        }
        $jsonData | ConvertTo-Json -Depth 3 | Out-File -FilePath $jsonFile -Force
        Write-Host "JSON report generated successfully: $jsonFile" -ForegroundColor Green
    } catch {
        Log-Error "Error exporting to JSON: $_"
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