<#
.SYNOPSIS
    Executes a PowerShell script on multiple remote machines as local administrator.
.DESCRIPTION
    This script attempts to connect to each machine in a provided list using administrator credentials,
    copies a specified PowerShell script to the remote machine, executes it, and reports the results.
.PARAMETER ComputerList
    Path to a text file containing list of target computers (one per line)
.PARAMETER ScriptPath
    Path to the PowerShell script that will be executed remotely
.PARAMETER Credential
    Administrator credentials to use for remote connections
.PARAMETER LogPath
    Path where the operation log will be saved (default: RemoteScriptExecution.log)
.EXAMPLE
    .\Execute-RemoteScript.ps1 -ComputerList .\computers.txt -ScriptPath .\deploy.ps1 -Credential (Get-Credential)
#>

param (
    [Parameter(Mandatory=$true)]
    [string]$ComputerList,
    
    [Parameter(Mandatory=$true)]
    [string]$ScriptPath,
    
    [Parameter(Mandatory=$true)]
    [System.Management.Automation.PSCredential]$Credential,
    
    [string]$LogPath = "RemoteScriptExecution.log"
)

# Initialize log file
$timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
"=== Remote Script Execution Log - $timestamp ===" | Out-File -FilePath $LogPath -Append
"Script to execute: $ScriptPath" | Out-File -FilePath $LogPath -Append
"Target computers: $ComputerList" | Out-File -FilePath $LogPath -Append
"" | Out-File -FilePath $LogPath -Append

# Verify the script file exists
if (-not (Test-Path -Path $ScriptPath -PathType Leaf)) {
    Write-Error "Script file not found at $ScriptPath"
    "ERROR: Script file not found at $ScriptPath" | Out-File -FilePath $LogPath -Append
    exit 1
}

# Read computer list
try {
    $computers = Get-Content -Path $ComputerList -ErrorAction Stop | Where-Object { $_ -match '\S' }
    if (-not $computers) {
        throw "Computer list is empty or contains only whitespace"
    }
}
catch {
    Write-Error "Error reading computer list: $_"
    "ERROR: Failed to read computer list - $_" | Out-File -FilePath $LogPath -Append
    exit 1
}

# Script block to execute remotely
$remoteScriptBlock = {
    param($scriptPath)
    
    try {
        # Execute the script
        $output = & $scriptPath 2>&1
        $result = @{
            Status = "Success"
            Output = $output
        }
    }
    catch {
        $result = @{
            Status = "Error"
            Output = $_.Exception.Message
        }
    }
    
    return $result
}

# Process each computer
foreach ($computer in $computers) {
    $computer = $computer.Trim()
    $logEntry = "Processing $computer : "
    
    try {
        # Test connection first
        if (-not (Test-Connection -ComputerName $computer -Count 1 -Quiet)) {
            throw "$computer is not reachable via ping"
        }
        
        # Create remote session
        $sessionParams = @{
            ComputerName = $computer
            Credential = $Credential
            ErrorAction = 'Stop'
        }
        
        $session = New-PSSession @sessionParams
        
        # Copy script to remote temp location
        $remoteScriptPath = "C:\Windows\Temp\$([System.IO.Path]::GetFileName($ScriptPath))"
        Copy-Item -Path $ScriptPath -Destination $remoteScriptPath -ToSession $session -Force
        
        # Execute script remotely
        $result = Invoke-Command -Session $session -ScriptBlock $remoteScriptBlock -ArgumentList $remoteScriptPath
        
        # Process results
        if ($result.Status -eq "Success") {
            $logEntry += "Script executed successfully"
            Write-Host "$computer : Success" -ForegroundColor Green
        }
        else {
            $logEntry += "Script execution failed: $($result.Output)"
            Write-Host "$computer : Error - $($result.Output)" -ForegroundColor Red
        }
        
        # Clean up
        Invoke-Command -Session $session -ScriptBlock {
            param($path)
            Remove-Item -Path $path -Force -ErrorAction SilentlyContinue
        } -ArgumentList $remoteScriptPath
        
        Remove-PSSession -Session $session
    }
    catch {
        $logEntry += "Connection/execution failed: $_"
        Write-Host "$computer : Error - $_" -ForegroundColor Red
    }
    
    # Log the result
    $logEntry | Out-File -FilePath $LogPath -Append
}

"`n=== Execution completed ===" | Out-File -FilePath $LogPath -Append
Write-Host "`nExecution complete. Results logged to $LogPath" -ForegroundColor Cyan