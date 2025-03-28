<#
.SYNOPSIS
    Scans the local network for computers listening on RDP port (3389).
.DESCRIPTION
    This script identifies all active computers in the local subnet that have RDP port open.
    Results are displayed as a comma-delimited table with IP, Hostname, and Status.
.NOTES
    Author         : Erick Perez - quadrianweb.com
    File Name      : Find-RDPComputers.ps1
    Prerequisite   : PowerShell 5.1 or later
    Run as Administrator for best results
#>

# Function to test if a port is open
function Test-Port {
    param (
        [string]$Computer,
        [int]$Port,
        [int]$Timeout = 100
    )
    
    try {
        $TCPClient = New-Object System.Net.Sockets.TcpClient
        $AsyncResult = $TCPClient.BeginConnect($Computer, $Port, $null, $null)
        $Wait = $AsyncResult.AsyncWaitHandle.WaitOne($Timeout, $false)
        if ($Wait) {
            $TCPClient.EndConnect($AsyncResult) | Out-Null
            $TCPClient.Close()
            return $true
        } else {
            $TCPClient.Close()
            return $false
        }
    } catch {
        return $false
    }
}

# Function to resolve hostname from IP
function Get-HostnameFromIP {
    param (
        [string]$IPAddress
    )
    
    try {
        $HostEntry = [System.Net.Dns]::GetHostEntry($IPAddress)
        return $HostEntry.HostName
    } catch {
        return "UNKNOWN"
    }
}

# Get local IP information to determine scan range
$LocalIP = (Get-NetIPAddress -AddressFamily IPv4 | Where-Object { $_.InterfaceAlias -notlike "*Loopback*" } | Select-Object -First 1).IPAddress
$Subnet = $LocalIP -replace '\.\d+$', '.'

# Output header
"IP Address,Hostname,RDP Status" | Out-File -FilePath "RDP_Computers.csv" -Encoding utf8

# Scan the subnet (1-254)
1..254 | ForEach-Object -Parallel {
    $IP = "$using:Subnet$_"
    if (Test-Port -Computer $IP -Port 3389 -Timeout 200) {
        $Hostname = Get-HostnameFromIP -IPAddress $IP
        "$IP,$Hostname,RDP Open" | Out-File -FilePath "RDP_Computers.csv" -Encoding utf8 -Append
    }
} -ThrottleLimit 64

# Display results
Write-Host "Scan complete. Results saved to RDP_Computers.csv"
Write-Host "`nDiscovered RDP-enabled computers:`n"
Get-Content "RDP_Computers.csv"