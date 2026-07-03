# ============================================================
# FANVIL AUTO DETECT AUDIT TOOL
# Models Supported : Fanvil X301G / Fanvil X3SG
# Purpose          : Run on each PC, find Fanvil phones in same segment,
#                    try admin/admin login, read phone pages,
#                    export confirmed Fanvil phone data to CSV.
# Output Folder    : C:\Fanvil_Audit
# Output CSV       : C:\Fanvil_Audit\Fanvil_Audit.csv
# Debug HTML       : C:\Fanvil_Audit\Debug_<IP>.html
# ============================================================

# ---------------- USER SETTINGS ----------------

$Username = "admin"
$Password = "admin"

$TimeoutSec = 3

$OutputFolder = "C:\Fanvil_Audit"
$OutputCsv    = "$OutputFolder\Fanvil_Audit.csv"

$SaveDebugHtml = $true

# Scan same /24 segment of PC IP
# Example PC IP 192.168.10.25 scans 192.168.10.1 to 192.168.10.254
$ScanStart = 1
$ScanEnd   = 254

# Supported models
$SupportedModels = @("X301G", "X3SG")

# ------------------------------------------------

Clear-Host

Write-Host "============================================================" -ForegroundColor Cyan
Write-Host " FANVIL AUTO DETECT AUDIT TOOL - X301G / X3SG" -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host ""

if (!(Test-Path $OutputFolder)) {
    New-Item -ItemType Directory -Force -Path $OutputFolder | Out-Null
}

# ---------------- FUNCTIONS ----------------

function Get-LocalIPv4 {
    try {
        $ip = Get-NetIPAddress -AddressFamily IPv4 |
            Where-Object {
                $_.IPAddress -notlike "127.*" -and
                $_.IPAddress -notlike "169.254.*" -and
                $_.IPAddress -match "^\d{1,3}(\.\d{1,3}){3}$"
            } |
            Sort-Object InterfaceMetric |
            Select-Object -First 1 -ExpandProperty IPAddress

        return $ip
    }
    catch {
        try {
            $ip = Get-WmiObject Win32_NetworkAdapterConfiguration |
                Where-Object {
                    $_.IPEnabled -eq $true -and
                    $_.IPAddress -ne $null
                } |
                ForEach-Object { $_.IPAddress } |
                Where-Object {
                    $_ -notlike "127.*" -and
                    $_ -notlike "169.254.*" -and
                    $_ -match "^\d{1,3}(\.\d{1,3}){3}$"
                } |
                Select-Object -First 1

            return $ip
        }
        catch {
            return $null
        }
    }
}

function Get-LocalMAC {
    try {
        $mac = Get-NetAdapter |
            Where-Object { $_.Status -eq "Up" } |
            Sort-Object InterfaceMetric |
            Select-Object -First 1 -ExpandProperty MacAddress

        if ($mac) { return $mac }
        return "N/A"
    }
    catch {
        try {
            $mac = Get-WmiObject Win32_NetworkAdapterConfiguration |
                Where-Object { $_.IPEnabled -eq $true -and $_.MACAddress } |
                Select-Object -First 1 -ExpandProperty MACAddress

            if ($mac) { return $mac }
            return "N/A"
        }
        catch {
            return "N/A"
        }
    }
}

function Get-IPBase {
    param([string]$IP)

    $parts = $IP.Split(".")
    return "$($parts[0]).$($parts[1]).$($parts[2])"
}

function Test-HttpDevice {
    param([string]$IP)

    try {
        $response = Invoke-WebRequest `
            -Uri "http://$IP" `
            -UseBasicParsing `
            -TimeoutSec 2 `
            -ErrorAction Stop

        return $true
    }
    catch {
        return $false
    }
}

function Invoke-FanvilLoginAttempts {
    param(
        [string]$IP,
        $Session
    )

    $loginUrls = @(
        "http://$IP/cgi-bin/dologin",
        "http://$IP/login",
        "http://$IP/index.htm",
        "http://$IP/index.html",
        "http://$IP/",
        "http://$IP/cgi-bin/ConfigManApp.com"
    )

    $loginBodies = @(
        @{ username = $Username; password = $Password },
        @{ user = $Username; pwd = $Password },
        @{ Username = $Username; Password = $Password },
        @{ UserName = $Username; Password = $Password },
        @{ userName = $Username; password = $Password },
        @{ account = $Username; password = $Password },
        @{ login = $Username; password = $Password },
        @{ name = $Username; pwd = $Password },
        @{ id = $Username; password = $Password }
    )

    foreach ($url in $loginUrls) {
        foreach ($body in $loginBodies) {
            try {
                Invoke-WebRequest `
                    -Uri $url `
                    -Method POST `
                    -Body $body `
                    -WebSession $Session `
                    -UseBasicParsing `
                    -TimeoutSec 2 `
                    -ErrorAction Stop | Out-Null
            }
            catch {
                # Continue trying other login methods
            }
        }
    }
}

function Get-FanvilPages {
    param(
        [string]$IP,
        $Session
    )

    $pages = @(
        "/",
        "/index.htm",
        "/index.html",
        "/status.htm",
        "/status.html",
        "/network.htm",
        "/network.html",
        "/cgi-bin/ConfigManApp.com?key=Status",
        "/cgi-bin/ConfigManApp.com?key=Network",
        "/cgi-bin/ConfigManApp.com?key=PhoneStatus",
        "/cgi-bin/ConfigManApp.com?key=NetworkInfo",
        "/cgi-bin/ConfigManApp.com?key=DeviceInfo",
        "/cgi-bin/ConfigManApp.com?key=Information",
        "/cgi-bin/ConfigManApp.com?key=System"
    )

    $allHtml = ""

    foreach ($page in $pages) {
        try {
            $url = "http://$IP$page"

            $response = Invoke-WebRequest `
                -Uri $url `
                -WebSession $Session `
                -UseBasicParsing `
                -TimeoutSec $TimeoutSec `
                -ErrorAction Stop

            if ($response.Content) {
                $allHtml += "`r`n`r`n===== PAGE: $url =====`r`n"
                $allHtml += $response.Content
            }
        }
        catch {
            # Ignore failed page and continue
        }
    }

    return $allHtml
}

function Convert-HtmlText {
    param([string]$Value)

    if (!$Value) { return "N/A" }

    $clean = $Value `
        -replace "<script[\s\S]*?</script>", " " `
        -replace "<style[\s\S]*?</style>", " " `
        -replace "<.*?>", " " `
        -replace "&nbsp;", " " `
        -replace "&amp;", "&" `
        -replace "&lt;", "<" `
        -replace "&gt;", ">" `
        -replace "&#58;", ":" `
        -replace "`r", " " `
        -replace "`n", " " `
        -replace "`t", " "

    $clean = [regex]::Replace($clean, "\s+", " ").Trim()

    if (!$clean) { return "N/A" }

    return $clean
}

function Extract-Value {
    param(
        [string]$Html,
        [string[]]$Patterns
    )

    foreach ($pattern in $Patterns) {
        try {
            $match = [regex]::Match($Html, $pattern, "IgnoreCase")
            if ($match.Success) {
                $value = Convert-HtmlText $match.Groups[1].Value
                if ($value -and $value -ne "N/A") {
                    return $value
                }
            }
        }
        catch {}
    }

    return "N/A"
}

function Detect-FanvilModel {
    param([string]$Html)

    if (!$Html) { return "N/A" }

    foreach ($model in $SupportedModels) {
        if ($Html -match [regex]::Escape($model)) {
            return $model
        }
    }

    $fromPage = Extract-Value $Html @(
        "Model\s*[:=]\s*</td>\s*<td[^>]*>(.*?)</td>",
        "Model\s*[:=]\s*([^<\r\n]+)",
        "Product\s*Model\s*[:=]\s*([^<\r\n]+)",
        "Device\s*Model\s*[:=]\s*([^<\r\n]+)",
        "Phone\s*Model\s*[:=]\s*([^<\r\n]+)"
    )

    if ($fromPage -match "X301G") { return "X301G" }
    if ($fromPage -match "X3SG")  { return "X3SG" }
    if ($Html -match "Fanvil")    { return "Fanvil" }

    return "N/A"
}

function Test-IsConfirmedFanvil {
    param([string]$Html)

    if (!$Html) { return $false }

    if ($Html -match "Fanvil") { return $true }
    if ($Html -match "X301G")  { return $true }
    if ($Html -match "X3SG")   { return $true }

    return $false
}

function Extract-PhoneInfo {
    param(
        [string]$Html,
        [string]$PhoneIP,
        [string]$DebugFile,
        [string]$PCIP,
        [string]$PCMAC
    )

    $model = Detect-FanvilModel $Html

    $phoneMac = Extract-Value $Html @(
        "MAC\s*Address\s*[:=]\s*</td>\s*<td[^>]*>(.*?)</td>",
        "MAC\s*Address\s*[:=]\s*([0-9A-Fa-f:\-]{12,17})",
        "MAC\s*[:=]\s*([0-9A-Fa-f:\-]{12,17})",
        "WAN\s*MAC\s*[:=]\s*([0-9A-Fa-f:\-]{12,17})",
        "LAN\s*MAC\s*[:=]\s*([0-9A-Fa-f:\-]{12,17})"
    )

    $firmware = Extract-Value $Html @(
        "Firmware\s*Version\s*[:=]\s*</td>\s*<td[^>]*>(.*?)</td>",
        "Firmware\s*Version\s*[:=]\s*([^<\r\n]+)",
        "Software\s*Version\s*[:=]\s*([^<\r\n]+)",
        "Version\s*[:=]\s*([^<\r\n]+)",
        "Application\s*Version\s*[:=]\s*([^<\r\n]+)"
    )

    $phoneIpFromPage = Extract-Value $Html @(
        "IP\s*Address\s*[:=]\s*</td>\s*<td[^>]*>(.*?)</td>",
        "IP\s*Address\s*[:=]\s*((?:\d{1,3}\.){3}\d{1,3})",
        "WAN\s*IP\s*[:=]\s*((?:\d{1,3}\.){3}\d{1,3})",
        "LAN\s*IP\s*[:=]\s*((?:\d{1,3}\.){3}\d{1,3})"
    )

    $subnetMask = Extract-Value $Html @(
        "Subnet\s*Mask\s*[:=]\s*</td>\s*<td[^>]*>(.*?)</td>",
        "Subnet\s*Mask\s*[:=]\s*((?:\d{1,3}\.){3}\d{1,3})",
        "Netmask\s*[:=]\s*((?:\d{1,3}\.){3}\d{1,3})"
    )

    $gateway = Extract-Value $Html @(
        "Gateway\s*[:=]\s*</td>\s*<td[^>]*>(.*?)</td>",
        "Gateway\s*[:=]\s*((?:\d{1,3}\.){3}\d{1,3})",
        "Default\s*Gateway\s*[:=]\s*((?:\d{1,3}\.){3}\d{1,3})"
    )

    $dns = Extract-Value $Html @(
        "Primary\s*DNS\s*[:=]\s*</td>\s*<td[^>]*>(.*?)</td>",
        "DNS\s*[:=]\s*</td>\s*<td[^>]*>(.*?)</td>",
        "DNS\s*[:=]\s*((?:\d{1,3}\.){3}\d{1,3})",
        "Primary\s*DNS\s*[:=]\s*((?:\d{1,3}\.){3}\d{1,3})"
    )

    $networkMode = Extract-Value $Html @(
        "Network\s*Mode\s*[:=]\s*</td>\s*<td[^>]*>(.*?)</td>",
        "Network\s*Mode\s*[:=]\s*([^<\r\n]+)",
        "IP\s*Mode\s*[:=]\s*([^<\r\n]+)",
        "DHCP\s*[:=]\s*([^<\r\n]+)"
    )

    $vlan = Extract-Value $Html @(
        "VLAN\s*ID\s*[:=]\s*</td>\s*<td[^>]*>(.*?)</td>",
        "VLAN\s*ID\s*[:=]\s*([^<\r\n]+)",
        "Voice\s*VLAN\s*[:=]\s*([^<\r\n]+)",
        "PC\s*VLAN\s*[:=]\s*([^<\r\n]+)"
    )

    $serial = Extract-Value $Html @(
        "Serial\s*Number\s*[:=]\s*</td>\s*<td[^>]*>(.*?)</td>",
        "Serial\s*Number\s*[:=]\s*([^<\r\n]+)",
        "SN\s*[:=]\s*([^<\r\n]+)"
    )

    return [PSCustomObject]@{
        AuditDate       = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        PCName          = $env:COMPUTERNAME
        LoggedUser      = $env:USERNAME
        PCIP            = $PCIP
        PCMAC           = $PCMAC
        PhoneIP         = $PhoneIP
        PhoneIPFromPage = $phoneIpFromPage
        PhoneMAC        = $phoneMac
        Model           = $model
        SerialNumber    = $serial
        Firmware        = $firmware
        NetworkMode     = $networkMode
        VLAN            = $vlan
        SubnetMask      = $subnetMask
        Gateway         = $gateway
        DNS             = $dns
        LoginStatus     = "Tried admin/admin"
        DebugFile       = $DebugFile
    }
}

function Save-Record {
    param($Record)

    if (Test-Path $OutputCsv) {
        $Record | Export-Csv $OutputCsv -Append -NoTypeInformation -Encoding UTF8
    }
    else {
        $Record | Export-Csv $OutputCsv -NoTypeInformation -Encoding UTF8
    }
}

# ---------------- MAIN ----------------

$pcIP  = Get-LocalIPv4
$pcMac = Get-LocalMAC

if (!$pcIP) {
    Write-Host "No valid IPv4 address found. Exiting." -ForegroundColor Red
    Pause
    exit
}

$ipBase = Get-IPBase $pcIP

Write-Host "PC Name : $env:COMPUTERNAME"
Write-Host "User    : $env:USERNAME"
Write-Host "PC IP   : $pcIP"
Write-Host "PC MAC  : $pcMac"
Write-Host "Scan    : $ipBase.$ScanStart - $ipBase.$ScanEnd"
Write-Host "Output  : $OutputCsv"
Write-Host ""

$confirmedCount = 0
$httpCount = 0

for ($i = $ScanStart; $i -le $ScanEnd; $i++) {

    $testIP = "$ipBase.$i"

    if ($testIP -eq $pcIP) {
        continue
    }

    Write-Host "Checking $testIP ..." -NoNewline

    $isHttp = Test-HttpDevice $testIP

    if (!$isHttp) {
        Write-Host " no response" -ForegroundColor DarkGray
        continue
    }

    $httpCount++
    Write-Host " HTTP device found, trying Fanvil login..." -ForegroundColor Yellow

    $session = New-Object Microsoft.PowerShell.Commands.WebRequestSession

    Invoke-FanvilLoginAttempts -IP $testIP -Session $session

    $html = Get-FanvilPages -IP $testIP -Session $session

    if (!$html) {
        Write-Host "  No readable page after login try." -ForegroundColor DarkGray
        continue
    }

    $safeIp = $testIP.Replace(".", "_")
    $debugFile = "$OutputFolder\Debug_$safeIp.html"

    if ($SaveDebugHtml) {
        $html | Out-File $debugFile -Encoding UTF8
    }
    else {
        $debugFile = "Disabled"
    }

    if (!(Test-IsConfirmedFanvil $html)) {
        Write-Host "  Not confirmed Fanvil device." -ForegroundColor DarkGray
        continue
    }

    $record = Extract-PhoneInfo `
        -Html $html `
        -PhoneIP $testIP `
        -DebugFile $debugFile `
        -PCIP $pcIP `
        -PCMAC $pcMac

    Save-Record $record
    $confirmedCount++

    Write-Host "  Confirmed Fanvil: $($record.Model) at $testIP" -ForegroundColor Green
}

Write-Host ""
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host " Scan Completed" -ForegroundColor Cyan
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host "HTTP devices found     : $httpCount"
Write-Host "Fanvil phones confirmed: $confirmedCount"
Write-Host "CSV saved at           : $OutputCsv"
Write-Host "Debug folder           : $OutputFolder"
Write-Host ""

if ($confirmedCount -eq 0) {
    Write-Host "No Fanvil phones confirmed." -ForegroundColor Yellow
    Write-Host "If your phone login page does not show Fanvil/model before login,"
    Write-Host "send one Debug_<IP>.html file from C:\Fanvil_Audit for exact firmware support."
}

Pause
