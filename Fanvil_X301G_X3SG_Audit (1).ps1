# ============================================================
# FANVIL LAN PHONE AUDIT TOOL
# Supported Models: X301G, X3SG
# Run this script on each PC
# Collects nearby Fanvil phone info and saves to CSV
# ============================================================

# ---------------- USER SETTINGS ----------------

$Username = "admin"
$Password = "admin"

# Master CSV save path
$OutputFolder = "C:\Fanvil_Audit"
$OutputFile   = "$OutputFolder\Fanvil_Phone_Audit_Master.csv"

# Save detected webpage HTML for checking/fine tuning
$SaveDebugHtml = $true

# Scan full subnet of the PC IP
$ScanFullSubnet = $true

# Timeout seconds per IP/page
$TimeoutSec = 3

# Supported Fanvil models
$SupportedModels = @("X301G", "X3SG")

# ------------------------------------------------

if (!(Test-Path $OutputFolder)) {
    New-Item -ItemType Directory -Path $OutputFolder -Force | Out-Null
}

function Get-LocalIPv4 {
    try {
        $ips = Get-NetIPAddress -AddressFamily IPv4 |
            Where-Object {
                $_.IPAddress -notlike "127.*" -and
                $_.IPAddress -notlike "169.254.*" -and
                $_.PrefixOrigin -ne "WellKnown"
            } |
            Select-Object -ExpandProperty IPAddress

        return $ips | Select-Object -First 1
    }
    catch {
        return $null
    }
}

function Get-IPBase {
    param([string]$IP)

    $parts = $IP.Split(".")
    return "$($parts[0]).$($parts[1]).$($parts[2])"
}

function Test-WebAlive {
    param([string]$IP)

    try {
        $url = "http://$IP"
        $res = Invoke-WebRequest -Uri $url -UseBasicParsing -TimeoutSec $TimeoutSec -ErrorAction Stop
        return $res.Content
    }
    catch {
        return $null
    }
}

function Test-IsFanvil {
    param([string]$Content)

    if (!$Content) {
        return $false
    }

    if ($Content -match "Fanvil|X301G|X3SG|VoIP|SIP Phone|IP Phone") {
        return $true
    }

    return $false
}

function Try-FanvilLogin {
    param([string]$IP)

    $session = New-Object Microsoft.PowerShell.Commands.WebRequestSession

    $loginUrls = @(
        "http://$IP/cgi-bin/dologin",
        "http://$IP/cgi-bin/ConfigManApp.com",
        "http://$IP/login",
        "http://$IP/",
        "http://$IP/index.htm",
        "http://$IP/index.html"
    )

    $loginBodies = @(
        @{ username = $Username; password = $Password },
        @{ user = $Username; pwd = $Password },
        @{ Username = $Username; Password = $Password },
        @{ account = $Username; password = $Password },
        @{ login = $Username; password = $Password },
        @{ userName = $Username; password = $Password }
    )

    foreach ($loginUrl in $loginUrls) {
        foreach ($body in $loginBodies) {
            try {
                Invoke-WebRequest `
                    -Uri $loginUrl `
                    -Method POST `
                    -Body $body `
                    -WebSession $session `
                    -UseBasicParsing `
                    -TimeoutSec $TimeoutSec `
                    -ErrorAction Stop | Out-Null

                return $session
            }
            catch {
                continue
            }
        }
    }

    return $session
}

function Get-FanvilPages {
    param(
        [string]$IP,
        $Session
    )

    $pages = @(
        "http://$IP/",
        "http://$IP/index.htm",
        "http://$IP/index.html",
        "http://$IP/status.htm",
        "http://$IP/status.html",
        "http://$IP/network.htm",
        "http://$IP/network.html",
        "http://$IP/cgi-bin/ConfigManApp.com?key=Status",
        "http://$IP/cgi-bin/ConfigManApp.com?key=Network",
        "http://$IP/cgi-bin/ConfigManApp.com?key=PhoneStatus",
        "http://$IP/cgi-bin/ConfigManApp.com?key=NetworkInfo"
    )

    $allContent = ""

    foreach ($page in $pages) {
        try {
            $res = Invoke-WebRequest `
                -Uri $page `
                -WebSession $Session `
                -UseBasicParsing `
                -TimeoutSec $TimeoutSec `
                -ErrorAction Stop

            if ($res.Content) {
                $allContent += "`n`n===== PAGE: $page =====`n"
                $allContent += $res.Content
            }
        }
        catch {
            continue
        }
    }

    return $allContent
}

function Extract-Value {
    param(
        [string]$Content,
        [string[]]$Patterns
    )

    foreach ($pattern in $Patterns) {
        $match = [regex]::Match($Content, $pattern, "IgnoreCase")
        if ($match.Success) {
            return ($match.Groups[1].Value -replace "<.*?>", "" -replace "&nbsp;", " " -replace "&amp;", "&" -replace "`r|`n|`t", " ").Trim()
        }
    }

    return "N/A"
}

function Detect-Model {
    param([string]$Content)

    foreach ($model in $SupportedModels) {
        if ($Content -match $model) {
            return $model
        }
    }

    $modelFromPage = Extract-Value $Content @(
        "Model\s*[:=]</td>\s*<td[^>]*>(.*?)</td>",
        "Model\s*[:=]\s*([^<\r\n]+)",
        "Product\s*Model\s*[:=]\s*([^<\r\n]+)",
        "Device\s*Model\s*[:=]\s*([^<\r\n]+)"
    )

    return $modelFromPage
}

function Get-PCInfo {
    $pcName = $env:COMPUTERNAME
    $userName = $env:USERNAME
    $ip = Get-LocalIPv4

    $mac = "N/A"
    try {
        $mac = Get-NetAdapter |
            Where-Object { $_.Status -eq "Up" } |
            Select-Object -First 1 -ExpandProperty MacAddress
    }
    catch {}

    return [PSCustomObject]@{
        PCName = $pcName
        UserName = $userName
        PCIP = $ip
        PCMAC = $mac
    }
}

# ---------------- MAIN PROCESS ----------------

$pcInfo = Get-PCInfo

if (!$pcInfo.PCIP) {
    Write-Host "No valid PC IP found." -ForegroundColor Red
    exit
}

$ipBase = Get-IPBase $pcInfo.PCIP

Write-Host "============================================="
Write-Host " FANVIL PHONE AUDIT TOOL - X301G / X3SG"
Write-Host "============================================="
Write-Host "PC Name : $($pcInfo.PCName)"
Write-Host "PC IP   : $($pcInfo.PCIP)"
Write-Host "PC MAC  : $($pcInfo.PCMAC)"
Write-Host "Scanning subnet: $ipBase.1 - $ipBase.254"
Write-Host ""

$possiblePhones = @()

if ($ScanFullSubnet) {
    for ($i = 1; $i -le 254; $i++) {
        $testIP = "$ipBase.$i"

        if ($testIP -eq $pcInfo.PCIP) {
            continue
        }

        Write-Host "Checking $testIP ..." -NoNewline

        $content = Test-WebAlive $testIP

        if ($content) {
            if (Test-IsFanvil $content) {
                Write-Host " possible Fanvil found" -ForegroundColor Green
                $possiblePhones += $testIP
            }
            else {
                Write-Host " web device found" -ForegroundColor Yellow
            }
        }
        else {
            Write-Host " no response" -ForegroundColor DarkGray
        }
    }
}

if ($possiblePhones.Count -eq 0) {
    Write-Host ""
    Write-Host "No Fanvil phone found automatically." -ForegroundColor Yellow
    $manualIP = Read-Host "Enter phone IP manually or press Enter to exit"

    if ($manualIP) {
        $possiblePhones += $manualIP
    }
    else {
        Write-Host "Completed. No phone IP entered."
        pause
        exit
    }
}

foreach ($phoneIP in $possiblePhones) {

    Write-Host ""
    Write-Host "Trying phone: $phoneIP" -ForegroundColor Cyan

    $session = Try-FanvilLogin $phoneIP

    if (!$session) {
        Write-Host "Login/session failed for $phoneIP" -ForegroundColor Red
        continue
    }

    Write-Host "Reading phone pages..." -ForegroundColor Green

    $content = Get-FanvilPages -IP $phoneIP -Session $session

    if (!$content) {
        Write-Host "No content received after login/page read." -ForegroundColor Red
        continue
    }

    $detectedModel = Detect-Model $content

    if ($SaveDebugHtml) {
        $safeIP = $phoneIP.Replace(".", "_")
        $debugFile = "$OutputFolder\Debug_${safeIP}_${detectedModel}.html"
        $content | Out-File $debugFile -Encoding UTF8
    }

    $phoneMac = Extract-Value $content @(
        "MAC\s*Address\s*[:=]</td>\s*<td[^>]*>(.*?)</td>",
        "MAC\s*Address\s*[:=]\s*([0-9A-Fa-f:\-]{12,17})",
        "MAC\s*[:=]\s*([0-9A-Fa-f:\-]{12,17})"
    )

    $firmware = Extract-Value $content @(
        "Firmware\s*Version\s*[:=]</td>\s*<td[^>]*>(.*?)</td>",
        "Firmware\s*Version\s*[:=]\s*([^<\r\n]+)",
        "Software\s*Version\s*[:=]\s*([^<\r\n]+)",
        "Version\s*[:=]\s*([^<\r\n]+)"
    )

    $networkMode = Extract-Value $content @(
        "Network\s*Mode\s*[:=]</td>\s*<td[^>]*>(.*?)</td>",
        "Network\s*Mode\s*[:=]\s*([^<\r\n]+)",
        "IP\s*Mode\s*[:=]\s*([^<\r\n]+)"
    )

    $phoneIpFromPage = Extract-Value $content @(
        "IP\s*Address\s*[:=]</td>\s*<td[^>]*>(.*?)</td>",
        "IP\s*Address\s*[:=]\s*((?:\d{1,3}\.){3}\d{1,3})"
    )

    $subnetMask = Extract-Value $content @(
        "Subnet\s*Mask\s*[:=]</td>\s*<td[^>]*>(.*?)</td>",
        "Subnet\s*Mask\s*[:=]\s*((?:\d{1,3}\.){3}\d{1,3})"
    )

    $gateway = Extract-Value $content @(
        "Gateway\s*[:=]</td>\s*<td[^>]*>(.*?)</td>",
        "Gateway\s*[:=]\s*((?:\d{1,3}\.){3}\d{1,3})"
    )

    $dns = Extract-Value $content @(
        "DNS\s*[:=]</td>\s*<td[^>]*>(.*?)</td>",
        "DNS\s*[:=]\s*((?:\d{1,3}\.){3}\d{1,3})"
    )

    $record = [PSCustomObject]@{
        AuditDate       = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        PCName          = $pcInfo.PCName
        LoggedUser      = $pcInfo.UserName
        PCIP            = $pcInfo.PCIP
        PCMAC           = $pcInfo.PCMAC
        PhoneIP         = $phoneIP
        PhoneIPFromPage = $phoneIpFromPage
        PhoneMAC        = $phoneMac
        Model           = $detectedModel
        Firmware        = $firmware
        NetworkMode     = $networkMode
        SubnetMask      = $subnetMask
        Gateway         = $gateway
        DNS             = $dns
        LoginStatus     = "Success"
    }

    if (!(Test-Path $OutputFile)) {
        $record | Export-Csv $OutputFile -NoTypeInformation -Encoding UTF8
    }
    else {
        $record | Export-Csv $OutputFile -NoTypeInformation -Encoding UTF8 -Append
    }

    Write-Host ""
    Write-Host "Data saved successfully:" -ForegroundColor Green
    Write-Host $OutputFile
}

Write-Host ""
Write-Host "Completed." -ForegroundColor Cyan
pause
