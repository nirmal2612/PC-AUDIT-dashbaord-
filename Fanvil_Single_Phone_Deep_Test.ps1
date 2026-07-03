# ============================================================
# FANVIL SINGLE PHONE TEST - DEEP PAGE READER
# Fix: Login success but page not reading / values N/A
# This version:
# 1. Opens one phone IP
# 2. Tries login admin/admin
# 3. Reads root/index/status pages
# 4. Auto-detects links, frames, iframes, scripts from HTML
# 5. Reads discovered pages also
# 6. Extracts Model, Ethernet MAC, Ethernet IP, Gateway, Software, Extension
# 7. Saves all debug pages to D:\fanvil\Fanvil_Debug_Pages
# ============================================================

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$script:OutputFolder = "D:\fanvil"
$script:DebugFolder = Join-Path $script:OutputFolder "Fanvil_Debug_Pages"
$script:LastResult = $null
$script:LastHtml = ""

New-Item -ItemType Directory -Path $script:OutputFolder -Force | Out-Null
New-Item -ItemType Directory -Path $script:DebugFolder -Force | Out-Null

function Clean-Text {
    param([string]$Text)
    if (!$Text) { return "N/A" }

    $t = $Text `
        -replace "<script[\s\S]*?</script>", " " `
        -replace "<style[\s\S]*?</style>", " " `
        -replace "<[^>]+>", " " `
        -replace "&nbsp;", " " `
        -replace "&amp;", "&" `
        -replace "&lt;", "<" `
        -replace "&gt;", ">" `
        -replace "&#58;", ":" `
        -replace "&#x3a;", ":" `
        -replace "`r", " " `
        -replace "`n", " " `
        -replace "`t", " "

    $t = [regex]::Replace($t, "\s+", " ").Trim()
    if (!$t) { return "N/A" }
    return $t
}

function Add-Log {
    param([string]$Text)
    $time = Get-Date -Format "HH:mm:ss"
    $txtLog.AppendText("[$time] $Text`r`n")
    $txtLog.SelectionStart = $txtLog.Text.Length
    $txtLog.ScrollToCaret()
    [System.Windows.Forms.Application]::DoEvents()
}

function Save-DebugPage {
    param(
        [string]$IP,
        [string]$Name,
        [string]$Content
    )

    try {
        $safeName = $Name -replace "[:\\\/\?\&\=\.\s]+", "_"
        if ($safeName.Length -gt 120) { $safeName = $safeName.Substring(0,120) }
        $file = Join-Path $script:DebugFolder ("Debug_" + $IP.Replace(".","_") + "_" + $safeName + ".html")
        $Content | Out-File $file -Encoding UTF8
        return $file
    } catch {
        return "Save failed"
    }
}

function Extract-Regex {
    param([string]$Text, [string[]]$Patterns)

    foreach ($p in $Patterns) {
        $m = [regex]::Match($Text, $p, "IgnoreCase")
        if ($m.Success) {
            $v = Clean-Text $m.Groups[1].Value
            if ($v -and $v -ne "N/A") { return $v }
        }
    }

    return "N/A"
}

function Extract-AfterLabel {
    param([string]$Plain, [string[]]$Labels)

    foreach ($label in $Labels) {
        $pattern = [regex]::Escape($label) + "\s*:?\s*([A-Za-z0-9\.\-_:\/@]+(?:\s+[A-Za-z0-9\.\-_:\/@]+)?)"
        $m = [regex]::Match($Plain, $pattern, "IgnoreCase")
        if ($m.Success) {
            $v = $m.Groups[1].Value.Trim()
            if ($v) { return $v }
        }
    }

    return "N/A"
}

function Invoke-GetPage {
    param(
        [string]$Url,
        $Session,
        [int]$Timeout
    )

    try {
        $headers = @{
            "User-Agent" = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/120 Safari/537.36"
            "Accept" = "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8"
            "Referer" = $Url
        }

        $r = Invoke-WebRequest `
            -Uri $Url `
            -WebSession $Session `
            -Headers $headers `
            -UseBasicParsing `
            -TimeoutSec $Timeout `
            -ErrorAction Stop

        if ($r.Content) {
            return [string]$r.Content
        }
    } catch {}

    return ""
}

function Invoke-PostLogin {
    param(
        [string]$IP,
        $Session,
        [string]$User,
        [string]$Pass,
        [int]$Timeout
    )

    $loginUrls = @(
        "http://$IP/cgi-bin/dologin",
        "http://$IP/cgi-bin/ConfigManApp.com",
        "http://$IP/login",
        "http://$IP/",
        "http://$IP/index.htm",
        "http://$IP/index.html"
    )

    $bodies = @(
        @{ username=$User; password=$Pass },
        @{ user=$User; pwd=$Pass },
        @{ User=$User; Password=$Pass },
        @{ Username=$User; Password=$Pass },
        @{ userName=$User; password=$Pass },
        @{ account=$User; password=$Pass },
        @{ login=$User; password=$Pass },
        @{ Login=$User; Password=$Pass },
        @{ name=$User; pwd=$Pass },
        @{ admin=$User; password=$Pass }
    )

    foreach ($url in $loginUrls) {
        foreach ($body in $bodies) {
            try {
                $headers = @{
                    "User-Agent" = "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 Chrome/120 Safari/537.36"
                    "Accept" = "*/*"
                    "Referer" = "http://$IP/"
                }

                Invoke-WebRequest `
                    -Uri $url `
                    -Method POST `
                    -Body $body `
                    -WebSession $Session `
                    -Headers $headers `
                    -UseBasicParsing `
                    -TimeoutSec $Timeout `
                    -ErrorAction Stop | Out-Null
            } catch {}
        }
    }
}

function Get-DiscoveredLinks {
    param(
        [string]$IP,
        [string]$Html
    )

    $links = New-Object System.Collections.ArrayList

    $patterns = @(
        'href\s*=\s*["'']([^"'']+)["'']',
        'src\s*=\s*["'']([^"'']+)["'']',
        'action\s*=\s*["'']([^"'']+)["'']',
        'url\s*:\s*["'']([^"'']+)["'']'
    )

    foreach ($pattern in $patterns) {
        $matches = [regex]::Matches($Html, $pattern, "IgnoreCase")
        foreach ($m in $matches) {
            $link = $m.Groups[1].Value.Trim()
            if (!$link) { continue }
            if ($link.StartsWith("#")) { continue }
            if ($link -match "javascript:|mailto:|tel:") { continue }

            if ($link.StartsWith("http://") -or $link.StartsWith("https://")) {
                if ($link -match [regex]::Escape($IP)) {
                    [void]$links.Add($link)
                }
            }
            elseif ($link.StartsWith("/")) {
                [void]$links.Add("http://$IP$link")
            }
            else {
                [void]$links.Add("http://$IP/$link")
            }
        }
    }

    # Fanvil common hidden/API pages
    $common = @(
        "http://$IP/",
        "http://$IP/index.htm",
        "http://$IP/index.html",
        "http://$IP/status.htm",
        "http://$IP/status.html",
        "http://$IP/information.htm",
        "http://$IP/information.html",
        "http://$IP/network.htm",
        "http://$IP/network.html",
        "http://$IP/line.htm",
        "http://$IP/line.html",
        "http://$IP/account.htm",
        "http://$IP/account.html",
        "http://$IP/cgi-bin/ConfigManApp.com?key=Status",
        "http://$IP/cgi-bin/ConfigManApp.com?key=Network",
        "http://$IP/cgi-bin/ConfigManApp.com?key=PhoneStatus",
        "http://$IP/cgi-bin/ConfigManApp.com?key=NetworkInfo",
        "http://$IP/cgi-bin/ConfigManApp.com?key=DeviceInfo",
        "http://$IP/cgi-bin/ConfigManApp.com?key=Information",
        "http://$IP/cgi-bin/ConfigManApp.com?key=System",
        "http://$IP/cgi-bin/ConfigManApp.com?key=Line",
        "http://$IP/cgi-bin/ConfigManApp.com?key=SIP",
        "http://$IP/cgi-bin/ConfigManApp.com?key=Account",
        "http://$IP/cgi-bin/ConfigManApp.com?key=Account1",
        "http://$IP/cgi-bin/ConfigManApp.com?key=RegisterStatus"
    )

    foreach ($c in $common) { [void]$links.Add($c) }

    return $links | Select-Object -Unique
}

function Get-DeepFanvilHtml {
    param(
        [string]$IP,
        $Session,
        [int]$Timeout
    )

    $all = ""
    $visited = @{}

    $startPages = @(
        "http://$IP/",
        "http://$IP/index.htm",
        "http://$IP/index.html"
    )

    foreach ($u in $startPages) {
        Add-Log "Reading start page: $u"
        $h = Invoke-GetPage -Url $u -Session $Session -Timeout $Timeout
        if ($h) {
            $visited[$u] = $true
            Save-DebugPage -IP $IP -Name $u -Content $h | Out-Null
            $all += "`r`n===== PAGE: $u =====`r`n$h"
        }
    }

    $links = Get-DiscoveredLinks -IP $IP -Html $all
    Add-Log "Discovered page count: $($links.Count)"

    foreach ($u in $links) {
        if ($visited.ContainsKey($u)) { continue }
        $visited[$u] = $true

        Add-Log "Reading: $u"
        $h = Invoke-GetPage -Url $u -Session $Session -Timeout $Timeout

        if ($h) {
            Save-DebugPage -IP $IP -Name $u -Content $h | Out-Null
            $all += "`r`n===== PAGE: $u =====`r`n$h"
        }
    }

    return $all
}

function Get-PhoneInfo {
    param([string]$IP, [string]$Html, [string]$DebugFile)

    $plain = Clean-Text $Html

    $model = "N/A"
    if ($plain -match "X301G") { $model = "X301G" }
    elseif ($plain -match "X3SG") { $model = "X3SG" }
    else { $model = Extract-AfterLabel $plain @("Model", "Phone Model", "Device Model") }

    $hardware = Extract-AfterLabel $plain @("Hardware")
    $software = Extract-AfterLabel $plain @("Software", "Firmware Version", "Firmware", "Version")
    $uboot = Extract-AfterLabel $plain @("Uboot")
    $uptime = Extract-Regex $plain @(
        "Uptime\s*:?\s*([0-9]+\s*:\s*[0-9]+\s*:\s*[0-9]+)",
        "Running\s*time\s*:?\s*([0-9]+\s*:\s*[0-9]+\s*:\s*[0-9]+)"
    )

    $networkMode = Extract-AfterLabel $plain @("Network mode", "Network Mode", "IP Mode")

    $ethernetMac = Extract-Regex $plain @(
        "Ethernet\s*MAC\s*:?\s*([0-9A-Fa-f]{2}[:-][0-9A-Fa-f]{2}[:-][0-9A-Fa-f]{2}[:-][0-9A-Fa-f]{2}[:-][0-9A-Fa-f]{2}[:-][0-9A-Fa-f]{2})",
        "MAC\s*Address\s*:?\s*([0-9A-Fa-f]{2}[:-][0-9A-Fa-f]{2}[:-][0-9A-Fa-f]{2}[:-][0-9A-Fa-f]{2}[:-][0-9A-Fa-f]{2}[:-][0-9A-Fa-f]{2})",
        "MAC\s*:?\s*([0-9A-Fa-f]{2}[:-][0-9A-Fa-f]{2}[:-][0-9A-Fa-f]{2}[:-][0-9A-Fa-f]{2}[:-][0-9A-Fa-f]{2}[:-][0-9A-Fa-f]{2})"
    )

    $ethernetIP = Extract-Regex $plain @(
        "Ethernet\s*IP\s*:?\s*((?:\d{1,3}\.){3}\d{1,3})",
        "IP\s*Address\s*:?\s*((?:\d{1,3}\.){3}\d{1,3})",
        "IPv4\s*:?\s*((?:\d{1,3}\.){3}\d{1,3})"
    )
    if ($ethernetIP -eq "N/A") { $ethernetIP = $IP }

    $subnet = Extract-Regex $plain @(
        "Subnet\s*mask\s*:?\s*((?:\d{1,3}\.){3}\d{1,3})",
        "Subnet\s*Mask\s*:?\s*((?:\d{1,3}\.){3}\d{1,3})",
        "Netmask\s*:?\s*((?:\d{1,3}\.){3}\d{1,3})"
    )

    $gateway = Extract-Regex $plain @(
        "Default\s*gateway\s*:?\s*((?:\d{1,3}\.){3}\d{1,3})",
        "Default\s*Gateway\s*:?\s*((?:\d{1,3}\.){3}\d{1,3})",
        "Gateway\s*:?\s*((?:\d{1,3}\.){3}\d{1,3})"
    )

    $extension = Extract-Regex $plain @(
        "Extension\s*:?\s*([0-9]{2,10})",
        "Account\s*1\s*:?\s*([0-9]{2,10})",
        "SIP\s*User\s*ID\s*:?\s*([0-9]{2,10})",
        "User\s*ID\s*:?\s*([0-9]{2,10})",
        "Register\s*Name\s*:?\s*([0-9]{2,10})",
        "Phone\s*Number\s*:?\s*([0-9]{2,10})",
        "Auth\s*User\s*:?\s*([0-9]{2,10})"
    )

    $lineStatus = Extract-Regex $plain @(
        "Register\s*Status\s*:?\s*([A-Za-z0-9\s]+)",
        "Registration\s*Status\s*:?\s*([A-Za-z0-9\s]+)",
        "SIP\s*Status\s*:?\s*([A-Za-z0-9\s]+)"
    )

    $isFanvil = "No"
    if ($plain -match "Fanvil|X301G|X3SG|Ethernet MAC|Ethernet IP|Default gateway") {
        $isFanvil = "Yes"
    }

    return [PSCustomObject]@{
        TestDate       = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        PhoneIP        = $IP
        IsFanvil       = $isFanvil
        Model          = $model
        Extension      = $extension
        EthernetMAC    = $ethernetMac
        EthernetIP     = $ethernetIP
        NetworkMode    = $networkMode
        SubnetMask     = $subnet
        DefaultGateway = $gateway
        Hardware       = $hardware
        Software       = $software
        Uboot          = $uboot
        Uptime         = $uptime
        LineStatus     = $lineStatus
        DebugFile      = $DebugFile
    }
}

function Show-Result {
    param($Info)

    $txtResult.Clear()
    if (!$Info) { return }

    $txtResult.Text = @"
Test Date       : $($Info.TestDate)
Phone IP        : $($Info.PhoneIP)
Is Fanvil       : $($Info.IsFanvil)

Model           : $($Info.Model)
Extension       : $($Info.Extension)

Ethernet MAC    : $($Info.EthernetMAC)
Ethernet IP     : $($Info.EthernetIP)
Network Mode    : $($Info.NetworkMode)
Subnet Mask     : $($Info.SubnetMask)
Default Gateway : $($Info.DefaultGateway)

Hardware        : $($Info.Hardware)
Software        : $($Info.Software)
Uboot           : $($Info.Uboot)
Uptime          : $($Info.Uptime)
Line Status     : $($Info.LineStatus)

Debug File      : $($Info.DebugFile)
Debug Folder    : $script:DebugFolder
"@
}

function Export-Result {
    if (!$script:LastResult) {
        [System.Windows.Forms.MessageBox]::Show("No result to export.", "Export", "OK", "Information") | Out-Null
        return
    }

    $stamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $csv = Join-Path $script:OutputFolder "Fanvil_Single_Deep_Test_$stamp.csv"
    $script:LastResult | Export-Csv -Path $csv -NoTypeInformation -Encoding UTF8
    [System.Windows.Forms.MessageBox]::Show("CSV saved:`r`n$csv", "Export Done", "OK", "Information") | Out-Null
}

$form = New-Object System.Windows.Forms.Form
$form.Text = "Fanvil Single Phone Deep Test"
$form.Size = New-Object System.Drawing.Size(900, 680)
$form.StartPosition = "CenterScreen"
$form.Font = New-Object System.Drawing.Font("Segoe UI", 9)

$lblTitle = New-Object System.Windows.Forms.Label
$lblTitle.Text = "Fanvil Single Phone Deep Test"
$lblTitle.Font = New-Object System.Drawing.Font("Segoe UI", 16, [System.Drawing.FontStyle]::Bold)
$lblTitle.AutoSize = $true
$lblTitle.Location = New-Object System.Drawing.Point(15, 15)
$form.Controls.Add($lblTitle)

$lblSub = New-Object System.Windows.Forms.Label
$lblSub.Text = "For login success but page not reading. It discovers and reads links/frames after login."
$lblSub.AutoSize = $true
$lblSub.Location = New-Object System.Drawing.Point(18, 48)
$form.Controls.Add($lblSub)

$lblIP = New-Object System.Windows.Forms.Label
$lblIP.Text = "Phone IP:"
$lblIP.AutoSize = $true
$lblIP.Location = New-Object System.Drawing.Point(20, 85)
$form.Controls.Add($lblIP)

$txtIP = New-Object System.Windows.Forms.TextBox
$txtIP.Location = New-Object System.Drawing.Point(100, 81)
$txtIP.Size = New-Object System.Drawing.Size(160, 25)
$form.Controls.Add($txtIP)

$lblUser = New-Object System.Windows.Forms.Label
$lblUser.Text = "User:"
$lblUser.AutoSize = $true
$lblUser.Location = New-Object System.Drawing.Point(280, 85)
$form.Controls.Add($lblUser)

$txtUser = New-Object System.Windows.Forms.TextBox
$txtUser.Location = New-Object System.Drawing.Point(325, 81)
$txtUser.Size = New-Object System.Drawing.Size(100, 25)
$txtUser.Text = "admin"
$form.Controls.Add($txtUser)

$lblPass = New-Object System.Windows.Forms.Label
$lblPass.Text = "Password:"
$lblPass.AutoSize = $true
$lblPass.Location = New-Object System.Drawing.Point(440, 85)
$form.Controls.Add($lblPass)

$txtPass = New-Object System.Windows.Forms.TextBox
$txtPass.Location = New-Object System.Drawing.Point(515, 81)
$txtPass.Size = New-Object System.Drawing.Size(100, 25)
$txtPass.Text = "admin"
$form.Controls.Add($txtPass)

$lblTimeout = New-Object System.Windows.Forms.Label
$lblTimeout.Text = "Timeout:"
$lblTimeout.AutoSize = $true
$lblTimeout.Location = New-Object System.Drawing.Point(20, 122)
$form.Controls.Add($lblTimeout)

$numTimeout = New-Object System.Windows.Forms.NumericUpDown
$numTimeout.Location = New-Object System.Drawing.Point(100, 118)
$numTimeout.Size = New-Object System.Drawing.Size(60, 25)
$numTimeout.Minimum = 1
$numTimeout.Maximum = 10
$numTimeout.Value = 3
$form.Controls.Add($numTimeout)

$btnTest = New-Object System.Windows.Forms.Button
$btnTest.Text = "Deep Test"
$btnTest.Location = New-Object System.Drawing.Point(650, 78)
$btnTest.Size = New-Object System.Drawing.Size(95, 32)
$form.Controls.Add($btnTest)

$btnExport = New-Object System.Windows.Forms.Button
$btnExport.Text = "Export"
$btnExport.Location = New-Object System.Drawing.Point(755, 78)
$btnExport.Size = New-Object System.Drawing.Size(75, 32)
$form.Controls.Add($btnExport)

$txtResult = New-Object System.Windows.Forms.TextBox
$txtResult.Location = New-Object System.Drawing.Point(20, 165)
$txtResult.Size = New-Object System.Drawing.Size(840, 300)
$txtResult.Multiline = $true
$txtResult.ScrollBars = "Vertical"
$txtResult.ReadOnly = $true
$txtResult.Font = New-Object System.Drawing.Font("Consolas", 10)
$form.Controls.Add($txtResult)

$txtLog = New-Object System.Windows.Forms.TextBox
$txtLog.Location = New-Object System.Drawing.Point(20, 485)
$txtLog.Size = New-Object System.Drawing.Size(840, 130)
$txtLog.Multiline = $true
$txtLog.ScrollBars = "Vertical"
$txtLog.ReadOnly = $true
$form.Controls.Add($txtLog)

$btnTest.Add_Click({
    $ip = $txtIP.Text.Trim()
    $user = $txtUser.Text.Trim()
    $pass = $txtPass.Text.Trim()
    $timeout = [int]$numTimeout.Value

    if (!$ip -or $ip -notmatch "^(\d{1,3}\.){3}\d{1,3}$") {
        [System.Windows.Forms.MessageBox]::Show("Enter valid phone IP. Example: 10.209.110.202", "Invalid IP", "OK", "Warning") | Out-Null
        return
    }

    $txtResult.Clear()
    $txtLog.Clear()
    $script:LastResult = $null
    $script:LastHtml = ""

    Add-Log "Testing IP: $ip"

    $session = New-Object Microsoft.PowerShell.Commands.WebRequestSession

    Add-Log "Reading before login."
    $before = Invoke-GetPage -Url "http://$ip/" -Session $session -Timeout $timeout
    if ($before) {
        Save-DebugPage -IP $ip -Name "before_login_root" -Content $before | Out-Null
        Add-Log "Before-login root page read."
    } else {
        Add-Log "No root page before login."
    }

    Add-Log "Trying login..."
    Invoke-PostLogin -IP $ip -Session $session -User $user -Pass $pass -Timeout $timeout

    Add-Log "Deep reading pages after login..."
    $html = Get-DeepFanvilHtml -IP $ip -Session $session -Timeout $timeout

    if (!$html) {
        Add-Log "Could not read pages after login."
        [System.Windows.Forms.MessageBox]::Show("Could not read pages after login.", "Test Failed", "OK", "Warning") | Out-Null
        return
    }

    $combinedFile = Save-DebugPage -IP $ip -Name "combined_after_login_all_pages" -Content $html
    $script:LastHtml = $html

    $info = Get-PhoneInfo -IP $ip -Html $html -DebugFile $combinedFile
    $script:LastResult = $info
    Show-Result $info

    Add-Log "Deep test completed."
    Add-Log "Combined debug file: $combinedFile"

    if ($info.EthernetMAC -eq "N/A" -or $info.Model -eq "N/A") {
        Add-Log "Some values still N/A. Send combined debug HTML from D:\fanvil\Fanvil_Debug_Pages."
    } else {
        Add-Log "Main values fetched successfully."
    }
})

$btnExport.Add_Click({ Export-Result })

Add-Log "Ready. Enter one phone IP and click Deep Test."

[void]$form.ShowDialog()
