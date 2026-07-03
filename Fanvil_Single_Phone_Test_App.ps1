# ============================================================
# FANVIL SINGLE PHONE TEST PROGRAM
# Purpose : Enter one Fanvil phone IP, login, fetch information,
#           show result, and export to D:\fanvil
# Models  : X301G / X3SG / Fanvil web phones
# Login   : Default admin / admin
# ============================================================

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$script:OutputFolder = "D:\fanvil"
$script:LastResult = $null
$script:LastHtml = ""

if (!(Test-Path $script:OutputFolder)) {
    New-Item -ItemType Directory -Path $script:OutputFolder -Force | Out-Null
}

function Clean-Text {
    param([string]$Text)

    if (!$Text) { return "N/A" }

    $t = $Text `
        -replace "<script[\s\S]*?</script>", " " `
        -replace "<style[\s\S]*?</style>", " " `
        -replace "<[^>]+>", " " `
        -replace "&nbsp;", " " `
        -replace "&amp;", "&" `
        -replace "&#58;", ":" `
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

function Extract-Regex {
    param(
        [string]$Text,
        [string[]]$Patterns
    )

    foreach ($p in $Patterns) {
        $m = [regex]::Match($Text, $p, "IgnoreCase")
        if ($m.Success) {
            $v = Clean-Text $m.Groups[1].Value
            if ($v -and $v -ne "N/A") {
                return $v
            }
        }
    }

    return "N/A"
}

function Extract-AfterLabel {
    param(
        [string]$Plain,
        [string[]]$Labels
    )

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

function Test-HttpAlive {
    param(
        [string]$IP,
        [int]$Timeout
    )

    try {
        Invoke-WebRequest `
            -Uri "http://$IP" `
            -UseBasicParsing `
            -TimeoutSec $Timeout `
            -ErrorAction Stop | Out-Null

        return $true
    }
    catch {
        return $false
    }
}

function Invoke-FanvilLogin {
    param(
        [string]$IP,
        $Session,
        [string]$User,
        [string]$Pass,
        [int]$Timeout
    )

    $loginUrls = @(
        "http://$IP/cgi-bin/dologin",
        "http://$IP/login",
        "http://$IP/index.htm",
        "http://$IP/index.html",
        "http://$IP/",
        "http://$IP/cgi-bin/ConfigManApp.com"
    )

    $bodies = @(
        @{ username=$User; password=$Pass },
        @{ user=$User; pwd=$Pass },
        @{ Username=$User; Password=$Pass },
        @{ UserName=$User; Password=$Pass },
        @{ userName=$User; password=$Pass },
        @{ account=$User; password=$Pass },
        @{ login=$User; password=$Pass },
        @{ name=$User; pwd=$Pass }
    )

    foreach ($url in $loginUrls) {
        foreach ($body in $bodies) {
            try {
                Invoke-WebRequest `
                    -Uri $url `
                    -Method POST `
                    -Body $body `
                    -WebSession $Session `
                    -UseBasicParsing `
                    -TimeoutSec $Timeout `
                    -ErrorAction Stop | Out-Null
            }
            catch {}
        }
    }
}

function Get-FanvilHtml {
    param(
        [string]$IP,
        $Session,
        [int]$Timeout
    )

    $pages = @(
        "/",
        "/index.htm",
        "/index.html",
        "/status.htm",
        "/status.html",
        "/network.htm",
        "/network.html",
        "/line.htm",
        "/line.html",
        "/account.htm",
        "/account.html",
        "/cgi-bin/ConfigManApp.com?key=Status",
        "/cgi-bin/ConfigManApp.com?key=Network",
        "/cgi-bin/ConfigManApp.com?key=PhoneStatus",
        "/cgi-bin/ConfigManApp.com?key=NetworkInfo",
        "/cgi-bin/ConfigManApp.com?key=DeviceInfo",
        "/cgi-bin/ConfigManApp.com?key=Information",
        "/cgi-bin/ConfigManApp.com?key=System",
        "/cgi-bin/ConfigManApp.com?key=Line",
        "/cgi-bin/ConfigManApp.com?key=SIP",
        "/cgi-bin/ConfigManApp.com?key=Account"
    )

    $all = ""

    foreach ($p in $pages) {
        try {
            $url = "http://$IP$p"
            $r = Invoke-WebRequest `
                -Uri $url `
                -WebSession $Session `
                -UseBasicParsing `
                -TimeoutSec $Timeout `
                -ErrorAction Stop

            if ($r.Content) {
                $all += "`r`n===== PAGE: $url =====`r`n"
                $all += [string]$r.Content
            }
        }
        catch {}
    }

    return $all
}

function Get-PhoneInfo {
    param(
        [string]$IP,
        [string]$Html,
        [string]$DebugFile
    )

    $plain = Clean-Text $Html

    $model = "N/A"
    if ($plain -match "X301G") {
        $model = "X301G"
    }
    elseif ($plain -match "X3SG") {
        $model = "X3SG"
    }
    else {
        $model = Extract-AfterLabel $plain @("Model", "Phone Model", "Device Model")
    }

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

    if ($ethernetIP -eq "N/A") {
        $ethernetIP = $IP
    }

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
    if ($plain -match "Fanvil|X301G|X3SG|Ethernet MAC|Ethernet IP") {
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

    if (!$Info) {
        return
    }

    $text = @"
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
"@

    $txtResult.Text = $text
}

function Export-Result {
    if (!$script:LastResult) {
        [System.Windows.Forms.MessageBox]::Show("No test result to export.", "Export", "OK", "Information") | Out-Null
        return
    }

    if (!(Test-Path $script:OutputFolder)) {
        New-Item -ItemType Directory -Path $script:OutputFolder -Force | Out-Null
    }

    $stamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $csv = Join-Path $script:OutputFolder "Fanvil_Single_Test_$stamp.csv"
    $xlsx = Join-Path $script:OutputFolder "Fanvil_Single_Test_$stamp.xlsx"

    $script:LastResult | Export-Csv -Path $csv -NoTypeInformation -Encoding UTF8

    $excelOk = $false

    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false

        $wb = $excel.Workbooks.Open($csv)
        $ws = $wb.Worksheets.Item(1)
        $ws.Name = "Fanvil Test"

        $ws.UsedRange.EntireColumn.AutoFit() | Out-Null
        $ws.Range("A1:O1").Font.Bold = $true
        $ws.Range("A1:O1").Interior.ColorIndex = 15

        $wb.SaveAs($xlsx, 51)
        $wb.Close($true)
        $excel.Quit()

        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($ws) | Out-Null
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($wb) | Out-Null
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null

        $excelOk = $true
    }
    catch {
        $excelOk = $false
    }

    if ($excelOk) {
        Add-Log "Excel saved: $xlsx"
        [System.Windows.Forms.MessageBox]::Show("Excel saved:`r`n$xlsx`r`n`r`nCSV also saved:`r`n$csv", "Export Done", "OK", "Information") | Out-Null
    }
    else {
        Add-Log "CSV saved: $csv"
        [System.Windows.Forms.MessageBox]::Show("CSV saved:`r`n$csv", "Export Done", "OK", "Information") | Out-Null
    }
}

# ---------------- GUI ----------------

$form = New-Object System.Windows.Forms.Form
$form.Text = "Fanvil Single Phone Test"
$form.Size = New-Object System.Drawing.Size(850, 650)
$form.StartPosition = "CenterScreen"
$form.Font = New-Object System.Drawing.Font("Segoe UI", 9)

$lblTitle = New-Object System.Windows.Forms.Label
$lblTitle.Text = "Fanvil Single Phone Test"
$lblTitle.Font = New-Object System.Drawing.Font("Segoe UI", 16, [System.Drawing.FontStyle]::Bold)
$lblTitle.AutoSize = $true
$lblTitle.Location = New-Object System.Drawing.Point(15, 15)
$form.Controls.Add($lblTitle)

$lblSub = New-Object System.Windows.Forms.Label
$lblSub.Text = "Enter one LAN phone IP, test login, fetch phone information, and export to D:\fanvil"
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
$btnTest.Text = "Test Fetch"
$btnTest.Location = New-Object System.Drawing.Point(640, 78)
$btnTest.Size = New-Object System.Drawing.Size(90, 32)
$form.Controls.Add($btnTest)

$btnExport = New-Object System.Windows.Forms.Button
$btnExport.Text = "Export"
$btnExport.Location = New-Object System.Drawing.Point(740, 78)
$btnExport.Size = New-Object System.Drawing.Size(75, 32)
$form.Controls.Add($btnExport)

$lblResult = New-Object System.Windows.Forms.Label
$lblResult.Text = "Fetched Information:"
$lblResult.AutoSize = $true
$lblResult.Location = New-Object System.Drawing.Point(20, 155)
$form.Controls.Add($lblResult)

$txtResult = New-Object System.Windows.Forms.TextBox
$txtResult.Location = New-Object System.Drawing.Point(20, 180)
$txtResult.Size = New-Object System.Drawing.Size(795, 300)
$txtResult.Multiline = $true
$txtResult.ScrollBars = "Vertical"
$txtResult.ReadOnly = $true
$txtResult.Font = New-Object System.Drawing.Font("Consolas", 10)
$form.Controls.Add($txtResult)

$lblLog = New-Object System.Windows.Forms.Label
$lblLog.Text = "Log:"
$lblLog.AutoSize = $true
$lblLog.Location = New-Object System.Drawing.Point(20, 495)
$form.Controls.Add($lblLog)

$txtLog = New-Object System.Windows.Forms.TextBox
$txtLog.Location = New-Object System.Drawing.Point(20, 520)
$txtLog.Size = New-Object System.Drawing.Size(795, 80)
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

    if (!(Test-HttpAlive -IP $ip -Timeout $timeout)) {
        Add-Log "No HTTP response from $ip"
        [System.Windows.Forms.MessageBox]::Show("No HTTP response from $ip", "Test Failed", "OK", "Warning") | Out-Null
        return
    }

    Add-Log "HTTP response OK."
    Add-Log "Trying login using $user / password."

    $session = New-Object Microsoft.PowerShell.Commands.WebRequestSession

    Invoke-FanvilLogin -IP $ip -Session $session -User $user -Pass $pass -Timeout $timeout

    Add-Log "Reading possible information pages."

    $html = Get-FanvilHtml -IP $ip -Session $session -Timeout $timeout

    if (!$html) {
        Add-Log "Could not read any page after login."
        [System.Windows.Forms.MessageBox]::Show("Could not read any page after login.", "Test Failed", "OK", "Warning") | Out-Null
        return
    }

    $debugFile = Join-Path $script:OutputFolder ("Single_Debug_" + $ip.Replace(".","_") + ".html")
    $html | Out-File $debugFile -Encoding UTF8

    $info = Get-PhoneInfo -IP $ip -Html $html -DebugFile $debugFile

    $script:LastResult = $info
    $script:LastHtml = $html

    Show-Result $info

    Add-Log "Test completed."
    Add-Log "Debug HTML saved: $debugFile"

    if ($info.IsFanvil -eq "Yes") {
        Add-Log "Fanvil page detected."
    }
    else {
        Add-Log "Fanvil not confirmed. Send the debug HTML if values are N/A."
    }
})

$btnExport.Add_Click({
    Export-Result
})

Add-Log "Ready. Enter one phone IP and click Test Fetch."

[void]$form.ShowDialog()
