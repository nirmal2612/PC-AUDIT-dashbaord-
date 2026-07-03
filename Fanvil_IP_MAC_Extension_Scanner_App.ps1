# ============================================================
# FANVIL SEGMENT SCANNER APPLICATION
# Purpose:
#   Enter one IP segment and scan all Fanvil phones.
#   Fetch:
#     - Phone IP
#     - Phone MAC address using ARP
#     - Model if visible from web page
#     - Extension number if available from Fanvil web/account/line pages
#
# Output:
#   D:\fanvil\Fanvil_Segment_Report_yyyyMMdd_HHmmss.csv
#   D:\fanvil\Fanvil_Segment_Report_yyyyMMdd_HHmmss.xlsx if Excel installed
#
# Notes:
#   MAC address is fetched using ARP, so it works even if web parsing fails.
#   Extension number depends on Fanvil firmware/page access.
#   If extension shows N/A, send the generated Debug HTML for exact firmware mapping.
# ============================================================

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$script:OutputFolder = "D:\fanvil"
$script:DebugFolder  = "D:\fanvil\Debug"
$script:Results = New-Object System.Collections.ArrayList
$script:IsScanning = $false
$script:CancelScan = $false

New-Item -ItemType Directory -Path $script:OutputFolder -Force | Out-Null
New-Item -ItemType Directory -Path $script:DebugFolder -Force | Out-Null

# ---------------- BASIC HELPERS ----------------

function Add-Log {
    param([string]$Text)
    $time = Get-Date -Format "HH:mm:ss"
    $txtLog.AppendText("[$time] $Text`r`n")
    $txtLog.SelectionStart = $txtLog.Text.Length
    $txtLog.ScrollToCaret()
    [System.Windows.Forms.Application]::DoEvents()
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

function Get-LocalIPv4 {
    try {
        return Get-NetIPAddress -AddressFamily IPv4 |
            Where-Object {
                $_.IPAddress -notlike "127.*" -and
                $_.IPAddress -notlike "169.254.*" -and
                $_.IPAddress -match "^\d{1,3}(\.\d{1,3}){3}$"
            } |
            Select-Object -First 1 -ExpandProperty IPAddress
    }
    catch {
        try {
            return Get-WmiObject Win32_NetworkAdapterConfiguration |
                Where-Object { $_.IPEnabled -eq $true -and $_.IPAddress } |
                ForEach-Object { $_.IPAddress } |
                Where-Object {
                    $_ -notlike "127.*" -and
                    $_ -notlike "169.254.*" -and
                    $_ -match "^\d{1,3}(\.\d{1,3}){3}$"
                } |
                Select-Object -First 1
        }
        catch { return $null }
    }
}

function Get-IPBase {
    param([string]$IP)
    $p = $IP.Split(".")
    return "$($p[0]).$($p[1]).$($p[2])"
}

function Parse-SegmentInput {
    param([string]$InputText)

    $input = $InputText.Trim()

    if ($input -match "^(\d{1,3}\.){2}\d{1,3}$") {
        return [PSCustomObject]@{ IsValid=$true; Base=$input; Start=1; End=254; Message="" }
    }

    if ($input -match "^((\d{1,3}\.){3})(\d{1,3})\s*-\s*(\d{1,3})$") {
        $base = $matches[1].TrimEnd(".")
        $start = [int]$matches[3]
        $end = [int]$matches[4]

        if ($start -lt 1 -or $end -gt 254 -or $start -gt $end) {
            return [PSCustomObject]@{ IsValid=$false; Message="Invalid range. Use 1-254." }
        }

        return [PSCustomObject]@{ IsValid=$true; Base=$base; Start=$start; End=$end; Message="" }
    }

    if ($input -match "^((\d{1,3}\.){3})(\d{1,3})$") {
        $p = $input.Split(".")
        return [PSCustomObject]@{ IsValid=$true; Base="$($p[0]).$($p[1]).$($p[2])"; Start=1; End=254; Message="" }
    }

    return [PSCustomObject]@{ IsValid=$false; Message="Enter segment like 10.209.110 or range like 10.209.110.1-254" }
}

# ---------------- NETWORK / ARP ----------------

function Test-DeviceReachable {
    param([string]$IP)

    try {
        Test-Connection -ComputerName $IP -Count 1 -Quiet -ErrorAction SilentlyContinue | Out-Null
    }
    catch {}

    try {
        ping $IP -n 1 -w 500 | Out-Null
    }
    catch {}
}

function Get-MacFromArp {
    param([string]$IP)

    Test-DeviceReachable -IP $IP

    try {
        $arp = arp -a $IP
        $m = ($arp | Select-String "([0-9a-fA-F]{2}-){5}[0-9a-fA-F]{2}").Matches.Value | Select-Object -First 1

        if ($m) {
            return ($m.ToUpper() -replace "-", ":")
        }
    }
    catch {}

    return "N/A"
}

function Test-HttpAlive {
    param([string]$IP, [int]$Timeout)

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

# ---------------- FANVIL WEB READ ----------------

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
        "http://$IP/cgi-bin/ConfigManApp.com",
        "http://$IP/login",
        "http://$IP/",
        "http://$IP/index.htm",
        "http://$IP/index.html"
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
                $headers = @{
                    "User-Agent" = "Mozilla/5.0"
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
            }
            catch {}
        }
    }
}

function Get-WebPage {
    param(
        [string]$Url,
        $Session,
        [int]$Timeout
    )

    try {
        $headers = @{
            "User-Agent" = "Mozilla/5.0"
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
    }
    catch {}

    return ""
}

function Get-FanvilHtml {
    param(
        [string]$IP,
        $Session,
        [int]$Timeout
    )

    $pages = @(
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
        "http://$IP/cgi-bin/ConfigManApp.com?key=RegisterStatus",

        "http://$IP/cgi-bin/ConfigManApp.com?key=SIP1",
        "http://$IP/cgi-bin/ConfigManApp.com?key=LINE1",
        "http://$IP/cgi-bin/ConfigManApp.com?key=VOIP",
        "http://$IP/cgi-bin/ConfigManApp.com?key=Phone"
    )

    $all = ""

    foreach ($url in $pages) {
        $h = Get-WebPage -Url $url -Session $Session -Timeout $Timeout
        if ($h) {
            $all += "`r`n===== PAGE: $url =====`r`n"
            $all += $h
        }
    }

    return $all
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
    param([string]$Text, [string[]]$Labels)

    foreach ($label in $Labels) {
        $pattern = [regex]::Escape($label) + "\s*:?\s*([A-Za-z0-9\.\-_:\/@]+(?:\s+[A-Za-z0-9\.\-_:\/@]+)?)"
        $m = [regex]::Match($Text, $pattern, "IgnoreCase")
        if ($m.Success) {
            $v = Clean-Text $m.Groups[1].Value
            if ($v -and $v -ne "N/A") { return $v }
        }
    }

    return "N/A"
}

function Extract-PhoneInfo {
    param(
        [string]$IP,
        [string]$Html,
        [string]$ArpMac,
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
    elseif ($plain -match "Fanvil") {
        $model = Extract-AfterLabel $plain @("Model", "Phone Model", "Device Model")
        if ($model -eq "N/A") { $model = "Fanvil" }
    }

    $webMac = Extract-Regex $plain @(
        "Ethernet\s*MAC\s*:?\s*([0-9A-Fa-f]{2}[:-][0-9A-Fa-f]{2}[:-][0-9A-Fa-f]{2}[:-][0-9A-Fa-f]{2}[:-][0-9A-Fa-f]{2}[:-][0-9A-Fa-f]{2})",
        "MAC\s*Address\s*:?\s*([0-9A-Fa-f]{2}[:-][0-9A-Fa-f]{2}[:-][0-9A-Fa-f]{2}[:-][0-9A-Fa-f]{2}[:-][0-9A-Fa-f]{2}[:-][0-9A-Fa-f]{2})",
        "MAC\s*:?\s*([0-9A-Fa-f]{2}[:-][0-9A-Fa-f]{2}[:-][0-9A-Fa-f]{2}[:-][0-9A-Fa-f]{2}[:-][0-9A-Fa-f]{2}[:-][0-9A-Fa-f]{2})"
    )

    $finalMac = $ArpMac
    if ($finalMac -eq "N/A" -and $webMac -ne "N/A") {
        $finalMac = $webMac
    }

    $ethernetIP = Extract-Regex $plain @(
        "Ethernet\s*IP\s*:?\s*((?:\d{1,3}\.){3}\d{1,3})",
        "IP\s*Address\s*:?\s*((?:\d{1,3}\.){3}\d{1,3})",
        "IPv4\s*:?\s*((?:\d{1,3}\.){3}\d{1,3})"
    )

    if ($ethernetIP -eq "N/A") { $ethernetIP = $IP }

    $extension = Extract-Regex $plain @(
        "Extension\s*:?\s*([0-9]{2,10})",
        "Ext\s*:?\s*([0-9]{2,10})",
        "Phone\s*Number\s*:?\s*([0-9]{2,10})",
        "SIP\s*User\s*ID\s*:?\s*([0-9]{2,10})",
        "User\s*ID\s*:?\s*([0-9]{2,10})",
        "Register\s*Name\s*:?\s*([0-9]{2,10})",
        "Auth\s*User\s*:?\s*([0-9]{2,10})",
        "Account\s*1\s*:?\s*([0-9]{2,10})",
        "Display\s*Name\s*:?\s*([0-9]{2,10})",
        "UserName\s*:?\s*([0-9]{2,10})",
        "username\s*[=:]\s*['""]?([0-9]{2,10})",
        "sipUser\s*[=:]\s*['""]?([0-9]{2,10})",
        "account\s*[=:]\s*['""]?([0-9]{2,10})"
    )

    $registerStatus = Extract-Regex $plain @(
        "Register\s*Status\s*:?\s*([A-Za-z0-9\s]+)",
        "Registration\s*Status\s*:?\s*([A-Za-z0-9\s]+)",
        "SIP\s*Status\s*:?\s*([A-Za-z0-9\s]+)"
    )

    $networkMode = Extract-AfterLabel $plain @("Network mode", "Network Mode", "IP Mode")
    $software = Extract-AfterLabel $plain @("Software", "Firmware", "Firmware Version", "Version")

    $isFanvil = "No"
    if ($plain -match "Fanvil|X301G|X3SG|Ethernet MAC|Ethernet IP|Default gateway") {
        $isFanvil = "Yes"
    }

    return [PSCustomObject]@{
        ScanDate       = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        PhoneIP        = $IP
        PhoneMAC       = $finalMac
        ArpMAC         = $ArpMac
        WebMAC         = $webMac
        Extension      = $extension
        Model          = $model
        EthernetIP     = $ethernetIP
        NetworkMode    = $networkMode
        RegisterStatus = $registerStatus
        Software       = $software
        IsFanvil       = $isFanvil
        DebugFile      = $DebugFile
    }
}

# ---------------- EXPORT ----------------

function Export-Results {
    if ($script:Results.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("No results to export.", "Export", "OK", "Information") | Out-Null
        return
    }

    $stamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $csv = Join-Path $script:OutputFolder "Fanvil_Segment_Report_$stamp.csv"
    $xlsx = Join-Path $script:OutputFolder "Fanvil_Segment_Report_$stamp.xlsx"

    $script:Results | Export-Csv -Path $csv -NoTypeInformation -Encoding UTF8

    $excelOk = $false

    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false

        $wb = $excel.Workbooks.Open($csv)
        $ws = $wb.Worksheets.Item(1)
        $ws.Name = "Fanvil Phones"
        $ws.UsedRange.EntireColumn.AutoFit() | Out-Null
        $ws.Range("A1:M1").Font.Bold = $true
        $ws.Range("A1:M1").Interior.ColorIndex = 15

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
        [System.Windows.Forms.MessageBox]::Show("CSV saved:`r`n$csv`r`n`r`nExcel was not available, so XLSX was not created.", "Export Done", "OK", "Information") | Out-Null
    }
}

# ---------------- GUI ----------------

$form = New-Object System.Windows.Forms.Form
$form.Text = "Fanvil IP MAC Extension Scanner"
$form.Size = New-Object System.Drawing.Size(1150, 720)
$form.StartPosition = "CenterScreen"
$form.Font = New-Object System.Drawing.Font("Segoe UI", 9)

$lblTitle = New-Object System.Windows.Forms.Label
$lblTitle.Text = "Fanvil IP MAC Extension Scanner"
$lblTitle.Font = New-Object System.Drawing.Font("Segoe UI", 16, [System.Drawing.FontStyle]::Bold)
$lblTitle.AutoSize = $true
$lblTitle.Location = New-Object System.Drawing.Point(15, 15)
$form.Controls.Add($lblTitle)

$lblSub = New-Object System.Windows.Forms.Label
$lblSub.Text = "Enter one IP segment. It fetches Fanvil IP, MAC using ARP, and extension if visible from phone web pages. Output: D:\fanvil"
$lblSub.AutoSize = $true
$lblSub.Location = New-Object System.Drawing.Point(18, 48)
$form.Controls.Add($lblSub)

$lblSeg = New-Object System.Windows.Forms.Label
$lblSeg.Text = "IP Segment:"
$lblSeg.AutoSize = $true
$lblSeg.Location = New-Object System.Drawing.Point(20, 85)
$form.Controls.Add($lblSeg)

$txtSegment = New-Object System.Windows.Forms.TextBox
$txtSegment.Location = New-Object System.Drawing.Point(105, 81)
$txtSegment.Size = New-Object System.Drawing.Size(180, 25)
$form.Controls.Add($txtSegment)

$localIp = Get-LocalIPv4
if ($localIp) { $txtSegment.Text = Get-IPBase $localIp }

$lblHint = New-Object System.Windows.Forms.Label
$lblHint.Text = "Example: 10.209.110 or 10.209.110.1-254"
$lblHint.AutoSize = $true
$lblHint.Location = New-Object System.Drawing.Point(295, 85)
$form.Controls.Add($lblHint)

$lblUser = New-Object System.Windows.Forms.Label
$lblUser.Text = "User:"
$lblUser.AutoSize = $true
$lblUser.Location = New-Object System.Drawing.Point(20, 122)
$form.Controls.Add($lblUser)

$txtUser = New-Object System.Windows.Forms.TextBox
$txtUser.Location = New-Object System.Drawing.Point(105, 118)
$txtUser.Size = New-Object System.Drawing.Size(90, 25)
$txtUser.Text = "admin"
$form.Controls.Add($txtUser)

$lblPass = New-Object System.Windows.Forms.Label
$lblPass.Text = "Password:"
$lblPass.AutoSize = $true
$lblPass.Location = New-Object System.Drawing.Point(210, 122)
$form.Controls.Add($lblPass)

$txtPass = New-Object System.Windows.Forms.TextBox
$txtPass.Location = New-Object System.Drawing.Point(285, 118)
$txtPass.Size = New-Object System.Drawing.Size(90, 25)
$txtPass.Text = "admin"
$form.Controls.Add($txtPass)

$lblTimeout = New-Object System.Windows.Forms.Label
$lblTimeout.Text = "Timeout:"
$lblTimeout.AutoSize = $true
$lblTimeout.Location = New-Object System.Drawing.Point(395, 122)
$form.Controls.Add($lblTimeout)

$numTimeout = New-Object System.Windows.Forms.NumericUpDown
$numTimeout.Location = New-Object System.Drawing.Point(455, 118)
$numTimeout.Size = New-Object System.Drawing.Size(60, 25)
$numTimeout.Minimum = 1
$numTimeout.Maximum = 10
$numTimeout.Value = 2
$form.Controls.Add($numTimeout)

$chkOnlyFanvil = New-Object System.Windows.Forms.CheckBox
$chkOnlyFanvil.Text = "Show only Fanvil detected"
$chkOnlyFanvil.AutoSize = $true
$chkOnlyFanvil.Location = New-Object System.Drawing.Point(540, 120)
$chkOnlyFanvil.Checked = $false
$form.Controls.Add($chkOnlyFanvil)

$btnScan = New-Object System.Windows.Forms.Button
$btnScan.Text = "Scan"
$btnScan.Location = New-Object System.Drawing.Point(780, 78)
$btnScan.Size = New-Object System.Drawing.Size(90, 32)
$form.Controls.Add($btnScan)

$btnCancel = New-Object System.Windows.Forms.Button
$btnCancel.Text = "Cancel"
$btnCancel.Location = New-Object System.Drawing.Point(880, 78)
$btnCancel.Size = New-Object System.Drawing.Size(90, 32)
$btnCancel.Enabled = $false
$form.Controls.Add($btnCancel)

$btnExport = New-Object System.Windows.Forms.Button
$btnExport.Text = "Export"
$btnExport.Location = New-Object System.Drawing.Point(980, 78)
$btnExport.Size = New-Object System.Drawing.Size(90, 32)
$form.Controls.Add($btnExport)

$progress = New-Object System.Windows.Forms.ProgressBar
$progress.Location = New-Object System.Drawing.Point(20, 155)
$progress.Size = New-Object System.Drawing.Size(1050, 22)
$progress.Minimum = 0
$progress.Maximum = 100
$form.Controls.Add($progress)

$lblStatus = New-Object System.Windows.Forms.Label
$lblStatus.Text = "Ready"
$lblStatus.AutoSize = $true
$lblStatus.Location = New-Object System.Drawing.Point(20, 183)
$form.Controls.Add($lblStatus)

$grid = New-Object System.Windows.Forms.DataGridView
$grid.Location = New-Object System.Drawing.Point(20, 210)
$grid.Size = New-Object System.Drawing.Size(1050, 330)
$grid.AllowUserToAddRows = $false
$grid.ReadOnly = $true
$grid.SelectionMode = "FullRowSelect"
$grid.AutoSizeColumnsMode = "Fill"
$form.Controls.Add($grid)

$table = New-Object System.Data.DataTable
@("S.No","Phone IP","Phone MAC","Extension","Model","Ethernet IP","Register Status","Software","Is Fanvil") | ForEach-Object {
    [void]$table.Columns.Add($_)
}
$grid.DataSource = $table

$txtLog = New-Object System.Windows.Forms.TextBox
$txtLog.Location = New-Object System.Drawing.Point(20, 555)
$txtLog.Size = New-Object System.Drawing.Size(1050, 100)
$txtLog.Multiline = $true
$txtLog.ScrollBars = "Vertical"
$txtLog.ReadOnly = $true
$form.Controls.Add($txtLog)

# ---------------- EVENTS ----------------

$btnScan.Add_Click({
    if ($script:IsScanning) { return }

    $parsed = Parse-SegmentInput $txtSegment.Text

    if (!$parsed.IsValid) {
        [System.Windows.Forms.MessageBox]::Show($parsed.Message, "Invalid Segment", "OK", "Warning") | Out-Null
        return
    }

    $script:Results.Clear()
    $table.Rows.Clear()
    $txtLog.Clear()
    $progress.Value = 0

    $script:IsScanning = $true
    $script:CancelScan = $false

    $btnScan.Enabled = $false
    $btnCancel.Enabled = $true
    $btnExport.Enabled = $false

    $base = $parsed.Base
    $start = [int]$parsed.Start
    $end = [int]$parsed.End
    $total = ($end - $start) + 1

    $user = $txtUser.Text.Trim()
    $pass = $txtPass.Text.Trim()
    $timeout = [int]$numTimeout.Value

    $found = 0
    $displayNo = 0
    $count = 0

    Add-Log "Scan started: $base.$start-$end"
    Add-Log "MAC will be fetched from ARP."
    Add-Log "Extension will be fetched from Fanvil web/account pages if available."

    for ($i = $start; $i -le $end; $i++) {
        if ($script:CancelScan) {
            Add-Log "Scan cancelled by user."
            break
        }

        $ip = "$base.$i"
        $count++

        $percent = [int](($count / $total) * 100)
        if ($percent -gt 100) { $percent = 100 }

        $progress.Value = $percent
        $lblStatus.Text = "Checking $ip ($count/$total)"
        [System.Windows.Forms.Application]::DoEvents()

        # ARP MAC check
        $mac = Get-MacFromArp -IP $ip

        if ($mac -eq "N/A") {
            continue
        }

        # Check web page. We only need web for extension/model.
        $isHttp = Test-HttpAlive -IP $ip -Timeout $timeout

        if (!$isHttp) {
            continue
        }

        Add-Log "Web device found: $ip | MAC: $mac"

        $session = New-Object Microsoft.PowerShell.Commands.WebRequestSession
        Invoke-FanvilLogin -IP $ip -Session $session -User $user -Pass $pass -Timeout $timeout

        $html = Get-FanvilHtml -IP $ip -Session $session -Timeout $timeout

        $debugFile = Join-Path $script:DebugFolder ("Debug_" + $ip.Replace(".","_") + ".html")
        if ($html) {
            $html | Out-File $debugFile -Encoding UTF8
        }
        else {
            "" | Out-File $debugFile -Encoding UTF8
        }

        $info = Extract-PhoneInfo -IP $ip -Html $html -ArpMac $mac -DebugFile $debugFile

        if ($chkOnlyFanvil.Checked -and $info.IsFanvil -ne "Yes") {
            Add-Log "Skipped non-Fanvil web device: $ip"
            continue
        }

        $found++
        [void]$script:Results.Add($info)

        $displayNo++
        $row = $table.NewRow()
        $row["S.No"] = $displayNo
        $row["Phone IP"] = $info.PhoneIP
        $row["Phone MAC"] = $info.PhoneMAC
        $row["Extension"] = $info.Extension
        $row["Model"] = $info.Model
        $row["Ethernet IP"] = $info.EthernetIP
        $row["Register Status"] = $info.RegisterStatus
        $row["Software"] = $info.Software
        $row["Is Fanvil"] = $info.IsFanvil
        $table.Rows.Add($row)

        Add-Log "Fetched: IP=$ip | MAC=$($info.PhoneMAC) | Ext=$($info.Extension) | Model=$($info.Model)"
    }

    $progress.Value = 100
    $lblStatus.Text = "Completed. Records found: $found"

    $script:IsScanning = $false
    $btnScan.Enabled = $true
    $btnCancel.Enabled = $false
    $btnExport.Enabled = $true

    Add-Log "Scan completed. Records found: $found"

    if ($found -gt 0) {
        Export-Results
    }
})

$btnCancel.Add_Click({
    $script:CancelScan = $true
    $lblStatus.Text = "Cancelling..."
})

$btnExport.Add_Click({
    Export-Results
})

Add-Log "Ready."
Add-Log "Enter segment and click Scan."

[void]$form.ShowDialog()
