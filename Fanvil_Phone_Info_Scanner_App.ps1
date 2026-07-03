# ============================================================
# FANVIL PHONE INFO SCANNER APPLICATION
# Models  : Fanvil X301G / X3SG / Fanvil web phones
# Purpose : Enter IP segment, click Scan, login admin/admin,
#           fetch Model, Ethernet MAC, IP, Gateway, Extension etc.
# Output  : D:\fanvil\Fanvil_Phone_Info_Report_yyyyMMdd_HHmmss.csv
#           D:\fanvil\Fanvil_Phone_Info_Report_yyyyMMdd_HHmmss.xlsx if Excel installed
# ============================================================

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$script:OutputFolder = "D:\fanvil"
$script:Results = New-Object System.Collections.ArrayList
$script:IsScanning = $false
$script:CancelScan = $false

if (!(Test-Path $script:OutputFolder)) {
    New-Item -ItemType Directory -Path $script:OutputFolder -Force | Out-Null
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
    } catch { return $null }
}

function Get-IPBase {
    param([string]$IP)
    $p = $IP.Split(".")
    return "$($p[0]).$($p[1]).$($p[2])"
}

function Get-PCMac {
    try {
        $mac = Get-NetAdapter | Where-Object { $_.Status -eq "Up" } | Select-Object -First 1 -ExpandProperty MacAddress
        if ($mac) { return $mac }
    } catch {}
    return "N/A"
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

    return [PSCustomObject]@{ IsValid=$false; Message="Enter like 10.209.110 or 10.209.110.1-254" }
}

function Test-HttpAlive {
    param([string]$IP, [int]$Timeout)
    try {
        Invoke-WebRequest -Uri "http://$IP" -UseBasicParsing -TimeoutSec $Timeout -ErrorAction Stop | Out-Null
        return $true
    } catch { return $false }
}

function Invoke-FanvilLogin {
    param([string]$IP, $Session, [string]$User, [string]$Pass, [int]$Timeout)

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
                Invoke-WebRequest -Uri $url -Method POST -Body $body -WebSession $Session -UseBasicParsing -TimeoutSec $Timeout -ErrorAction Stop | Out-Null
            } catch {}
        }
    }
}

function Get-FanvilHtml {
    param([string]$IP, $Session, [int]$Timeout)

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
        "/cgi-bin/ConfigManApp.com?key=Status",
        "/cgi-bin/ConfigManApp.com?key=Network",
        "/cgi-bin/ConfigManApp.com?key=PhoneStatus",
        "/cgi-bin/ConfigManApp.com?key=NetworkInfo",
        "/cgi-bin/ConfigManApp.com?key=Line",
        "/cgi-bin/ConfigManApp.com?key=SIP",
        "/cgi-bin/ConfigManApp.com?key=Account"
    )

    $all = ""
    foreach ($p in $pages) {
        try {
            $url = "http://$IP$p"
            $r = Invoke-WebRequest -Uri $url -WebSession $Session -UseBasicParsing -TimeoutSec $Timeout -ErrorAction Stop
            if ($r.Content) {
                $all += "`r`n===== PAGE: $url =====`r`n"
                $all += [string]$r.Content
            }
        } catch {}
    }
    return $all
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

function Get-PhoneInfo {
    param([string]$IP, [string]$Html, [string]$DebugFile, [string]$PCIP, [string]$PCMAC)

    $plain = Clean-Text $Html

    $model = "N/A"
    if ($plain -match "X301G") { $model = "X301G" }
    elseif ($plain -match "X3SG") { $model = "X3SG" }
    else { $model = Extract-AfterLabel $plain @("Model") }

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

    $networkMode = Extract-AfterLabel $plain @("Network mode", "Network Mode", "IP Mode")
    $hardware = Extract-AfterLabel $plain @("Hardware")
    $software = Extract-AfterLabel $plain @("Software", "Firmware Version", "Firmware", "Version")
    $uboot = Extract-AfterLabel $plain @("Uboot")
    $uptime = Extract-Regex $plain @("Uptime\s*:?\s*([0-9]+\s*:\s*[0-9]+\s*:\s*[0-9]+)")

    $extension = Extract-Regex $plain @(
        "Extension\s*:?\s*([0-9]{2,10})",
        "Account\s*1\s*:?\s*([0-9]{2,10})",
        "SIP\s*User\s*ID\s*:?\s*([0-9]{2,10})",
        "User\s*ID\s*:?\s*([0-9]{2,10})",
        "Register\s*Name\s*:?\s*([0-9]{2,10})",
        "Phone\s*Number\s*:?\s*([0-9]{2,10})"
    )

    $lineStatus = Extract-Regex $plain @(
        "Register\s*Status\s*:?\s*([A-Za-z0-9\s]+)",
        "Registration\s*Status\s*:?\s*([A-Za-z0-9\s]+)",
        "SIP\s*Status\s*:?\s*([A-Za-z0-9\s]+)"
    )

    $isFanvil = "No"
    if ($plain -match "Fanvil|X301G|X3SG|Ethernet MAC|Ethernet IP") { $isFanvil = "Yes" }

    return [PSCustomObject]@{
        ScanDate       = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        PCName         = $env:COMPUTERNAME
        PCIP           = $PCIP
        PCMAC          = $PCMAC
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

function Export-Results {
    if ($script:Results.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("No results to export.", "Export", "OK", "Information") | Out-Null
        return
    }

    if (!(Test-Path $script:OutputFolder)) {
        New-Item -ItemType Directory -Path $script:OutputFolder -Force | Out-Null
    }

    $stamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $csv = Join-Path $script:OutputFolder "Fanvil_Phone_Info_Report_$stamp.csv"
    $xlsx = Join-Path $script:OutputFolder "Fanvil_Phone_Info_Report_$stamp.xlsx"

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
        $ws.Range("A1:S1").Font.Bold = $true
        $ws.Range("A1:S1").Interior.ColorIndex = 15
        $wb.SaveAs($xlsx, 51)
        $wb.Close($true)
        $excel.Quit()
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($ws) | Out-Null
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($wb) | Out-Null
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
        $excelOk = $true
    } catch { $excelOk = $false }

    if ($excelOk) {
        Add-Log "Excel saved: $xlsx"
        [System.Windows.Forms.MessageBox]::Show("Excel saved:`r`n$xlsx`r`n`r`nCSV also saved:`r`n$csv", "Export Done", "OK", "Information") | Out-Null
    } else {
        Add-Log "CSV saved: $csv"
        [System.Windows.Forms.MessageBox]::Show("CSV saved:`r`n$csv", "Export Done", "OK", "Information") | Out-Null
    }
}

$form = New-Object System.Windows.Forms.Form
$form.Text = "Fanvil Phone Info Scanner"
$form.Size = New-Object System.Drawing.Size(1180, 720)
$form.StartPosition = "CenterScreen"
$form.Font = New-Object System.Drawing.Font("Segoe UI", 9)

$lblTitle = New-Object System.Windows.Forms.Label
$lblTitle.Text = "Fanvil Phone Info Scanner"
$lblTitle.Font = New-Object System.Drawing.Font("Segoe UI", 16, [System.Drawing.FontStyle]::Bold)
$lblTitle.AutoSize = $true
$lblTitle.Location = New-Object System.Drawing.Point(15, 15)
$form.Controls.Add($lblTitle)

$lblSub = New-Object System.Windows.Forms.Label
$lblSub.Text = "Scans segment, logs in using admin/admin, fetches Model, Ethernet MAC, IP, Gateway, Extension, and exports to D:\fanvil"
$lblSub.AutoSize = $true
$lblSub.Location = New-Object System.Drawing.Point(18, 48)
$form.Controls.Add($lblSub)

$lblSegment = New-Object System.Windows.Forms.Label
$lblSegment.Text = "IP Segment / Range:"
$lblSegment.AutoSize = $true
$lblSegment.Location = New-Object System.Drawing.Point(18, 85)
$form.Controls.Add($lblSegment)

$txtSegment = New-Object System.Windows.Forms.TextBox
$txtSegment.Location = New-Object System.Drawing.Point(140, 81)
$txtSegment.Size = New-Object System.Drawing.Size(230, 25)
$form.Controls.Add($txtSegment)

$localIp = Get-LocalIPv4
if ($localIp) { $txtSegment.Text = Get-IPBase $localIp }

$lblHint = New-Object System.Windows.Forms.Label
$lblHint.Text = "Example: 10.209.110 or 10.209.110.1-254"
$lblHint.AutoSize = $true
$lblHint.Location = New-Object System.Drawing.Point(380, 85)
$form.Controls.Add($lblHint)

$lblCred = New-Object System.Windows.Forms.Label
$lblCred.Text = "Login:"
$lblCred.AutoSize = $true
$lblCred.Location = New-Object System.Drawing.Point(18, 122)
$form.Controls.Add($lblCred)

$txtUser = New-Object System.Windows.Forms.TextBox
$txtUser.Location = New-Object System.Drawing.Point(140, 118)
$txtUser.Size = New-Object System.Drawing.Size(100, 25)
$txtUser.Text = "admin"
$form.Controls.Add($txtUser)

$txtPass = New-Object System.Windows.Forms.TextBox
$txtPass.Location = New-Object System.Drawing.Point(250, 118)
$txtPass.Size = New-Object System.Drawing.Size(100, 25)
$txtPass.Text = "admin"
$form.Controls.Add($txtPass)

$lblTimeout = New-Object System.Windows.Forms.Label
$lblTimeout.Text = "Timeout:"
$lblTimeout.AutoSize = $true
$lblTimeout.Location = New-Object System.Drawing.Point(370, 122)
$form.Controls.Add($lblTimeout)

$numTimeout = New-Object System.Windows.Forms.NumericUpDown
$numTimeout.Location = New-Object System.Drawing.Point(430, 118)
$numTimeout.Size = New-Object System.Drawing.Size(60, 25)
$numTimeout.Minimum = 1
$numTimeout.Maximum = 10
$numTimeout.Value = 2
$form.Controls.Add($numTimeout)

$btnScan = New-Object System.Windows.Forms.Button
$btnScan.Text = "Scan"
$btnScan.Location = New-Object System.Drawing.Point(795, 78)
$btnScan.Size = New-Object System.Drawing.Size(100, 34)
$form.Controls.Add($btnScan)

$btnCancel = New-Object System.Windows.Forms.Button
$btnCancel.Text = "Cancel"
$btnCancel.Location = New-Object System.Drawing.Point(905, 78)
$btnCancel.Size = New-Object System.Drawing.Size(100, 34)
$btnCancel.Enabled = $false
$form.Controls.Add($btnCancel)

$btnExport = New-Object System.Windows.Forms.Button
$btnExport.Text = "Export"
$btnExport.Location = New-Object System.Drawing.Point(1015, 78)
$btnExport.Size = New-Object System.Drawing.Size(100, 34)
$form.Controls.Add($btnExport)

$progress = New-Object System.Windows.Forms.ProgressBar
$progress.Location = New-Object System.Drawing.Point(20, 155)
$progress.Size = New-Object System.Drawing.Size(1095, 22)
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
$grid.Size = New-Object System.Drawing.Size(1095, 335)
$grid.AllowUserToAddRows = $false
$grid.ReadOnly = $true
$grid.SelectionMode = "FullRowSelect"
$grid.AutoSizeColumnsMode = "Fill"
$form.Controls.Add($grid)

$table = New-Object System.Data.DataTable
@("S.No","Phone IP","Model","Extension","Ethernet MAC","Ethernet IP","Network Mode","Subnet Mask","Gateway","Software","Line Status","Is Fanvil") | ForEach-Object {
    [void]$table.Columns.Add($_)
}
$grid.DataSource = $table

$txtLog = New-Object System.Windows.Forms.TextBox
$txtLog.Location = New-Object System.Drawing.Point(20, 560)
$txtLog.Size = New-Object System.Drawing.Size(1095, 95)
$txtLog.Multiline = $true
$txtLog.ScrollBars = "Vertical"
$txtLog.ReadOnly = $true
$form.Controls.Add($txtLog)

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

    $user = $txtUser.Text.Trim()
    $pass = $txtPass.Text.Trim()
    $timeout = [int]$numTimeout.Value

    $base = $parsed.Base
    $start = [int]$parsed.Start
    $end = [int]$parsed.End
    $total = ($end - $start) + 1
    $count = 0
    $found = 0

    $pcIp = Get-LocalIPv4
    $pcMac = Get-PCMac

    Add-Log "Scan started: $base.$start-$end"
    Add-Log "Output path: $script:OutputFolder"

    for ($i = $start; $i -le $end; $i++) {
        if ($script:CancelScan) {
            Add-Log "Scan cancelled."
            break
        }

        $ip = "$base.$i"
        $count++
        $percent = [int](($count / $total) * 100)
        if ($percent -gt 100) { $percent = 100 }
        $progress.Value = $percent
        $lblStatus.Text = "Scanning $ip ($count/$total)"
        [System.Windows.Forms.Application]::DoEvents()

        if (!(Test-HttpAlive -IP $ip -Timeout $timeout)) {
            continue
        }

        Add-Log "Web device found: $ip - trying login and info page"

        $session = New-Object Microsoft.PowerShell.Commands.WebRequestSession
        Invoke-FanvilLogin -IP $ip -Session $session -User $user -Pass $pass -Timeout $timeout

        $html = Get-FanvilHtml -IP $ip -Session $session -Timeout $timeout
        if (!$html) { continue }

        $debugFile = Join-Path $script:OutputFolder ("Debug_" + $ip.Replace(".","_") + ".html")
        $html | Out-File $debugFile -Encoding UTF8

        $info = Get-PhoneInfo -IP $ip -Html $html -DebugFile $debugFile -PCIP $pcIp -PCMAC $pcMac

        if ($info.IsFanvil -ne "Yes") {
            Add-Log "Skipped non-Fanvil web device: $ip"
            continue
        }

        $found++
        [void]$script:Results.Add($info)

        $row = $table.NewRow()
        $row["S.No"] = $found
        $row["Phone IP"] = $info.PhoneIP
        $row["Model"] = $info.Model
        $row["Extension"] = $info.Extension
        $row["Ethernet MAC"] = $info.EthernetMAC
        $row["Ethernet IP"] = $info.EthernetIP
        $row["Network Mode"] = $info.NetworkMode
        $row["Subnet Mask"] = $info.SubnetMask
        $row["Gateway"] = $info.DefaultGateway
        $row["Software"] = $info.Software
        $row["Line Status"] = $info.LineStatus
        $row["Is Fanvil"] = $info.IsFanvil
        $table.Rows.Add($row)

        Add-Log "Fetched: $ip | $($info.Model) | MAC $($info.EthernetMAC) | Ext $($info.Extension)"
    }

    $lblStatus.Text = "Completed. Fanvil phones found: $found"
    $progress.Value = 100

    $script:IsScanning = $false
    $btnScan.Enabled = $true
    $btnCancel.Enabled = $false
    $btnExport.Enabled = $true

    Add-Log "Scan completed. Fanvil phones found: $found"

    if ($found -gt 0) { Export-Results }
})

$btnCancel.Add_Click({
    $script:CancelScan = $true
    $lblStatus.Text = "Cancelling..."
})

$btnExport.Add_Click({ Export-Results })

Add-Log "Application ready."
Add-Log "Enter segment and click Scan."

[void]$form.ShowDialog()
