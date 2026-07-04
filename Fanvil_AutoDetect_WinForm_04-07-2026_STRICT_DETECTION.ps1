# Fanvil Auto Detect WinForms Scanner
# Updated: 04-07-2026 STRICT FANVIL DETECTION
# Purpose:
#   Enter IP segment, click Start Scan, detect web devices,
#   get IP related MAC using ping + arp -a,
#   mark Brand as Fanvil only when confirmed Fanvil text/model is detected,
#   export Excel/CSV to D:\fanvil.
#
# Output folder: D:\fanvil
# Requirement: Windows PowerShell 5.1, .NET WinForms. Excel is optional.

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$ErrorActionPreference = 'Continue'
$OUT_DIR = 'D:\fanvil'
if (!(Test-Path $OUT_DIR)) { New-Item -Path $OUT_DIR -ItemType Directory -Force | Out-Null }

$script:CancelScan = $false
$script:IsScanning = $false
$script:LastOutputFile = ''

function Write-AppLog {
    param([string]$Message)
    $time = Get-Date -Format 'HH:mm:ss'
    $script:txtLog.AppendText("[$time] $Message`r`n")
    $script:txtLog.SelectionStart = $script:txtLog.Text.Length
    $script:txtLog.ScrollToCaret()
    [System.Windows.Forms.Application]::DoEvents()
}

function Get-ArpMac {
    param([string]$Ip)

    try {
        ping $Ip -n 1 -w 700 | Out-Null
        Start-Sleep -Milliseconds 120

        $arp = arp -a $Ip 2>$null | Out-String
        $mac = $null

        foreach ($line in ($arp -split "`r?`n")) {
            if ($line -match [regex]::Escape($Ip) -and $line -match '([0-9a-fA-F]{2}[-:]){5}[0-9a-fA-F]{2}') {
                $mac = $Matches[0].ToUpper().Replace('-',':')
                break
            }
        }

        if ($mac) { return $mac }
        return ''
    }
    catch {
        return ''
    }
}

function Test-Port {
    param([string]$Ip, [int]$Port)

    try {
        $client = New-Object System.Net.Sockets.TcpClient
        $iar = $client.BeginConnect($Ip, $Port, $null, $null)
        $ok = $iar.AsyncWaitHandle.WaitOne(800, $false)

        if ($ok) {
            $client.EndConnect($iar)
            $client.Close()
            return $true
        }

        $client.Close()
        return $false
    }
    catch {
        return $false
    }
}

function Invoke-WebSafe {
    param(
        [string]$Url,
        [int]$TimeoutSec = 4
    )

    try {
        return Invoke-WebRequest -Uri $Url -TimeoutSec $TimeoutSec -UseBasicParsing -ErrorAction Stop
    }
    catch {
        return $null
    }
}

function Get-PageTitle {
    param([string]$Html)

    if (!$Html) { return '' }

    try {
        $m = [regex]::Match($Html, '<title[^>]*>(.*?)</title>', 'IgnoreCase,Singleline')
        if ($m.Success) {
            $title = $m.Groups[1].Value
            $title = $title -replace '<[^>]+>', ''
            $title = $title -replace '&nbsp;', ' '
            $title = $title -replace '&amp;', '&'
            $title = [regex]::Replace($title, '\s+', ' ').Trim()
            return $title
        }
    }
    catch {}

    return ''
}

function Detect-FanvilBasic {
    param([string]$Ip)

    $ports = @(80, 8080)

    foreach ($port in $ports) {
        if (Test-Port -Ip $Ip -Port $port) {

            $url = "http://$Ip"
            if ($port -ne 80) { $url = "http://$Ip`:$port" }

            $resp = Invoke-WebSafe -Url $url -TimeoutSec 4
            $html = if ($resp) { [string]$resp.Content } else { '' }
            $server = if ($resp -and $resp.Headers['Server']) { [string]$resp.Headers['Server'] } else { '' }
            $title = Get-PageTitle -Html $html

            $isFanvil = $false

            # Strict Fanvil detection:
            # Do NOT mark generic SIP/VoIP/GoAhead/Boa pages as Fanvil.
            # Some routers/printers/cameras also show generic web server names.
            # Brand becomes Fanvil only if real Fanvil text/model is found.
            $fanvilCheckText = ($html + " " + $title + " " + $server)

            if ($fanvilCheckText -match '(?i)\bFanvil\b') {
                $isFanvil = $true
            }
            elseif ($fanvilCheckText -match '(?i)\bX301G\b|\bX3SG\b|\bX301\b|\bX3S\b') {
                $isFanvil = $true
            }
            else {
                $isFanvil = $false
            }

            return [PSCustomObject]@{
                WebOpen  = $true
                Url      = $url
                IsFanvil = $isFanvil
                Html     = $html
                Server   = $server
                Title    = $title
                Status   = 'HTTP OK'
            }
        }
    }

    return [PSCustomObject]@{
        WebOpen  = $false
        Url      = ''
        IsFanvil = $false
        Html     = ''
        Server   = ''
        Title    = ''
        Status   = ''
    }
}

function Export-Results {
    param([array]$Rows)

    $stamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    $csvPath = Join-Path $OUT_DIR "Fanvil_Web_Devices_$stamp.csv"
    $xlsxPath = Join-Path $OUT_DIR "Fanvil_Web_Devices_$stamp.xlsx"

    $Rows | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8

    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false

        $wb = $excel.Workbooks.Open($csvPath)
        $ws = $wb.Worksheets.Item(1)
        $ws.Name = 'Fanvil Devices'
        $ws.UsedRange.EntireColumn.AutoFit() | Out-Null
        $ws.Range('A1:K1').Font.Bold = $true
        $ws.Range('A1:K1').Interior.ColorIndex = 15

        $wb.SaveAs($xlsxPath, 51)
        $wb.Close($true)
        $excel.Quit()

        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($ws) | Out-Null
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($wb) | Out-Null
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null

        return $xlsxPath
    }
    catch {
        return $csvPath
    }
}

function Start-Scan {
    if ($script:IsScanning) { return }

    $segment = $script:txtSegment.Text.Trim()
    $start = [int]$script:numStart.Value
    $end = [int]$script:numEnd.Value

    if ($segment -notmatch '^\d{1,3}\.\d{1,3}\.\d{1,3}$') {
        [System.Windows.Forms.MessageBox]::Show('Enter segment like 10.209.110','Invalid Segment','OK','Warning') | Out-Null
        return
    }

    if ($start -gt $end) {
        [System.Windows.Forms.MessageBox]::Show('Start IP must be less than End IP','Invalid Range','OK','Warning') | Out-Null
        return
    }

    $script:IsScanning = $true
    $script:CancelScan = $false
    $script:grid.Rows.Clear()
    $script:btnStart.Enabled = $false
    $script:btnStop.Enabled = $true
    $script:progress.Value = 0

    $total = ($end - $start + 1)
    $count = 0
    $results = New-Object System.Collections.ArrayList

    Write-AppLog "🚀 Scan started for $segment.$start to $segment.$end"
    Write-AppLog "📌 MAC method: ping IP -n 1, then arp -a IP"
    Write-AppLog "📁 Output folder: $OUT_DIR"

    for ($i=$start; $i -le $end; $i++) {

        if ($script:CancelScan) {
            Write-AppLog "⛔ Scan stopped by user."
            break
        }

        $ip = "$segment.$i"
        $count++
        $script:progress.Value = [Math]::Min(100, [int](($count / $total) * 100))
        [System.Windows.Forms.Application]::DoEvents()

        $mac = Get-ArpMac -Ip $ip
        $detect = Detect-FanvilBasic -Ip $ip

        if (-not $detect.WebOpen) {
            continue
        }

        $brand = ''
        $loginStatus = ''

        if ($detect.IsFanvil) {
            $brand = 'Fanvil'
            $loginStatus = 'Confirmed'
        }
        else {
            $brand = 'Web Device'
            $loginStatus = ''
        }

        # Requested: keep these values empty for now.
        $model = ''
        $extension = ''
        $sipUser = ''

        $row = [PSCustomObject]@{
            'S.No'             = $results.Count + 1
            'IP Address'       = $ip
            'MAC Address'      = $mac
            'Brand'            = $brand
            'Model'            = $model
            'Extension Number' = $extension
            'SIP Username'     = $sipUser
            'URL'              = $detect.Url
            'Login Status'     = $loginStatus
            'Web Status'       = $detect.Status
            'Scan Time'        = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
        }

        [void]$results.Add($row)

        [void]$script:grid.Rows.Add(
            $row.'S.No',
            $row.'IP Address',
            $row.'MAC Address',
            $row.Brand,
            $row.Model,
            $row.'Extension Number',
            $row.'SIP Username',
            $row.URL,
            $row.'Login Status',
            $row.'Web Status'
        )

        Write-AppLog "✅ Found: $ip | MAC=$mac | Brand=$brand | URL=$($detect.Url)"
    }

    if ($results.Count -gt 0) {
        $file = Export-Results -Rows $results
        $script:LastOutputFile = $file
        Write-AppLog "✅ Export completed: $file"
        [System.Windows.Forms.MessageBox]::Show("Scan completed.`nFound: $($results.Count)`nSaved: $file",'Completed','OK','Information') | Out-Null
    }
    else {
        Write-AppLog 'No web devices found.'
        [System.Windows.Forms.MessageBox]::Show('No web devices found.','Completed','OK','Information') | Out-Null
    }

    $script:IsScanning = $false
    $script:btnStart.Enabled = $true
    $script:btnStop.Enabled = $false
    $script:progress.Value = 100
}

function Stop-Scan {
    $script:CancelScan = $true
    Write-AppLog "Stop requested. Please wait..."
}

function Open-OutputFolder {
    if (!(Test-Path $OUT_DIR)) {
        New-Item -Path $OUT_DIR -ItemType Directory -Force | Out-Null
    }
    Invoke-Item $OUT_DIR
}

# UI
$form = New-Object System.Windows.Forms.Form
$form.Text = '📞 Fanvil IP MAC Scanner'
$form.Size = New-Object System.Drawing.Size(1060,660)
$form.StartPosition = 'CenterScreen'
$form.MinimumSize = New-Object System.Drawing.Size(980,600)
$form.Font = New-Object System.Drawing.Font('Segoe UI',9)
$form.ShowIcon = $false

$title = New-Object System.Windows.Forms.Label
$title.Text = '📞 Fanvil IP MAC Scanner'
$title.Font = New-Object System.Drawing.Font('Segoe UI',16,[System.Drawing.FontStyle]::Bold)
$title.Location = New-Object System.Drawing.Point(18,15)
$title.Size = New-Object System.Drawing.Size(650,34)
$form.Controls.Add($title)

$subTitle = New-Object System.Windows.Forms.Label
$subTitle.Text = 'Scan web devices, fetch IP related MAC address, and export Excel to D:\fanvil'
$subTitle.Location = New-Object System.Drawing.Point(22,45)
$subTitle.Size = New-Object System.Drawing.Size(780,22)
$form.Controls.Add($subTitle)

$lblSegment = New-Object System.Windows.Forms.Label
$lblSegment.Text = 'IP Segment'
$lblSegment.Location = New-Object System.Drawing.Point(20,82)
$lblSegment.Size = New-Object System.Drawing.Size(90,25)
$form.Controls.Add($lblSegment)

$txtSegment = New-Object System.Windows.Forms.TextBox
$txtSegment.Text = '10.209.110'
$txtSegment.Location = New-Object System.Drawing.Point(110,79)
$txtSegment.Size = New-Object System.Drawing.Size(140,25)
$form.Controls.Add($txtSegment)
$script:txtSegment = $txtSegment

$lblRange = New-Object System.Windows.Forms.Label
$lblRange.Text = 'Range'
$lblRange.Location = New-Object System.Drawing.Point(270,82)
$lblRange.Size = New-Object System.Drawing.Size(50,25)
$form.Controls.Add($lblRange)

$numStart = New-Object System.Windows.Forms.NumericUpDown
$numStart.Minimum = 1
$numStart.Maximum = 254
$numStart.Value = 1
$numStart.Location = New-Object System.Drawing.Point(320,79)
$numStart.Size = New-Object System.Drawing.Size(60,25)
$form.Controls.Add($numStart)
$script:numStart = $numStart

$numEnd = New-Object System.Windows.Forms.NumericUpDown
$numEnd.Minimum = 1
$numEnd.Maximum = 254
$numEnd.Value = 254
$numEnd.Location = New-Object System.Drawing.Point(390,79)
$numEnd.Size = New-Object System.Drawing.Size(60,25)
$form.Controls.Add($numEnd)
$script:numEnd = $numEnd

$btnStart = New-Object System.Windows.Forms.Button
$btnStart.Text = '▶ Start Scan'
$btnStart.Location = New-Object System.Drawing.Point(480,74)
$btnStart.Size = New-Object System.Drawing.Size(120,34)
$btnStart.Add_Click({ Start-Scan })
$form.Controls.Add($btnStart)
$script:btnStart = $btnStart

$btnStop = New-Object System.Windows.Forms.Button
$btnStop.Text = '⏹ Stop Scan'
$btnStop.Location = New-Object System.Drawing.Point(615,74)
$btnStop.Size = New-Object System.Drawing.Size(120,34)
$btnStop.Enabled = $false
$btnStop.Add_Click({ Stop-Scan })
$form.Controls.Add($btnStop)
$script:btnStop = $btnStop

$btnOpenOutput = New-Object System.Windows.Forms.Button
$btnOpenOutput.Text = '📂 Open Output'
$btnOpenOutput.Location = New-Object System.Drawing.Point(750,74)
$btnOpenOutput.Size = New-Object System.Drawing.Size(140,34)
$btnOpenOutput.Add_Click({ Open-OutputFolder })
$form.Controls.Add($btnOpenOutput)
$script:btnOpenOutput = $btnOpenOutput

$progress = New-Object System.Windows.Forms.ProgressBar
$progress.Location = New-Object System.Drawing.Point(20,122)
$progress.Size = New-Object System.Drawing.Size(1000,18)
$form.Controls.Add($progress)
$script:progress = $progress

$grid = New-Object System.Windows.Forms.DataGridView
$grid.Location = New-Object System.Drawing.Point(20,155)
$grid.Size = New-Object System.Drawing.Size(1000,335)
$grid.Anchor = 'Top,Left,Right,Bottom'
$grid.AllowUserToAddRows = $false
$grid.ReadOnly = $true
$grid.AutoSizeColumnsMode = 'Fill'

$cols = @(
    'S.No',
    'IP Address',
    'MAC Address',
    'Brand',
    'Model',
    'Extension Number',
    'SIP Username',
    'URL',
    'Login Status',
    'Web Status'
)

foreach ($c in $cols) {
    [void]$grid.Columns.Add($c,$c)
}

$form.Controls.Add($grid)
$script:grid = $grid

$txtLog = New-Object System.Windows.Forms.TextBox
$txtLog.Location = New-Object System.Drawing.Point(20,505)
$txtLog.Size = New-Object System.Drawing.Size(1000,85)
$txtLog.Anchor = 'Left,Right,Bottom'
$txtLog.Multiline = $true
$txtLog.ScrollBars = 'Vertical'
$txtLog.ReadOnly = $true
$form.Controls.Add($txtLog)
$script:txtLog = $txtLog

$note = New-Object System.Windows.Forms.Label
$note.Text = '📁 Output: D:\fanvil  |  🧾 Excel/CSV columns include IP Address, MAC Address, Brand, URL, Login Status and Scan Time.'
$note.Location = New-Object System.Drawing.Point(20,602)
$note.Size = New-Object System.Drawing.Size(980,25)
$note.Anchor = 'Left,Right,Bottom'
$form.Controls.Add($note)

Write-AppLog 'Ready. Enter segment and click Start Scan.'
[void]$form.ShowDialog()
