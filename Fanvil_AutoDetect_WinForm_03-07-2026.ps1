
# Fanvil Auto Detect WinForms Scanner
# Created: 03-07-2026
# Purpose: Scan IP segment, detect Fanvil web phones, fetch IP/MAC and try to fetch extension details after login.
# Output folder: D:\fanvil
# Requirement: Windows PowerShell 5.1, .NET WinForms. Excel is optional.

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$ErrorActionPreference = 'Continue'
$OUT_DIR = 'D:\fanvil'
if (!(Test-Path $OUT_DIR)) { New-Item -Path $OUT_DIR -ItemType Directory -Force | Out-Null }

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
        Start-Sleep -Milliseconds 80
        $arp = arp -a $Ip 2>$null | Out-String
        $mac = $null
        foreach ($line in ($arp -split "`r?`n")) {
            if ($line -match [regex]::Escape($Ip) -and $line -match '([0-9a-fA-F]{2}[-:]){5}[0-9a-fA-F]{2}') {
                $mac = $Matches[0].ToUpper().Replace('-',':')
                break
            }
        }
        return $mac
    } catch { return $null }
}

function Test-Port {
    param([string]$Ip, [int]$Port)
    try {
        $client = New-Object System.Net.Sockets.TcpClient
        $iar = $client.BeginConnect($Ip, $Port, $null, $null)
        $ok = $iar.AsyncWaitHandle.WaitOne(800, $false)
        if ($ok) { $client.EndConnect($iar); $client.Close(); return $true }
        $client.Close(); return $false
    } catch { return $false }
}

function Invoke-WebSafe {
    param(
        [string]$Url,
        [Microsoft.PowerShell.Commands.WebRequestSession]$Session = $null,
        [string]$Method = 'GET',
        [object]$Body = $null,
        [int]$TimeoutSec = 5
    )
    try {
        $params = @{ Uri=$Url; Method=$Method; TimeoutSec=$TimeoutSec; UseBasicParsing=$true }
        if ($Session) { $params.WebSession = $Session }
        if ($Body) { $params.Body = $Body }
        return Invoke-WebRequest @params
    } catch {
        return $null
    }
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
            $isFanvil = $false
            if ($html -match '(?i)fanvil|X301|X3S|X4|X5|X6|SIP Phone|voip') { $isFanvil = $true }
            if ($server -match '(?i)fanvil|GoAhead|Boa') { $isFanvil = $true }
            return [PSCustomObject]@{ WebOpen=$true; Url=$url; IsFanvil=$isFanvil; Html=$html; Server=$server }
        }
    }
    return [PSCustomObject]@{ WebOpen=$false; Url=''; IsFanvil=$false; Html=''; Server='' }
}

function Extract-DetailsFromHtml {
    param([string]$Html)
    $model = ''
    $ext = ''
    $sipUser = ''

    if ($Html -match '(?i)(X301G|X3SG|X3S|X4U|X5U|X6U|X7|V62|V63|V64|V65|H3|H5|H2U)') { $model = $Matches[1] }

    $patterns = @(
        '(?i)(?:Extension|Ext|Account|Display\s*Name|SIP\s*User|User\s*Name|Auth\s*Name)[^0-9A-Za-z]{0,30}([0-9]{2,8})',
        '(?i)name=["'']?(?:account|Account|SIP[\w]*|phone|Phone|username|user)["'']?[^>]{0,120}value=["'']?([0-9]{2,8})',
        '(?i)value=["'']?([0-9]{2,8})["'']?[^>]{0,120}(?:Extension|SIP|Account|User)'
    )
    foreach ($p in $patterns) {
        if ($Html -match $p) { $ext = $Matches[1]; $sipUser = $Matches[1]; break }
    }
    return [PSCustomObject]@{ Model=$model; Extension=$ext; SipUser=$sipUser }
}

function Try-FanvilLoginAndFetch {
    param([string]$BaseUrl, [string]$Username, [string]$Password)

    $session = New-Object Microsoft.PowerShell.Commands.WebRequestSession
    $loginOk = $false
    $allHtml = ''

    # Fanvil firmware differs by model/version. Try several common login endpoints safely.
    $loginBodies = @(
        @{ username=$Username; password=$Password },
        @{ Username=$Username; Password=$Password },
        @{ user=$Username; pwd=$Password },
        @{ UserName=$Username; Password=$Password }
    )
    $loginPaths = @('/login.cgi','/cgi-bin/login.cgi','/dologin.htm','/login','/index.htm','/')

    foreach ($path in $loginPaths) {
        foreach ($body in $loginBodies) {
            $url = $BaseUrl.TrimEnd('/') + $path
            $r = Invoke-WebSafe -Url $url -Session $session -Method 'POST' -Body $body -TimeoutSec 5
            if ($r) {
                $content = [string]$r.Content
                if ($content -notmatch '(?i)invalid|incorrect|login failed|password error') {
                    $loginOk = $true
                    $allHtml += "`n" + $content
                    break
                }
            }
        }
        if ($loginOk) { break }
    }

    # Read likely status/account pages. Even if login detection is uncertain, session cookies may work.
    $pages = @(
        '/', '/index.htm', '/status.htm', '/Status.htm', '/cgi-bin/status.cgi',
        '/phone_status.htm', '/system.htm', '/config.htm', '/sip.htm', '/SIP.htm',
        '/account.htm', '/Account.htm', '/line.htm', '/Line.htm',
        '/cgi-bin/ConfigManApp.com?key=SIP',
        '/cgi-bin/ConfigManApp.com?key=LINE'
    )
    foreach ($p in $pages) {
        $u = $BaseUrl.TrimEnd('/') + $p
        $resp = Invoke-WebSafe -Url $u -Session $session -TimeoutSec 5
        if ($resp) { $allHtml += "`n" + ([string]$resp.Content) }
    }

    $details = Extract-DetailsFromHtml -Html $allHtml
    return [PSCustomObject]@{
        LoginStatus = if ($loginOk) { 'Tried / Session Created' } else { 'Not Confirmed' }
        Model = $details.Model
        Extension = $details.Extension
        SipUser = $details.SipUser
    }
}

function Export-Results {
    param([array]$Rows)
    $stamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    $csvPath = Join-Path $OUT_DIR "Fanvil_Phones_$stamp.csv"
    $xlsxPath = Join-Path $OUT_DIR "Fanvil_Phones_$stamp.xlsx"

    $Rows | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8

    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $wb = $excel.Workbooks.Open($csvPath)
        $ws = $wb.Worksheets.Item(1)
        $ws.UsedRange.EntireColumn.AutoFit() | Out-Null
        $wb.SaveAs($xlsxPath, 51)
        $wb.Close($true)
        $excel.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($ws) | Out-Null
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($wb) | Out-Null
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
        return $xlsxPath
    } catch {
        return $csvPath
    }
}

function Start-Scan {
    $segment = $script:txtSegment.Text.Trim()
    $start = [int]$script:numStart.Value
    $end = [int]$script:numEnd.Value
    $user = $script:txtUser.Text.Trim()
    $pass = $script:txtPass.Text

    if ($segment -notmatch '^\d{1,3}\.\d{1,3}\.\d{1,3}$') {
        [System.Windows.Forms.MessageBox]::Show('Enter segment like 10.209.110','Invalid Segment','OK','Warning') | Out-Null
        return
    }
    if ($start -gt $end) {
        [System.Windows.Forms.MessageBox]::Show('Start IP must be less than End IP','Invalid Range','OK','Warning') | Out-Null
        return
    }

    $script:grid.Rows.Clear()
    $script:btnScan.Enabled = $false
    $script:progress.Value = 0
    $total = ($end - $start + 1)
    $count = 0
    $results = New-Object System.Collections.ArrayList

    Write-AppLog "Scan started for $segment.$start to $segment.$end"

    for ($i=$start; $i -le $end; $i++) {
        $ip = "$segment.$i"
        $count++
        $script:progress.Value = [Math]::Min(100, [int](($count / $total) * 100))
        [System.Windows.Forms.Application]::DoEvents()

        $alive = Test-Connection -ComputerName $ip -Count 1 -Quiet -ErrorAction SilentlyContinue
        if (-not $alive) { continue }

        $mac = Get-ArpMac -Ip $ip
        $detect = Detect-FanvilBasic -Ip $ip
        if (-not $detect.WebOpen) { continue }

        $brand = if ($detect.IsFanvil) { 'Fanvil / Possible Fanvil' } else { 'Web Device' }
        $model = ''
        $extension = ''
        $sipUser = ''
        $loginStatus = 'Not Tried'

        if ($detect.IsFanvil -and $user -and $pass) {
            Write-AppLog "Fanvil possible: $ip - trying login/read"
            $fetch = Try-FanvilLoginAndFetch -BaseUrl $detect.Url -Username $user -Password $pass
            $model = $fetch.Model
            $extension = $fetch.Extension
            $sipUser = $fetch.SipUser
            $loginStatus = $fetch.LoginStatus
        } else {
            $d = Extract-DetailsFromHtml -Html $detect.Html
            $model = $d.Model
            $extension = $d.Extension
            $sipUser = $d.SipUser
        }

        $row = [PSCustomObject]@{
            'S.No' = $results.Count + 1
            'IP Address' = $ip
            'MAC Address' = $mac
            'Brand' = $brand
            'Model' = $model
            'Extension Number' = $extension
            'SIP Username' = $sipUser
            'Web URL' = $detect.Url
            'Login Status' = $loginStatus
            'Scan Time' = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
        }
        [void]$results.Add($row)
        [void]$script:grid.Rows.Add($row.'S.No',$row.'IP Address',$row.'MAC Address',$row.Brand,$row.Model,$row.'Extension Number',$row.'SIP Username',$row.'Login Status')
        Write-AppLog "Found: $ip MAC=$mac Brand=$brand Ext=$extension"
    }

    if ($results.Count -gt 0) {
        $file = Export-Results -Rows $results
        Write-AppLog "Export completed: $file"
        [System.Windows.Forms.MessageBox]::Show("Scan completed.`nFound: $($results.Count)`nSaved: $file",'Completed','OK','Information') | Out-Null
    } else {
        Write-AppLog 'No web/Fanvil devices found.'
        [System.Windows.Forms.MessageBox]::Show('No web/Fanvil devices found.','Completed','OK','Information') | Out-Null
    }
    $script:btnScan.Enabled = $true
}

# UI
$form = New-Object System.Windows.Forms.Form
$form.Text = 'Fanvil Auto Detect Scanner'
$form.Size = New-Object System.Drawing.Size(980,650)
$form.StartPosition = 'CenterScreen'
$form.MinimumSize = New-Object System.Drawing.Size(900,580)
$form.Font = New-Object System.Drawing.Font('Segoe UI',9)

$title = New-Object System.Windows.Forms.Label
$title.Text = 'Fanvil LAN Phone Auto Detect Scanner'
$title.Font = New-Object System.Drawing.Font('Segoe UI',16,[System.Drawing.FontStyle]::Bold)
$title.Location = New-Object System.Drawing.Point(18,15)
$title.Size = New-Object System.Drawing.Size(650,34)
$form.Controls.Add($title)

$lblSegment = New-Object System.Windows.Forms.Label
$lblSegment.Text = 'IP Segment'
$lblSegment.Location = New-Object System.Drawing.Point(20,65)
$lblSegment.Size = New-Object System.Drawing.Size(90,25)
$form.Controls.Add($lblSegment)

$txtSegment = New-Object System.Windows.Forms.TextBox
$txtSegment.Text = '10.209.110'
$txtSegment.Location = New-Object System.Drawing.Point(110,62)
$txtSegment.Size = New-Object System.Drawing.Size(140,25)
$form.Controls.Add($txtSegment)
$script:txtSegment = $txtSegment

$lblRange = New-Object System.Windows.Forms.Label
$lblRange.Text = 'Range'
$lblRange.Location = New-Object System.Drawing.Point(270,65)
$lblRange.Size = New-Object System.Drawing.Size(50,25)
$form.Controls.Add($lblRange)

$numStart = New-Object System.Windows.Forms.NumericUpDown
$numStart.Minimum = 1; $numStart.Maximum = 254; $numStart.Value = 1
$numStart.Location = New-Object System.Drawing.Point(320,62)
$numStart.Size = New-Object System.Drawing.Size(60,25)
$form.Controls.Add($numStart)
$script:numStart = $numStart

$numEnd = New-Object System.Windows.Forms.NumericUpDown
$numEnd.Minimum = 1; $numEnd.Maximum = 254; $numEnd.Value = 254
$numEnd.Location = New-Object System.Drawing.Point(390,62)
$numEnd.Size = New-Object System.Drawing.Size(60,25)
$form.Controls.Add($numEnd)
$script:numEnd = $numEnd

$lblUser = New-Object System.Windows.Forms.Label
$lblUser.Text = 'Username'
$lblUser.Location = New-Object System.Drawing.Point(470,65)
$lblUser.Size = New-Object System.Drawing.Size(75,25)
$form.Controls.Add($lblUser)

$txtUser = New-Object System.Windows.Forms.TextBox
$txtUser.Text = 'admin'
$txtUser.Location = New-Object System.Drawing.Point(545,62)
$txtUser.Size = New-Object System.Drawing.Size(90,25)
$form.Controls.Add($txtUser)
$script:txtUser = $txtUser

$lblPass = New-Object System.Windows.Forms.Label
$lblPass.Text = 'Password'
$lblPass.Location = New-Object System.Drawing.Point(650,65)
$lblPass.Size = New-Object System.Drawing.Size(75,25)
$form.Controls.Add($lblPass)

$txtPass = New-Object System.Windows.Forms.TextBox
$txtPass.Text = 'admin'
$txtPass.UseSystemPasswordChar = $true
$txtPass.Location = New-Object System.Drawing.Point(725,62)
$txtPass.Size = New-Object System.Drawing.Size(95,25)
$form.Controls.Add($txtPass)
$script:txtPass = $txtPass

$btnScan = New-Object System.Windows.Forms.Button
$btnScan.Text = 'SCAN & EXPORT'
$btnScan.Location = New-Object System.Drawing.Point(835,59)
$btnScan.Size = New-Object System.Drawing.Size(120,32)
$btnScan.Add_Click({ Start-Scan })
$form.Controls.Add($btnScan)
$script:btnScan = $btnScan

$progress = New-Object System.Windows.Forms.ProgressBar
$progress.Location = New-Object System.Drawing.Point(20,105)
$progress.Size = New-Object System.Drawing.Size(935,18)
$form.Controls.Add($progress)
$script:progress = $progress

$grid = New-Object System.Windows.Forms.DataGridView
$grid.Location = New-Object System.Drawing.Point(20,140)
$grid.Size = New-Object System.Drawing.Size(935,330)
$grid.Anchor = 'Top,Left,Right,Bottom'
$grid.AllowUserToAddRows = $false
$grid.ReadOnly = $true
$grid.AutoSizeColumnsMode = 'Fill'
$cols = @('S.No','IP Address','MAC Address','Brand','Model','Extension Number','SIP Username','Login Status')
foreach ($c in $cols) { [void]$grid.Columns.Add($c,$c) }
$form.Controls.Add($grid)
$script:grid = $grid

$txtLog = New-Object System.Windows.Forms.TextBox
$txtLog.Location = New-Object System.Drawing.Point(20,485)
$txtLog.Size = New-Object System.Drawing.Size(935,85)
$txtLog.Anchor = 'Left,Right,Bottom'
$txtLog.Multiline = $true
$txtLog.ScrollBars = 'Vertical'
$txtLog.ReadOnly = $true
$form.Controls.Add($txtLog)
$script:txtLog = $txtLog

$note = New-Object System.Windows.Forms.Label
$note.Text = 'Output: D:\fanvil | IP+MAC works without login. Extension fetch depends on Fanvil firmware web page/API and correct credential.'
$note.Location = New-Object System.Drawing.Point(20,580)
$note.Size = New-Object System.Drawing.Size(900,25)
$note.Anchor = 'Left,Right,Bottom'
$form.Controls.Add($note)

Write-AppLog 'Ready. Enter segment and click SCAN & EXPORT.'
[void]$form.ShowDialog()
