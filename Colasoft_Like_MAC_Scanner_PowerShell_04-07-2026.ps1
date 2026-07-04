# Colasoft-like MAC Scanner - PowerShell WinForms
# Works on Windows PowerShell 5.1
# Purpose: Scan LAN IP range, ping devices, read ARP table, show IP/MAC/Hostname/Vendor, export CSV/Excel CSV.

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Hide PowerShell console window
try {
    Add-Type @"
using System;
using System.Runtime.InteropServices;
public class Win32Window {
    [DllImport("kernel32.dll")] public static extern IntPtr GetConsoleWindow();
    [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
}
"@
    $consolePtr = [Win32Window]::GetConsoleWindow()
    if ($consolePtr -ne [IntPtr]::Zero) { [Win32Window]::ShowWindow($consolePtr, 0) | Out-Null }
} catch {}

$script:StopScan = $false
$script:Results = New-Object System.Collections.Generic.List[object]

# Common vendor OUI map. Add more OUIs whenever needed.
$VendorMap = @{
    '00:A8:59' = 'Fanvil'
    '0C:38:3E' = 'Fanvil'
    '00:09:6E' = 'Fanvil'
    '00:1A:A0' = 'Dell'
    'F8:B1:56' = 'Dell'
    '3C:2C:30' = 'Dell'
    'B8:27:EB' = 'Raspberry Pi'
    'DC:A6:32' = 'Raspberry Pi'
    '00:50:56' = 'VMware'
    '00:0C:29' = 'VMware'
    '08:00:27' = 'VirtualBox'
    '00:1B:21' = 'Intel'
    '3C:52:82' = 'Intel'
    '10:7B:44' = 'HP'
    'D8:9D:67' = 'HP'
    'FC:15:B4' = 'Hikvision'
    'BC:AD:28' = 'Hangzhou Hikvision'
    '00:18:E7' = 'Cameo / Network Device'
}

function Normalize-Mac {
    param([string]$Mac)
    if ([string]::IsNullOrWhiteSpace($Mac)) { return '' }
    $m = $Mac.Trim().ToUpper() -replace '-', ':'
    if ($m -match '^[0-9A-F]{12}$') {
        return (($m -split '(.{2})' | Where-Object { $_ }) -join ':')
    }
    return $m
}

function Get-Oui {
    param([string]$Mac)
    $m = Normalize-Mac $Mac
    if ($m.Length -ge 8) { return $m.Substring(0,8) }
    return ''
}

function Get-VendorName {
    param([string]$Mac)
    $oui = Get-Oui $Mac
    if ($VendorMap.ContainsKey($oui)) { return $VendorMap[$oui] }
    if ($oui) { return 'Unknown Vendor' }
    return ''
}

function Get-ArpMacForIp {
    param([string]$Ip)
    try {
        $arp = arp -a $Ip 2>$null | Out-String
        foreach ($line in ($arp -split "`r?`n")) {
            if ($line -match "\b$([regex]::Escape($Ip))\b\s+([0-9A-Fa-f:-]{17})\s+") {
                return Normalize-Mac $matches[1]
            }
        }
    } catch {}
    return ''
}

function Resolve-HostNameSafe {
    param([string]$Ip)
    try {
        $entry = [System.Net.Dns]::GetHostEntry($Ip)
        if ($entry.HostName) { return $entry.HostName }
    } catch {}
    return ''
}

function Test-PortQuick {
    param([string]$Ip, [int]$Port, [int]$TimeoutMs = 250)
    try {
        $client = New-Object System.Net.Sockets.TcpClient
        $iar = $client.BeginConnect($Ip, $Port, $null, $null)
        $ok = $iar.AsyncWaitHandle.WaitOne($TimeoutMs, $false)
        if ($ok -and $client.Connected) {
            $client.EndConnect($iar)
            $client.Close()
            return $true
        }
        $client.Close()
    } catch {}
    return $false
}

function Ping-Ip {
    param([string]$Ip)
    try {
        return Test-Connection -ComputerName $Ip -Count 1 -Quiet -ErrorAction SilentlyContinue
    } catch {
        try {
            $p = New-Object System.Net.NetworkInformation.Ping
            $r = $p.Send($Ip, 500)
            return ($r.Status -eq 'Success')
        } catch { return $false }
    }
}

function Add-Log {
    param([string]$Text)
    $time = Get-Date -Format 'HH:mm:ss'
    $txtLog.AppendText("[$time] $Text`r`n")
    $txtLog.SelectionStart = $txtLog.Text.Length
    $txtLog.ScrollToCaret()
    [System.Windows.Forms.Application]::DoEvents()
}

function Add-ResultRow {
    param($Obj)
    $script:Results.Add($Obj) | Out-Null
    $rowIndex = $grid.Rows.Add()
    $row = $grid.Rows[$rowIndex]
    $row.Cells['SNo'].Value = $Obj.SNo
    $row.Cells['IPAddress'].Value = $Obj.IPAddress
    $row.Cells['MACAddress'].Value = $Obj.MACAddress
    $row.Cells['Vendor'].Value = $Obj.Vendor
    $row.Cells['HostName'].Value = $Obj.HostName
    $row.Cells['Status'].Value = $Obj.Status
    $row.Cells['HTTP'].Value = $Obj.HTTP
    $row.Cells['HTTPS'].Value = $Obj.HTTPS
    $row.Cells['URL'].Value = $Obj.URL

    if ($Obj.Status -eq 'Online') { $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(225,255,225) }
    elseif ($Obj.Status -eq 'ARP Only') { $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(230,240,255) }
    else { $row.DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(255,235,235) }
}

# ---------------- GUI ----------------
$form = New-Object System.Windows.Forms.Form
$form.Text = 'LAN MAC Scanner'
$form.Size = New-Object System.Drawing.Size(1220, 760)
$form.StartPosition = 'CenterScreen'
$form.Font = New-Object System.Drawing.Font('Segoe UI', 9)
$form.BackColor = [System.Drawing.Color]::WhiteSmoke

$lblTitle = New-Object System.Windows.Forms.Label
$lblTitle.Text = 'LAN MAC Address Scanner'
$lblTitle.Font = New-Object System.Drawing.Font('Segoe UI', 20, [System.Drawing.FontStyle]::Bold)
$lblTitle.AutoSize = $true
$lblTitle.Location = New-Object System.Drawing.Point(20, 18)
$form.Controls.Add($lblTitle)

$lblSub = New-Object System.Windows.Forms.Label
$lblSub.Text = 'Scan IP range, detect MAC address, hostname, vendor and web ports'
$lblSub.Font = New-Object System.Drawing.Font('Segoe UI', 10)
$lblSub.AutoSize = $true
$lblSub.Location = New-Object System.Drawing.Point(24, 58)
$form.Controls.Add($lblSub)

$panel = New-Object System.Windows.Forms.Panel
$panel.Location = New-Object System.Drawing.Point(20, 90)
$panel.Size = New-Object System.Drawing.Size(1160, 82)
$panel.BorderStyle = 'FixedSingle'
$panel.BackColor = [System.Drawing.Color]::White
$form.Controls.Add($panel)

$lblSegment = New-Object System.Windows.Forms.Label
$lblSegment.Text = 'IP Segment'
$lblSegment.Location = New-Object System.Drawing.Point(15, 15)
$lblSegment.AutoSize = $true
$panel.Controls.Add($lblSegment)

$txtSegment = New-Object System.Windows.Forms.TextBox
$txtSegment.Text = '10.209.110'
$txtSegment.Location = New-Object System.Drawing.Point(15, 38)
$txtSegment.Size = New-Object System.Drawing.Size(180, 25)
$panel.Controls.Add($txtSegment)

$lblFrom = New-Object System.Windows.Forms.Label
$lblFrom.Text = 'From'
$lblFrom.Location = New-Object System.Drawing.Point(225, 15)
$lblFrom.AutoSize = $true
$panel.Controls.Add($lblFrom)

$numFrom = New-Object System.Windows.Forms.NumericUpDown
$numFrom.Minimum = 1
$numFrom.Maximum = 254
$numFrom.Value = 1
$numFrom.Location = New-Object System.Drawing.Point(225, 38)
$numFrom.Size = New-Object System.Drawing.Size(80, 25)
$panel.Controls.Add($numFrom)

$lblTo = New-Object System.Windows.Forms.Label
$lblTo.Text = 'To'
$lblTo.Location = New-Object System.Drawing.Point(330, 15)
$lblTo.AutoSize = $true
$panel.Controls.Add($lblTo)

$numTo = New-Object System.Windows.Forms.NumericUpDown
$numTo.Minimum = 1
$numTo.Maximum = 254
$numTo.Value = 254
$numTo.Location = New-Object System.Drawing.Point(330, 38)
$numTo.Size = New-Object System.Drawing.Size(80, 25)
$panel.Controls.Add($numTo)

$chkHostname = New-Object System.Windows.Forms.CheckBox
$chkHostname.Text = 'Resolve Hostname'
$chkHostname.Checked = $true
$chkHostname.Location = New-Object System.Drawing.Point(445, 38)
$chkHostname.Size = New-Object System.Drawing.Size(145, 25)
$panel.Controls.Add($chkHostname)

$chkPorts = New-Object System.Windows.Forms.CheckBox
$chkPorts.Text = 'Check HTTP/HTTPS'
$chkPorts.Checked = $true
$chkPorts.Location = New-Object System.Drawing.Point(600, 38)
$chkPorts.Size = New-Object System.Drawing.Size(160, 25)
$panel.Controls.Add($chkPorts)

$btnStart = New-Object System.Windows.Forms.Button
$btnStart.Text = 'Start Scan'
$btnStart.Location = New-Object System.Drawing.Point(795, 25)
$btnStart.Size = New-Object System.Drawing.Size(110, 38)
$btnStart.BackColor = [System.Drawing.Color]::FromArgb(42, 160, 80)
$btnStart.ForeColor = [System.Drawing.Color]::White
$btnStart.FlatStyle = 'Flat'
$panel.Controls.Add($btnStart)

$btnStop = New-Object System.Windows.Forms.Button
$btnStop.Text = 'Stop'
$btnStop.Enabled = $false
$btnStop.Location = New-Object System.Drawing.Point(915, 25)
$btnStop.Size = New-Object System.Drawing.Size(90, 38)
$btnStop.BackColor = [System.Drawing.Color]::FromArgb(220, 70, 70)
$btnStop.ForeColor = [System.Drawing.Color]::White
$btnStop.FlatStyle = 'Flat'
$panel.Controls.Add($btnStop)

$btnExport = New-Object System.Windows.Forms.Button
$btnExport.Text = 'Export CSV'
$btnExport.Location = New-Object System.Drawing.Point(1015, 25)
$btnExport.Size = New-Object System.Drawing.Size(120, 38)
$btnExport.BackColor = [System.Drawing.Color]::FromArgb(40, 105, 190)
$btnExport.ForeColor = [System.Drawing.Color]::White
$btnExport.FlatStyle = 'Flat'
$panel.Controls.Add($btnExport)

$progress = New-Object System.Windows.Forms.ProgressBar
$progress.Location = New-Object System.Drawing.Point(20, 185)
$progress.Size = New-Object System.Drawing.Size(1160, 20)
$form.Controls.Add($progress)

$lblStatus = New-Object System.Windows.Forms.Label
$lblStatus.Text = 'Ready'
$lblStatus.Location = New-Object System.Drawing.Point(20, 210)
$lblStatus.Size = New-Object System.Drawing.Size(1160, 25)
$form.Controls.Add($lblStatus)

$grid = New-Object System.Windows.Forms.DataGridView
$grid.Location = New-Object System.Drawing.Point(20, 240)
$grid.Size = New-Object System.Drawing.Size(1160, 360)
$grid.AllowUserToAddRows = $false
$grid.AllowUserToDeleteRows = $false
$grid.ReadOnly = $true
$grid.SelectionMode = 'FullRowSelect'
$grid.AutoSizeColumnsMode = 'Fill'
$grid.RowHeadersVisible = $false
$form.Controls.Add($grid)

$cols = @(
    @{Name='SNo'; Header='S.No'; Width=50},
    @{Name='IPAddress'; Header='IP Address'; Width=120},
    @{Name='MACAddress'; Header='MAC Address'; Width=150},
    @{Name='Vendor'; Header='Vendor'; Width=150},
    @{Name='HostName'; Header='Host Name'; Width=190},
    @{Name='Status'; Header='Status'; Width=95},
    @{Name='HTTP'; Header='HTTP'; Width=70},
    @{Name='HTTPS'; Header='HTTPS'; Width=70},
    @{Name='URL'; Header='URL'; Width=190}
)
foreach ($c in $cols) {
    $col = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $col.Name = $c.Name
    $col.HeaderText = $c.Header
    $col.FillWeight = $c.Width
    [void]$grid.Columns.Add($col)
}

$txtLog = New-Object System.Windows.Forms.TextBox
$txtLog.Location = New-Object System.Drawing.Point(20, 615)
$txtLog.Size = New-Object System.Drawing.Size(1160, 90)
$txtLog.Multiline = $true
$txtLog.ScrollBars = 'Vertical'
$txtLog.ReadOnly = $true
$txtLog.BackColor = [System.Drawing.Color]::White
$form.Controls.Add($txtLog)

$btnStop.Add_Click({
    $script:StopScan = $true
    Add-Log 'Stop requested by user.'
})

$btnExport.Add_Click({
    if ($script:Results.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show('No data to export.', 'Export', 'OK', 'Information') | Out-Null
        return
    }
    $outDir = Join-Path $env:USERPROFILE 'Desktop\LAN_MAC_Scanner_Output'
    if (!(Test-Path $outDir)) { New-Item -ItemType Directory -Path $outDir -Force | Out-Null }
    $file = Join-Path $outDir ('LAN_MAC_Scan_' + (Get-Date -Format 'dd-MM-yyyy_HH-mm-ss') + '.csv')
    $script:Results | Export-Csv -Path $file -NoTypeInformation -Encoding UTF8
    Add-Log "Exported: $file"
    Start-Process explorer.exe "/select,`"$file`""
})

$btnStart.Add_Click({
    $script:StopScan = $false
    $script:Results.Clear()
    $grid.Rows.Clear()
    $txtLog.Clear()

    $segment = $txtSegment.Text.Trim().TrimEnd('.')
    $from = [int]$numFrom.Value
    $to = [int]$numTo.Value
    if ($from -gt $to) {
        [System.Windows.Forms.MessageBox]::Show('From value cannot be greater than To value.', 'Invalid Range', 'OK', 'Warning') | Out-Null
        return
    }
    if ($segment -notmatch '^\d{1,3}\.\d{1,3}\.\d{1,3}$') {
        [System.Windows.Forms.MessageBox]::Show('Enter IP segment like 10.209.110', 'Invalid Segment', 'OK', 'Warning') | Out-Null
        return
    }

    $btnStart.Enabled = $false
    $btnStop.Enabled = $true
    $progress.Minimum = 0
    $progress.Maximum = ($to - $from + 1)
    $progress.Value = 0

    Add-Log "Scan started: $segment.$from to $segment.$to"
    $sno = 0
    $total = $to - $from + 1
    $done = 0

    for ($i = $from; $i -le $to; $i++) {
        if ($script:StopScan) { break }
        $ip = "$segment.$i"
        $done++
        $lblStatus.Text = "Scanning $ip   ($done / $total)"
        [System.Windows.Forms.Application]::DoEvents()

        $isOnline = Ping-Ip $ip
        $mac = Get-ArpMacForIp $ip

        # If ping failed, still try ARP after a short ping attempt. This catches devices blocking ping but present in ARP.
        if (-not $mac) {
            try { ping -n 1 -w 250 $ip | Out-Null } catch {}
            Start-Sleep -Milliseconds 30
            $mac = Get-ArpMacForIp $ip
        }

        $status = 'Offline'
        if ($isOnline) { $status = 'Online' }
        elseif ($mac) { $status = 'ARP Only' }

        $http = ''
        $https = ''
        $url = ''
        if ($chkPorts.Checked -and ($isOnline -or $mac)) {
            $httpOpen = Test-PortQuick $ip 80 250
            $httpsOpen = Test-PortQuick $ip 443 250
            $http = if ($httpOpen) { 'Open' } else { 'Closed' }
            $https = if ($httpsOpen) { 'Open' } else { 'Closed' }
            if ($httpOpen) { $url = "http://$ip/" }
            elseif ($httpsOpen) { $url = "https://$ip/" }
        }

        $host = ''
        if ($chkHostname.Checked -and ($isOnline -or $mac)) { $host = Resolve-HostNameSafe $ip }

        if ($isOnline -or $mac) {
            $sno++
            $vendor = Get-VendorName $mac
            $obj = [pscustomobject]@{
                SNo        = $sno
                IPAddress  = $ip
                MACAddress = $mac
                Vendor     = $vendor
                HostName   = $host
                Status     = $status
                HTTP       = $http
                HTTPS      = $https
                URL        = $url
                ScanTime   = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
            }
            Add-ResultRow $obj
            Add-Log "Found $ip  MAC=$mac  Vendor=$vendor  Status=$status"
        }

        if ($progress.Value -lt $progress.Maximum) { $progress.Value++ }
        [System.Windows.Forms.Application]::DoEvents()
    }

    $lblStatus.Text = "Completed. Found $($script:Results.Count) device(s)."
    Add-Log "Scan completed. Found $($script:Results.Count) device(s)."
    $btnStart.Enabled = $true
    $btnStop.Enabled = $false
})

[void]$form.ShowDialog()
