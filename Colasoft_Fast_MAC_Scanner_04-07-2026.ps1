# Colasoft-Style Fast MAC Scanner - PowerShell WinForms
# Fast ARP/Ping scanner with hostname, workgroup/domain, manufacturer OUI lookup, compare result, CSV export
# Tested target: Windows PowerShell 5.1

# Hide PowerShell console
try {
    Add-Type @"
using System;
using System.Runtime.InteropServices;
public class Win32ShowWindow {
    [DllImport("kernel32.dll")] public static extern IntPtr GetConsoleWindow();
    [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
}
"@ -ErrorAction SilentlyContinue
    $consolePtr = [Win32ShowWindow]::GetConsoleWindow()
    if ($consolePtr -ne [IntPtr]::Zero) { [Win32ShowWindow]::ShowWindow($consolePtr, 0) | Out-Null }
} catch {}

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

[System.Windows.Forms.Application]::EnableVisualStyles()

$Script:AppDir = Join-Path $env:ProgramData "FastMACScanner"
$Script:HistoryPath = Join-Path $Script:AppDir "last_scan.csv"
$Script:OuiCsvPath = Join-Path $Script:AppDir "oui.csv"
$Script:ScanStop = $false
$Script:OuiMap = @{}
$Script:LastResults = @()
$Script:PreviousMap = @{}

if (!(Test-Path $Script:AppDir)) { New-Item -ItemType Directory -Path $Script:AppDir -Force | Out-Null }

function Load-BuiltinOui {
    $Script:OuiMap.Clear()
    $builtin = @{
        '00A859'='Fanvil Technology Co., Ltd.'
        '0C383E'='Fanvil Technology Co., Ltd.'
        '00096E'='Fanvil Technology Co., Ltd.'
        '001A79'='Cisco Systems, Inc.'
        '001B54'='Cisco Systems, Inc.'
        '005056'='VMware, Inc.'
        '000C29'='VMware, Inc.'
        '080027'='PCS Systemtechnik GmbH / VirtualBox'
        'F0D5BF'='Intel Corporate'
        '3C970E'='Wistron InfoComm'
        'D8BBC1'='Micro-Star INTL CO., LTD.'
        'B827EB'='Raspberry Pi Foundation'
        'D850E6'='ASUSTek COMPUTER INC.'
        'F8BC12'='Dell Inc.'
        'B8CA3A'='Dell Inc.'
        'A4BADB'='Dell Inc.'
        '6C2B59'='Hewlett Packard'
        '3CD92B'='Hewlett Packard'
        '10E7C6'='Hewlett Packard'
        '28D244'='LCFC(HeFei) Electronics Technology co., ltd'
        'E8B1FC'='Intel Corporate'
        '00155D'='Microsoft Corporation'
        '7C1E52'='Microsoft Corporation'
        'D067E5'='Dell Inc.'
        '001F29'='Hewlett Packard'
        '2C27D7'='Hewlett Packard'
        '002481'='Hewlett Packard'
        '001E0B'='Hewlett Packard'
        '001F3A'='Hon Hai Precision Ind. Co.,Ltd.'
        '18DBF2'='Dell Inc.'
        '8C1645'='Liteon Technology Corporation'
        '70B3D5'='IEEE Registration Authority'
        '0016B6'='Cisco-Linksys, LLC'
        '0024E8'='Dell Inc.'
        '34E6D7'='Dell Inc.'
        'D06726'='Dell Inc.'
        'F44D30'='Elitegroup Computer Systems Co.,Ltd.'
        'C85B76'='LCFC(HeFei) Electronics Technology co., ltd'
        'A0D3C1'='Hewlett Packard'
        '5CBAEF'='Chongqing Fugui Electronics Co.,Ltd.'
        '704D7B'='ASUSTek COMPUTER INC.'
        '001E8C'='ASUSTek COMPUTER INC.'
    }
    foreach ($k in $builtin.Keys) { $Script:OuiMap[$k] = $builtin[$k] }
}

function Import-OuiDatabase {
    Load-BuiltinOui
    if (Test-Path $Script:OuiCsvPath) {
        try {
            Import-Csv $Script:OuiCsvPath | ForEach-Object {
                $assignment = ($_.Assignment -replace '[^A-Fa-f0-9]', '').ToUpper()
                $org = $_.'Organization Name'
                if ($assignment.Length -ge 6 -and $org) { $Script:OuiMap[$assignment.Substring(0,6)] = $org }
            }
        } catch {}
    }
}

function Download-OuiDatabase {
    try {
        $url = 'https://standards-oui.ieee.org/oui/oui.csv'
        Invoke-WebRequest -Uri $url -OutFile $Script:OuiCsvPath -UseBasicParsing -TimeoutSec 30
        Import-OuiDatabase
        return $true
    } catch {
        return $false
    }
}

function Get-Manufacturer {
    param([string]$Mac)
    if ([string]::IsNullOrWhiteSpace($Mac)) { return '' }
    $clean = ($Mac -replace '[^A-Fa-f0-9]', '').ToUpper()
    if ($clean.Length -lt 6) { return 'Unknown' }
    $oui = $clean.Substring(0,6)
    if ($Script:OuiMap.ContainsKey($oui)) { return $Script:OuiMap[$oui] }
    return 'Unknown'
}

function Normalize-Mac {
    param([string]$Mac)
    if ([string]::IsNullOrWhiteSpace($Mac)) { return '' }
    $clean = ($Mac -replace '[^A-Fa-f0-9]', '').ToUpper()
    if ($clean.Length -lt 12) { return $Mac.ToUpper() }
    return (($clean.Substring(0,2),$clean.Substring(2,2),$clean.Substring(4,2),$clean.Substring(6,2),$clean.Substring(8,2),$clean.Substring(10,2)) -join ':')
}

function Get-LocalSubnets {
    $list = New-Object System.Collections.ArrayList
    try {
        $configs = Get-WmiObject Win32_NetworkAdapterConfiguration -Filter "IPEnabled=True" -ErrorAction Stop
        foreach ($c in $configs) {
            for ($i=0; $i -lt $c.IPAddress.Count; $i++) {
                $ip = $c.IPAddress[$i]
                $mask = $c.IPSubnet[$i]
                if ($ip -match '^\d+\.\d+\.\d+\.\d+$' -and $mask -match '^\d+\.\d+\.\d+\.\d+$') {
                    $parts = $ip.Split('.')
                    [void]$list.Add("$($parts[0]).$($parts[1]).$($parts[2]).1-254")
                }
            }
        }
    } catch {}
    if ($list.Count -eq 0) { [void]$list.Add('192.168.1.1-254') }
    return ($list | Select-Object -Unique)
}

function Expand-IPRange {
    param([string]$RangeText)
    $ips = New-Object System.Collections.ArrayList
    $text = $RangeText.Trim()
    if ($text -match '^(\d+\.\d+\.\d+\.)(\d+)\-(\d+)$') {
        $prefix = $matches[1]; $start=[int]$matches[2]; $end=[int]$matches[3]
        if ($start -lt 0) { $start=0 }; if ($end -gt 255) { $end=255 }
        for ($i=$start; $i -le $end; $i++) { [void]$ips.Add("$prefix$i") }
    } elseif ($text -match '^(\d+\.\d+\.\d+\.\d+)$') {
        [void]$ips.Add($text)
    } elseif ($text -match '^(\d+\.\d+\.\d+)\.\*$') {
        for ($i=1; $i -le 254; $i++) { [void]$ips.Add("$($matches[1]).$i") }
    }
    return $ips
}

function Get-ArpTable {
    $map = @{}
    try {
        $arp = arp -a 2>$null
        foreach ($line in $arp) {
            if ($line -match '(\d+\.\d+\.\d+\.\d+)\s+([a-fA-F0-9]{2}[-:][a-fA-F0-9]{2}[-:][a-fA-F0-9]{2}[-:][a-fA-F0-9]{2}[-:][a-fA-F0-9]{2}[-:][a-fA-F0-9]{2})') {
                $map[$matches[1]] = Normalize-Mac $matches[2]
            }
        }
    } catch {}
    return $map
}

function Resolve-HostNameFast {
    param([string]$IP)
    try {
        $entry = [System.Net.Dns]::GetHostEntry($IP)
        if ($entry.HostName) { return $entry.HostName }
    } catch {}
    try {
        $nbt = nbtstat -A $IP 2>$null
        foreach ($line in $nbt) {
            if ($line -match '^\s*([^\s<]+)\s+<00>\s+UNIQUE') { return $matches[1].Trim() }
        }
    } catch {}
    return ''
}

function Resolve-WorkgroupFast {
    param([string]$IP)
    try {
        $nbt = nbtstat -A $IP 2>$null
        foreach ($line in $nbt) {
            if ($line -match '^\s*([^\s<]+)\s+<00>\s+GROUP') { return $matches[1].Trim() }
        }
    } catch {}
    return ''
}

function Load-PreviousScan {
    $Script:PreviousMap.Clear()
    if (Test-Path $Script:HistoryPath) {
        try {
            Import-Csv $Script:HistoryPath | ForEach-Object {
                if ($_.MACAddress) { $Script:PreviousMap[$_.MACAddress] = $_ }
            }
        } catch {}
    }
}

function Get-CompareResult {
    param($Result)
    if (!$Result.MACAddress) { return '' }
    if ($Script:PreviousMap.ContainsKey($Result.MACAddress)) {
        $old = $Script:PreviousMap[$Result.MACAddress]
        if ($old.IPAddress -ne $Result.IPAddress) { return 'IP Changed' }
        return 'Existing'
    }
    return 'New Device'
}

function New-ScanResult {
    param([string]$IP,[string]$Mac,[string]$Host,[string]$Workgroup,[string]$Status,[int]$ResponseTime)
    $man = Get-Manufacturer $Mac
    $obj = [PSCustomObject]@{
        IPAddress = $IP
        MACAddress = $Mac
        HostName = $Host
        Workgroup = $Workgroup
        Manufacturer = $man
        Status = $Status
        ResponseTimeMs = $ResponseTime
        CompareResult = ''
        URL = "http://$IP/"
    }
    $obj.CompareResult = Get-CompareResult $obj
    return $obj
}

function Start-FastScan {
    param(
        [string[]]$IPs,
        [int]$TimeoutMs = 300,
        [int]$Threads = 100,
        [scriptblock]$OnResult,
        [scriptblock]$OnProgress
    )
    $Script:ScanStop = $false
    $total = $IPs.Count
    $done = 0
    $queue = [System.Collections.Queue]::Synchronized((New-Object System.Collections.Queue))
    foreach ($ip in $IPs) { $queue.Enqueue($ip) }

    $pool = [RunspaceFactory]::CreateRunspacePool(1, $Threads)
    $pool.ApartmentState = 'MTA'
    $pool.Open()
    $jobs = New-Object System.Collections.ArrayList

    $worker = {
        param($Queue, $Timeout)
        $results = @()
        while ($true) {
            $ip = $null
            [System.Threading.Monitor]::Enter($Queue.SyncRoot)
            try {
                if ($Queue.Count -gt 0) { $ip = $Queue.Dequeue() }
            } finally { [System.Threading.Monitor]::Exit($Queue.SyncRoot) }
            if (!$ip) { break }
            $sw = [System.Diagnostics.Stopwatch]::StartNew()
            $online = $false
            try {
                $p = New-Object System.Net.NetworkInformation.Ping
                $reply = $p.Send($ip, $Timeout)
                if ($reply.Status -eq 'Success') { $online = $true }
            } catch {}
            $sw.Stop()
            $results += [PSCustomObject]@{ IP=$ip; Online=$online; Ms=[int]$sw.ElapsedMilliseconds }
        }
        return $results
    }

    for ($i=0; $i -lt $Threads; $i++) {
        $ps = [PowerShell]::Create()
        $ps.RunspacePool = $pool
        [void]$ps.AddScript($worker).AddArgument($queue).AddArgument($TimeoutMs)
        $handle = $ps.BeginInvoke()
        [void]$jobs.Add([PSCustomObject]@{ PowerShell=$ps; Handle=$handle })
    }

    $allPing = @()
    foreach ($j in $jobs) {
        $out = $j.PowerShell.EndInvoke($j.Handle)
        $j.PowerShell.Dispose()
        $allPing += $out
        $done += $out.Count
        if ($OnProgress) { & $OnProgress $done $total }
        [System.Windows.Forms.Application]::DoEvents()
    }
    $pool.Close(); $pool.Dispose()

    $arpMap = Get-ArpTable
    $live = $allPing | Where-Object { $_.Online -or $arpMap.ContainsKey($_.IP) }
    $count = 0
    foreach ($r in $live) {
        if ($Script:ScanStop) { break }
        $mac = ''
        if ($arpMap.ContainsKey($r.IP)) { $mac = $arpMap[$r.IP] }
        $host = Resolve-HostNameFast $r.IP
        $wg = Resolve-WorkgroupFast $r.IP
        $status = if ($r.Online) { 'Online' } else { 'ARP Only' }
        $obj = New-ScanResult -IP $r.IP -Mac $mac -Host $host -Workgroup $wg -Status $status -ResponseTime $r.Ms
        if ($OnResult) { & $OnResult $obj }
        $count++
        if ($OnProgress) { & $OnProgress $count $live.Count }
        [System.Windows.Forms.Application]::DoEvents()
    }
}

Import-OuiDatabase
Load-PreviousScan

$form = New-Object System.Windows.Forms.Form
$form.Text = 'Fast MAC Scanner'
$form.Size = New-Object System.Drawing.Size(1120,720)
$form.StartPosition = 'CenterScreen'
$form.BackColor = [System.Drawing.Color]::FromArgb(245,247,250)
$form.Font = New-Object System.Drawing.Font('Segoe UI',9)

$header = New-Object System.Windows.Forms.Panel
$header.Dock = 'Top'
$header.Height = 70
$header.BackColor = [System.Drawing.Color]::FromArgb(35,45,60)
$form.Controls.Add($header)

$title = New-Object System.Windows.Forms.Label
$title.Text = 'Fast MAC Scanner - Colasoft Style'
$title.ForeColor = [System.Drawing.Color]::White
$title.Font = New-Object System.Drawing.Font('Segoe UI',16,[System.Drawing.FontStyle]::Bold)
$title.AutoSize = $true
$title.Location = New-Object System.Drawing.Point(18,16)
$header.Controls.Add($title)

$subtitle = New-Object System.Windows.Forms.Label
$subtitle.Text = 'IP, MAC, Host Name, Workgroup, Manufacturer, Compare Result'
$subtitle.ForeColor = [System.Drawing.Color]::Gainsboro
$subtitle.AutoSize = $true
$subtitle.Location = New-Object System.Drawing.Point(21,45)
$header.Controls.Add($subtitle)

$topPanel = New-Object System.Windows.Forms.Panel
$topPanel.Dock = 'Top'
$topPanel.Height = 90
$topPanel.BackColor = [System.Drawing.Color]::White
$form.Controls.Add($topPanel)

$lblRange = New-Object System.Windows.Forms.Label
$lblRange.Text = 'IP Range'
$lblRange.Location = New-Object System.Drawing.Point(18,14)
$lblRange.AutoSize = $true
$topPanel.Controls.Add($lblRange)

$cmbRange = New-Object System.Windows.Forms.ComboBox
$cmbRange.Location = New-Object System.Drawing.Point(18,36)
$cmbRange.Size = New-Object System.Drawing.Size(260,25)
$cmbRange.DropDownStyle = 'DropDown'
Get-LocalSubnets | ForEach-Object { [void]$cmbRange.Items.Add($_) }
if ($cmbRange.Items.Count -gt 0) { $cmbRange.SelectedIndex = 0 }
$topPanel.Controls.Add($cmbRange)

$lblTimeout = New-Object System.Windows.Forms.Label
$lblTimeout.Text = 'Timeout ms'
$lblTimeout.Location = New-Object System.Drawing.Point(295,14)
$lblTimeout.AutoSize = $true
$topPanel.Controls.Add($lblTimeout)

$numTimeout = New-Object System.Windows.Forms.NumericUpDown
$numTimeout.Location = New-Object System.Drawing.Point(295,36)
$numTimeout.Minimum = 50
$numTimeout.Maximum = 2000
$numTimeout.Value = 300
$numTimeout.Increment = 50
$topPanel.Controls.Add($numTimeout)

$lblThreads = New-Object System.Windows.Forms.Label
$lblThreads.Text = 'Threads'
$lblThreads.Location = New-Object System.Drawing.Point(410,14)
$lblThreads.AutoSize = $true
$topPanel.Controls.Add($lblThreads)

$numThreads = New-Object System.Windows.Forms.NumericUpDown
$numThreads.Location = New-Object System.Drawing.Point(410,36)
$numThreads.Minimum = 10
$numThreads.Maximum = 200
$numThreads.Value = 100
$numThreads.Increment = 10
$topPanel.Controls.Add($numThreads)

$btnScan = New-Object System.Windows.Forms.Button
$btnScan.Text = 'Start Scan'
$btnScan.Location = New-Object System.Drawing.Point(535,31)
$btnScan.Size = New-Object System.Drawing.Size(110,32)
$btnScan.BackColor = [System.Drawing.Color]::FromArgb(40,150,90)
$btnScan.ForeColor = [System.Drawing.Color]::White
$btnScan.FlatStyle = 'Flat'
$topPanel.Controls.Add($btnScan)

$btnStop = New-Object System.Windows.Forms.Button
$btnStop.Text = 'Stop'
$btnStop.Location = New-Object System.Drawing.Point(655,31)
$btnStop.Size = New-Object System.Drawing.Size(85,32)
$btnStop.Enabled = $false
$btnStop.BackColor = [System.Drawing.Color]::FromArgb(190,60,60)
$btnStop.ForeColor = [System.Drawing.Color]::White
$btnStop.FlatStyle = 'Flat'
$topPanel.Controls.Add($btnStop)

$btnExport = New-Object System.Windows.Forms.Button
$btnExport.Text = 'Export CSV'
$btnExport.Location = New-Object System.Drawing.Point(750,31)
$btnExport.Size = New-Object System.Drawing.Size(100,32)
$btnExport.FlatStyle = 'Flat'
$topPanel.Controls.Add($btnExport)

$btnOui = New-Object System.Windows.Forms.Button
$btnOui.Text = 'Update OUI'
$btnOui.Location = New-Object System.Drawing.Point(860,31)
$btnOui.Size = New-Object System.Drawing.Size(100,32)
$btnOui.FlatStyle = 'Flat'
$topPanel.Controls.Add($btnOui)

$lblSearch = New-Object System.Windows.Forms.Label
$lblSearch.Text = 'Search'
$lblSearch.Location = New-Object System.Drawing.Point(970,14)
$lblSearch.AutoSize = $true
$topPanel.Controls.Add($lblSearch)

$txtSearch = New-Object System.Windows.Forms.TextBox
$txtSearch.Location = New-Object System.Drawing.Point(970,36)
$txtSearch.Size = New-Object System.Drawing.Size(120,25)
$topPanel.Controls.Add($txtSearch)

$grid = New-Object System.Windows.Forms.DataGridView
$grid.Dock = 'Fill'
$grid.AllowUserToAddRows = $false
$grid.AllowUserToDeleteRows = $false
$grid.ReadOnly = $true
$grid.SelectionMode = 'FullRowSelect'
$grid.MultiSelect = $false
$grid.AutoSizeColumnsMode = 'Fill'
$grid.BackgroundColor = [System.Drawing.Color]::White
$grid.RowHeadersVisible = $false
$form.Controls.Add($grid)

$table = New-Object System.Data.DataTable
@('IPAddress','MACAddress','HostName','Workgroup','Manufacturer','Status','ResponseTimeMs','CompareResult','URL') | ForEach-Object { [void]$table.Columns.Add($_) }
$binding = New-Object System.Windows.Forms.BindingSource
$binding.DataSource = $table
$grid.DataSource = $binding

$statusPanel = New-Object System.Windows.Forms.Panel
$statusPanel.Dock = 'Bottom'
$statusPanel.Height = 54
$statusPanel.BackColor = [System.Drawing.Color]::White
$form.Controls.Add($statusPanel)

$progress = New-Object System.Windows.Forms.ProgressBar
$progress.Location = New-Object System.Drawing.Point(18,16)
$progress.Size = New-Object System.Drawing.Size(360,20)
$statusPanel.Controls.Add($progress)

$lblStatus = New-Object System.Windows.Forms.Label
$lblStatus.Text = 'Ready'
$lblStatus.Location = New-Object System.Drawing.Point(395,16)
$lblStatus.AutoSize = $true
$statusPanel.Controls.Add($lblStatus)

$lblCount = New-Object System.Windows.Forms.Label
$lblCount.Text = 'Devices: 0'
$lblCount.Location = New-Object System.Drawing.Point(800,16)
$lblCount.AutoSize = $true
$statusPanel.Controls.Add($lblCount)

$ctx = New-Object System.Windows.Forms.ContextMenuStrip
$miOpen = New-Object System.Windows.Forms.ToolStripMenuItem('Open Web Page')
$miCopyIP = New-Object System.Windows.Forms.ToolStripMenuItem('Copy IP')
$miCopyMAC = New-Object System.Windows.Forms.ToolStripMenuItem('Copy MAC')
[void]$ctx.Items.Add($miOpen); [void]$ctx.Items.Add($miCopyIP); [void]$ctx.Items.Add($miCopyMAC)
$grid.ContextMenuStrip = $ctx

$miOpen.Add_Click({
    if ($grid.SelectedRows.Count -gt 0) {
        $url = $grid.SelectedRows[0].Cells['URL'].Value
        if ($url) { Start-Process $url }
    }
})
$miCopyIP.Add_Click({ if ($grid.SelectedRows.Count -gt 0) { [System.Windows.Forms.Clipboard]::SetText($grid.SelectedRows[0].Cells['IPAddress'].Value) } })
$miCopyMAC.Add_Click({ if ($grid.SelectedRows.Count -gt 0) { [System.Windows.Forms.Clipboard]::SetText($grid.SelectedRows[0].Cells['MACAddress'].Value) } })
$grid.Add_CellDoubleClick({ if ($_.RowIndex -ge 0) { $url=$grid.Rows[$_.RowIndex].Cells['URL'].Value; if ($url) { Start-Process $url } } })

$txtSearch.Add_TextChanged({
    $s = $txtSearch.Text.Replace("'", "''")
    if ([string]::IsNullOrWhiteSpace($s)) { $binding.Filter = $null }
    else { $binding.Filter = "IPAddress LIKE '%$s%' OR MACAddress LIKE '%$s%' OR HostName LIKE '%$s%' OR Workgroup LIKE '%$s%' OR Manufacturer LIKE '%$s%' OR Status LIKE '%$s%' OR CompareResult LIKE '%$s%'" }
})

$btnOui.Add_Click({
    $lblStatus.Text = 'Downloading official IEEE OUI database...'
    [System.Windows.Forms.Application]::DoEvents()
    if (Download-OuiDatabase) { [System.Windows.Forms.MessageBox]::Show('OUI database updated successfully.','OUI Update','OK','Information') }
    else { [System.Windows.Forms.MessageBox]::Show('OUI download failed. Built-in OUI list will be used.','OUI Update','OK','Warning') }
    $lblStatus.Text = 'Ready'
})

$btnExport.Add_Click({
    if ($table.Rows.Count -eq 0) { [System.Windows.Forms.MessageBox]::Show('No scan results to export.','Export','OK','Information'); return }
    $save = New-Object System.Windows.Forms.SaveFileDialog
    $save.Filter = 'CSV files (*.csv)|*.csv'
    $save.FileName = 'MAC_Scan_' + (Get-Date -Format 'yyyyMMdd_HHmmss') + '.csv'
    if ($save.ShowDialog() -eq 'OK') {
        $rows = foreach ($r in $table.Rows) {
            [PSCustomObject]@{
                IPAddress=$r.IPAddress; MACAddress=$r.MACAddress; HostName=$r.HostName; Workgroup=$r.Workgroup; Manufacturer=$r.Manufacturer; Status=$r.Status; ResponseTimeMs=$r.ResponseTimeMs; CompareResult=$r.CompareResult; URL=$r.URL
            }
        }
        $rows | Export-Csv -Path $save.FileName -NoTypeInformation -Encoding UTF8
        [System.Windows.Forms.MessageBox]::Show('Export completed.','Export','OK','Information')
    }
})

$btnStop.Add_Click({
    $Script:ScanStop = $true
    $lblStatus.Text = 'Stopping scan...'
})

$btnScan.Add_Click({
    $range = $cmbRange.Text.Trim()
    $ips = @(Expand-IPRange $range)
    if ($ips.Count -eq 0) { [System.Windows.Forms.MessageBox]::Show('Enter valid range. Example: 10.209.110.1-254','Invalid Range','OK','Warning'); return }
    $table.Rows.Clear()
    $Script:LastResults = @()
    Load-PreviousScan
    $btnScan.Enabled = $false
    $btnStop.Enabled = $true
    $progress.Value = 0
    $progress.Maximum = 100
    $lblStatus.Text = "Scanning $($ips.Count) IPs..."
    $swTotal = [System.Diagnostics.Stopwatch]::StartNew()
    try {
        Start-FastScan -IPs $ips -TimeoutMs ([int]$numTimeout.Value) -Threads ([int]$numThreads.Value) `
            -OnResult {
                param($obj)
                $row = $table.NewRow()
                foreach ($c in $table.Columns) { $row[$c.ColumnName] = $obj.($c.ColumnName) }
                $table.Rows.Add($row)
                $Script:LastResults += $obj
                $lblCount.Text = "Devices: $($table.Rows.Count)"
            } `
            -OnProgress {
                param($done,$total)
                if ($total -gt 0) {
                    $pct = [Math]::Min(100, [int](($done / $total) * 100))
                    $progress.Value = $pct
                    $lblStatus.Text = "Progress: $pct%"
                }
            }
    } catch {
        [System.Windows.Forms.MessageBox]::Show($_.Exception.Message,'Scan Error','OK','Error')
    }
    $swTotal.Stop()
    if ($Script:LastResults.Count -gt 0) {
        try { $Script:LastResults | Export-Csv -Path $Script:HistoryPath -NoTypeInformation -Encoding UTF8 } catch {}
    }
    $progress.Value = 100
    $lblStatus.Text = "Completed in $([Math]::Round($swTotal.Elapsed.TotalSeconds,1)) sec"
    $btnScan.Enabled = $true
    $btnStop.Enabled = $false
})

$form.Add_Shown({ $form.Activate() })
[void][System.Windows.Forms.Application]::Run($form)
