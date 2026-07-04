# Advanced Colasoft-Style Device Scanner - PowerShell WinForms
# Fast LAN scanner with MAC, Hostname, Manufacturer, Device Type, Web Port, Compare Result, Export
# Works best in Windows PowerShell 5.1. Run as normal user. For best MAC discovery, same subnet is required.

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName Microsoft.VisualBasic

# Hide PowerShell console
try {
    Add-Type @"
using System;
using System.Runtime.InteropServices;
public class Win32Console {
    [DllImport("kernel32.dll")] public static extern IntPtr GetConsoleWindow();
    [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
}
"@
    $consolePtr = [Win32Console]::GetConsoleWindow()
    if ($consolePtr -ne [IntPtr]::Zero) { [Win32Console]::ShowWindow($consolePtr, 0) | Out-Null }
} catch {}

$ErrorActionPreference = 'SilentlyContinue'
[System.Windows.Forms.Application]::EnableVisualStyles()

$AppRoot = Join-Path $env:LOCALAPPDATA "AdvancedDeviceScanner"
$OutputRoot = Join-Path $AppRoot "Output"
$HistoryPath = Join-Path $AppRoot "last_scan.csv"
$OuiPath = Join-Path $AppRoot "oui.csv"
New-Item -ItemType Directory -Force -Path $AppRoot,$OutputRoot | Out-Null

$script:CancelScan = $false
$script:CurrentResults = New-Object System.Collections.Generic.List[object]
$script:OuiMap = @{}

function Add-Log {
    param([string]$Message)
    $time = Get-Date -Format "HH:mm:ss"
    if ($txtLog.InvokeRequired) {
        $txtLog.BeginInvoke([Action]{ $txtLog.AppendText("[$time] $Message`r`n") }) | Out-Null
    } else {
        $txtLog.AppendText("[$time] $Message`r`n")
    }
}

function Get-LocalSubnetBase {
    try {
        $ip = Get-NetIPAddress -AddressFamily IPv4 | Where-Object {
            $_.IPAddress -notlike '169.254*' -and $_.IPAddress -ne '127.0.0.1' -and $_.PrefixLength -le 24
        } | Select-Object -First 1
        if ($ip) {
            $parts = $ip.IPAddress.Split('.')
            return "$($parts[0]).$($parts[1]).$($parts[2])"
        }
    } catch {}
    try {
        $ip = [System.Net.Dns]::GetHostAddresses($env:COMPUTERNAME) | Where-Object { $_.AddressFamily -eq 'InterNetwork' -and $_.ToString() -notlike '169.254*' } | Select-Object -First 1
        if ($ip) { $p = $ip.ToString().Split('.'); return "$($p[0]).$($p[1]).$($p[2])" }
    } catch {}
    return "10.209.110"
}

function Normalize-Mac {
    param([string]$Mac)
    if ([string]::IsNullOrWhiteSpace($Mac)) { return "" }
    $hex = ($Mac -replace '[^0-9A-Fa-f]', '').ToUpper()
    if ($hex.Length -lt 12) { return $Mac.ToUpper() }
    return (($hex.Substring(0,2),$hex.Substring(2,2),$hex.Substring(4,2),$hex.Substring(6,2),$hex.Substring(8,2),$hex.Substring(10,2)) -join ':')
}

function Get-OuiPrefix {
    param([string]$Mac)
    $hex = ($Mac -replace '[^0-9A-Fa-f]', '').ToUpper()
    if ($hex.Length -ge 6) { return $hex.Substring(0,6) }
    return ""
}

function Load-BuiltinOui {
    $pairs = @{
        '00A859'='Fanvil Technology Co., Ltd.'; '0C383E'='Fanvil Technology Co., Ltd.'; '00096E'='Fanvil Technology Co., Ltd.'
        'F8B156'='HP Inc.'; '3CD92B'='HP Inc.'; '2C768A'='HP Inc.'; 'D48564'='Hewlett Packard'; 'B05ADA'='HP Inc.'
        '3C5282'='Dell Inc.'; 'F8CAB8'='Dell Inc.'; 'D067E5'='Dell Inc.'; 'B8CA3A'='Dell Inc.'; '002564'='Dell Inc.'
        '98FA9B'='Lenovo'; '54EE75'='Lenovo'; '6C4B90'='Liteon/Lenovo'; 'C85B76'='Lenovo'; '8C1645'='Lenovo'
        '001B24'='Cisco Systems'; 'F4CFD2'='Cisco Systems'; 'A41875'='Cisco Systems'; '3C5EC3'='Cisco Systems'
        '805EC0'='Yealink Network Technology'; '001565'='Yealink Network Technology'; '249AD8'='Yealink Network Technology'
        'C074AD'='Grandstream Networks'; '000B82'='Grandstream Networks'
        'E0CA3C'='Hikvision Digital Technology'; 'BCAD28'='Hikvision Digital Technology'; 'C056E3'='Hangzhou Hikvision'
        'D4E853'='Dahua Technology'; '3C1B0D'='Dahua Technology'; '4C11BF'='Zhejiang Dahua'
        '00606E'='ZKTeco'; 'E8ABFA'='ZKTeco'; '0027E3'='ZKTeco'
        '0050C2'='IEEE Assigned'; '000C29'='VMware'; '005056'='VMware'; '001C42'='Parallels'; '080027'='VirtualBox'
        '3C7C3F'='Samsung'; 'F0D1A9'='Apple'; 'DC2B2A'='Apple'; 'A4C361'='Apple'; 'F45C89'='Apple'
        'D850E6'='ASUSTek'; '2C4D54'='ASUSTek'; 'F46D04'='ASUSTek'; '7054D2'='Pegatron/ASUS'
        '001E8C'='ASUSTek'; '001F3B'='Intel Corporate'; 'A0369F'='Intel Corporate'; 'B49691'='Intel Corporate'
        '50E549'='Giga-Byte Technology'; '18C04D'='Giga-Byte Technology'; 'D8CB8A'='Micro-Star International'
        'B827EB'='Raspberry Pi Foundation'; 'DCA632'='Raspberry Pi Trading'; 'E45F01'='Raspberry Pi Trading'
        '001A79'='Zebra Technologies'; '002258'='Brother Industries'; '001B7A'='Canon'; '001E8F'='Canon'
        '00805F'='Kyocera'; '0001E6'='Hewlett Packard Printer'; '0025B3'='Hewlett Packard Printer'
        'F4F5D8'='Google'; 'D8EB46'='Google'; '60A4B7'='TP-Link'; 'F4F26D'='TP-Link'; '50C7BF'='TP-Link'
        'B4FBF9'='Ubiquiti'; 'F09FC2'='Ubiquiti'; '24A43C'='Ubiquiti'; 'E063DA'='Ubiquiti'
        'C4AD34'='Routerboard MikroTik'; 'D4CA6D'='MikroTik'; '18FD74'='Routerboard MikroTik'
        '0004F2'='Polycom'; '64167F'='Polycom'; '00085D'='Aastra'; '000413'='Snom Technology'
    }
    foreach($k in $pairs.Keys){ $script:OuiMap[$k] = $pairs[$k] }
}

function Load-OuiDatabase {
    Load-BuiltinOui
    if (Test-Path $OuiPath) {
        try {
            Import-Csv $OuiPath | ForEach-Object {
                if ($_.Assignment -and $_.'Organization Name') { $script:OuiMap[$_.Assignment.ToUpper()] = $_.'Organization Name' }
            }
            Add-Log "Loaded OUI database: $($script:OuiMap.Count) entries"
        } catch { Add-Log "OUI CSV load failed, using built-in database" }
    } else {
        Add-Log "Using built-in OUI database. Click Download OUI for full manufacturer list."
    }
}

function Download-OuiDatabase {
    try {
        Add-Log "Downloading IEEE OUI database..."
        Invoke-WebRequest "https://standards-oui.ieee.org/oui/oui.csv" -OutFile $OuiPath -UseBasicParsing -TimeoutSec 30
        $script:OuiMap.Clear(); Load-OuiDatabase
        [System.Windows.Forms.MessageBox]::Show("OUI database downloaded successfully.","OUI Database") | Out-Null
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Download failed. Internet may be blocked. Built-in OUI will continue.","OUI Database") | Out-Null
    }
}

function Get-Manufacturer {
    param([string]$Mac)
    $prefix = Get-OuiPrefix $Mac
    if ($prefix -and $script:OuiMap.ContainsKey($prefix)) { return $script:OuiMap[$prefix] }
    return "Unknown"
}

function Get-DeviceType {
    param([string]$Manufacturer,[string]$HostName,[string]$OpenPorts,[string]$WebTitle,[string]$Mac)
    $s = (($Manufacturer + ' ' + $HostName + ' ' + $OpenPorts + ' ' + $WebTitle) -as [string]).ToLower()
    $prefix = Get-OuiPrefix $Mac
    if ($prefix -in @('00A859','0C383E','00096E') -or $s -match 'fanvil') { return 'IP Phone - Fanvil' }
    if ($s -match 'zkteco|zkaccess|port 4370|4370') { return 'Biometric - ZKTeco' }
    if ($s -match 'saviour|essl|matrix|suprema|biometric|attendance') { return 'Biometric/Attendance' }
    if ($s -match 'dell|hp inc|hewlett|lenovo|asus|acer|intel|micro-star|gigabyte') { return 'PC/Laptop' }
    if ($s -match 'printer|canon|epson|brother|zebra|kyocera|ricoh|xerox') { return 'Printer' }
    if ($s -match 'hikvision|dahua|axis|uniview|cp plus|camera|nvr|dvr') { return 'CCTV/NVR' }
    if ($s -match 'cisco|aruba|juniper|netgear|d-link|tp-link|ubiquiti|ruckus|mikrotik|routerboard') { return 'Network Device' }
    if ($s -match 'vmware|virtualbox|hyper-v|parallels') { return 'Virtual Machine' }
    if ($s -match 'apple|samsung|xiaomi|oppo|vivo|oneplus') { return 'Mobile/Smart Device' }
    if ($OpenPorts -match '80|443|8080') { return 'Web Device' }
    return 'Unknown Device'
}

function Get-ArpMac {
    param([string]$IP)
    try {
        $arp = arp -a $IP | Out-String
        $line = ($arp -split "`r?`n") | Where-Object { $_ -match [regex]::Escape($IP) } | Select-Object -First 1
        if ($line -match '([0-9a-fA-F]{2}[-:]){5}[0-9a-fA-F]{2}') { return Normalize-Mac $matches[0] }
    } catch {}
    return ""
}

function Resolve-HostNameFast {
    param([string]$IP)
    try {
        $h = [System.Net.Dns]::GetHostEntry($IP).HostName
        if ($h) { return $h }
    } catch {}
    try {
        $nbt = nbtstat -A $IP 2>$null | Out-String
        $lines = $nbt -split "`r?`n"
        foreach($l in $lines) {
            if ($l -match '^\s*([^\s<]+)\s+<00>\s+UNIQUE') { return $matches[1].Trim() }
        }
    } catch {}
    return ""
}

function Test-PortFast {
    param([string]$IP,[int]$Port,[int]$TimeoutMs=250)
    try {
        $client = New-Object System.Net.Sockets.TcpClient
        $iar = $client.BeginConnect($IP,$Port,$null,$null)
        $ok = $iar.AsyncWaitHandle.WaitOne($TimeoutMs,$false)
        if ($ok -and $client.Connected) { $client.Close(); return $true }
        $client.Close()
    } catch {}
    return $false
}

function Get-WebTitleFast {
    param([string]$IP,[string]$Ports)
    if ($Ports -notmatch '80|8080|443') { return "" }
    foreach($p in @(80,8080,443)) {
        if ($Ports -notmatch "\b$p\b") { continue }
        try {
            $scheme = if($p -eq 443){'https'}else{'http'}
            $url = "${scheme}://${IP}:$p/"
            if ($p -eq 80) { $url = "http://${IP}/" }
            if ($p -eq 443) { $url = "https://${IP}/" }
            $req = [System.Net.WebRequest]::Create($url)
            $req.Timeout = 700
            $req.UserAgent = 'Mozilla/5.0'
            $resp = $req.GetResponse()
            $stream = $resp.GetResponseStream()
            $reader = New-Object System.IO.StreamReader($stream)
            $html = $reader.ReadToEnd()
            $reader.Close(); $resp.Close()
            if ($html -match '<title[^>]*>(.*?)</title>') { return ([System.Web.HttpUtility]::HtmlDecode($matches[1])).Trim() }
            if ($html -match 'ZKTeco|Fanvil|Saviour|Hikvision|Dahua|eSSL|Matrix') { return $matches[0] }
        } catch {}
    }
    return ""
}

function Get-WmiPcInfo {
    param([string]$IP,[string]$HostName)
    $target = if($HostName){$HostName}else{$IP}
    $info = @{ Brand=''; Model=''; Workgroup='' }
    try {
        $cs = Get-CimInstance Win32_ComputerSystem -ComputerName $target -OperationTimeoutSec 1
        if($cs) { $info.Brand = $cs.Manufacturer; $info.Model = $cs.Model; $info.Workgroup = if($cs.Domain){$cs.Domain}else{$cs.Workgroup} }
    } catch {}
    return $info
}

function Load-PreviousMap {
    $map = @{}
    if(Test-Path $HistoryPath) {
        try { Import-Csv $HistoryPath | ForEach-Object { if($_.MACAddress){ $map[$_.MACAddress] = $_ } } } catch {}
    }
    return $map
}

function Get-CompareResult {
    param($PreviousMap,[string]$Mac,[string]$IP)
    if(-not $Mac){ return "No MAC" }
    if(-not $PreviousMap.ContainsKey($Mac)) { return "New Device" }
    $old = $PreviousMap[$Mac]
    if($old.IPAddress -ne $IP) { return "IP Changed" }
    return "Existing"
}

function Scan-OneIP {
    param([string]$IP,[hashtable]$PreviousMap,[bool]$DeepWmi)
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    $pingOk = $false
    try { $pingOk = Test-Connection -ComputerName $IP -Count 1 -Quiet -BufferSize 16 -TimeToLive 64 }
    catch { $pingOk = $false }
    # Read ARP even if ping false because some devices block ICMP but arp may exist after connect attempt
    $mac = Get-ArpMac $IP
    if(-not $pingOk -and -not $mac) { return $null }
    $host = Resolve-HostNameFast $IP
    $ports = New-Object System.Collections.Generic.List[int]
    foreach($port in @(80,443,8080,4370,22,23,9100,161)) { if(Test-PortFast $IP $port 220) { $ports.Add($port) } }
    $portText = ($ports -join ',')
    $title = Get-WebTitleFast $IP $portText
    $manu = Get-Manufacturer $mac
    $pcBrand = '';$model='';$workgroup=''
    if($DeepWmi -and ($manu -match 'Dell|HP|Hewlett|Lenovo|ASUS|Acer|Intel|Micro-Star|Gigabyte' -or $host)) {
        $pc = Get-WmiPcInfo $IP $host
        $pcBrand=$pc.Brand; $model=$pc.Model; $workgroup=$pc.Workgroup
        if($pcBrand){ $manu = $pcBrand }
    }
    $dtype = Get-DeviceType $manu $host $portText $title $mac
    $status = if($pingOk){'Online'}else{'MAC/Web Only'}
    $compare = Get-CompareResult $PreviousMap $mac $IP
    $url = if($portText -match '443') { "https://$IP/" } elseif($portText -match '80') { "http://$IP/" } elseif($portText -match '8080') { "http://${IP}:8080/" } else { "" }
    $sw.Stop()
    [PSCustomObject]@{
        IPAddress=$IP; MACAddress=$mac; HostName=$host; Workgroup=$workgroup; Manufacturer=$manu; PCBrand=$pcBrand; Model=$model; DeviceType=$dtype;
        Status=$status; ResponseMS=$sw.ElapsedMilliseconds; OpenPorts=$portText; WebTitle=$title; URL=$url; CompareResult=$compare; LastSeen=(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
    }
}

function Export-ResultsCsv {
    if($script:CurrentResults.Count -eq 0){ [System.Windows.Forms.MessageBox]::Show('No results to export.','Export') | Out-Null; return }
    $path = Join-Path $OutputRoot ("DeviceScan_" + (Get-Date -Format 'yyyyMMdd_HHmmss') + ".csv")
    $script:CurrentResults | Export-Csv $path -NoTypeInformation -Encoding UTF8
    Copy-Item $path $HistoryPath -Force
    [System.Windows.Forms.MessageBox]::Show("Exported:`n$path","Export") | Out-Null
    Start-Process explorer.exe "/select,`"$path`""
}

function Export-ResultsHtml {
    if($script:CurrentResults.Count -eq 0){ return }
    $path = Join-Path $OutputRoot ("DeviceScan_" + (Get-Date -Format 'yyyyMMdd_HHmmss') + ".html")
    $html = $script:CurrentResults | ConvertTo-Html -Title 'Advanced Device Scan Report' -PreContent "<h2>Advanced Device Scan Report</h2><p>Generated: $(Get-Date)</p>" | Out-String
    $html | Set-Content $path -Encoding UTF8
    Start-Process $path
}

function Add-GridRow {
    param($r)
    if($grid.InvokeRequired){ $grid.BeginInvoke([Action]{ Add-GridRow $r }) | Out-Null; return }
    $idx = $grid.Rows.Add($r.IPAddress,$r.MACAddress,$r.HostName,$r.Workgroup,$r.Manufacturer,$r.PCBrand,$r.Model,$r.DeviceType,$r.Status,$r.ResponseMS,$r.OpenPorts,$r.WebTitle,$r.URL,$r.CompareResult,$r.LastSeen)
    if($r.DeviceType -match 'Fanvil|Biometric') { $grid.Rows[$idx].DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(232,245,233) }
    elseif($r.CompareResult -eq 'New Device') { $grid.Rows[$idx].DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(255,248,225) }
    elseif($r.DeviceType -eq 'Web Device') { $grid.Rows[$idx].DefaultCellStyle.BackColor = [System.Drawing.Color]::FromArgb(227,242,253) }
}

function Start-ScanAsync {
    $script:CancelScan = $false
    $grid.Rows.Clear(); $script:CurrentResults.Clear()
    $base = $txtSubnet.Text.Trim()
    $from = [int]$numFrom.Value; $to = [int]$numTo.Value
    $threads = [int]$numThreads.Value
    $deepWmi = $chkWmi.Checked
    if($from -gt $to){ [System.Windows.Forms.MessageBox]::Show('From value must be less than To value.','Range') | Out-Null; return }
    $btnStart.Enabled=$false; $btnStop.Enabled=$true; $progress.Value=0
    $lblStatus.Text = 'Status: Scanning...'
    $ips = for($i=$from;$i -le $to;$i++){ "$base.$i" }
    $total = $ips.Count
    $prev = Load-PreviousMap
    Add-Log "Scan started: $base.$from - $base.$to with $threads workers"

    $job = [System.ComponentModel.BackgroundWorker]::new()
    $job.WorkerReportsProgress = $true
    $job.DoWork += {
        $pool = [RunspaceFactory]::CreateRunspacePool(1,$threads)
        $pool.Open()
        $tasks = New-Object System.Collections.Generic.List[object]
        $funcs = ${function:Normalize-Mac}.ToString()+${function:Get-OuiPrefix}.ToString()+${function:Get-ArpMac}.ToString()+${function:Resolve-HostNameFast}.ToString()+${function:Test-PortFast}.ToString()+${function:Get-WebTitleFast}.ToString()+${function:Get-Manufacturer}.ToString()+${function:Get-DeviceType}.ToString()+${function:Get-WmiPcInfo}.ToString()+${function:Get-CompareResult}.ToString()+${function:Scan-OneIP}.ToString()
        foreach($ip in $ips){
            $ps = [PowerShell]::Create()
            $ps.RunspacePool = $pool
            [void]$ps.AddScript($funcs)
            [void]$ps.AddScript({ param($IP,$PreviousMap,$DeepWmi,$OuiMap) $script:OuiMap=$OuiMap; Scan-OneIP -IP $IP -PreviousMap $PreviousMap -DeepWmi $DeepWmi })
            [void]$ps.AddArgument($ip); [void]$ps.AddArgument($prev); [void]$ps.AddArgument($deepWmi); [void]$ps.AddArgument($script:OuiMap)
            $tasks.Add([PSCustomObject]@{ PowerShell=$ps; Handle=$ps.BeginInvoke(); IP=$ip }) | Out-Null
        }
        $done=0
        while($tasks.Count -gt 0){
            for($x=$tasks.Count-1;$x -ge 0;$x--){
                if($script:CancelScan){ break }
                $t=$tasks[$x]
                if($t.Handle.IsCompleted){
                    $res=$t.PowerShell.EndInvoke($t.Handle)
                    $t.PowerShell.Dispose()
                    $tasks.RemoveAt($x)
                    $done++
                    if($res){ foreach($rr in $res){ $job.ReportProgress([int](($done/$total)*100),$rr) } } else { $job.ReportProgress([int](($done/$total)*100),$null) }
                }
            }
            if($script:CancelScan){ break }
            Start-Sleep -Milliseconds 50
        }
        foreach($t in $tasks){ try{$t.PowerShell.Stop();$t.PowerShell.Dispose()}catch{} }
        $pool.Close(); $pool.Dispose()
    }
    $job.ProgressChanged += {
        if($_.ProgressPercentage -ge 0 -and $_.ProgressPercentage -le 100){ $progress.Value=$_.ProgressPercentage }
        if($_.UserState){ $script:CurrentResults.Add($_.UserState); Add-GridRow $_.UserState }
        $lblCount.Text = "Devices: $($script:CurrentResults.Count)"
    }
    $job.RunWorkerCompleted += {
        $btnStart.Enabled=$true; $btnStop.Enabled=$false; $progress.Value=100
        $lblStatus.Text = if($script:CancelScan){'Status: Stopped'}else{'Status: Completed'}
        Add-Log "Scan completed. Devices found: $($script:CurrentResults.Count)"
        if($script:CurrentResults.Count -gt 0){ $script:CurrentResults | Export-Csv $HistoryPath -NoTypeInformation -Encoding UTF8 }
    }
    $job.RunWorkerAsync()
}

# UI
$form = New-Object System.Windows.Forms.Form
$form.Text = 'Advanced Device Scanner'
$form.Size = New-Object System.Drawing.Size(1280,760)
$form.StartPosition='CenterScreen'
$form.BackColor=[System.Drawing.Color]::FromArgb(245,247,250)
$form.Font = New-Object System.Drawing.Font('Segoe UI',9)
try { $form.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon((Get-Command powershell.exe).Source) } catch {}

$header = New-Object System.Windows.Forms.Panel
$header.Dock='Top'; $header.Height=72; $header.BackColor=[System.Drawing.Color]::FromArgb(32,44,57)
$form.Controls.Add($header)

$title = New-Object System.Windows.Forms.Label
$title.Text='Advanced Device Scanner'
$title.ForeColor=[System.Drawing.Color]::White
$title.Font=New-Object System.Drawing.Font('Segoe UI Semibold',18)
$title.AutoSize=$true; $title.Location=New-Object System.Drawing.Point(18,12)
$header.Controls.Add($title)

$sub = New-Object System.Windows.Forms.Label
$sub.Text='Colasoft-style fast MAC, hostname, manufacturer, PC, Fanvil, ZKTeco, Saviour and web device discovery'
$sub.ForeColor=[System.Drawing.Color]::Gainsboro
$sub.AutoSize=$true; $sub.Location=New-Object System.Drawing.Point(22,45)
$header.Controls.Add($sub)

$top = New-Object System.Windows.Forms.Panel
$top.Dock='Top'; $top.Height=86; $top.BackColor=[System.Drawing.Color]::White
$form.Controls.Add($top)

function New-Label($text,$x,$y){ $l=New-Object System.Windows.Forms.Label; $l.Text=$text; $l.Location=New-Object System.Drawing.Point($x,$y); $l.AutoSize=$true; $top.Controls.Add($l); return $l }
New-Label 'Subnet' 18 14 | Out-Null
$txtSubnet=New-Object System.Windows.Forms.TextBox; $txtSubnet.Location=New-Object System.Drawing.Point(18,36); $txtSubnet.Width=115; $txtSubnet.Text=Get-LocalSubnetBase; $top.Controls.Add($txtSubnet)
New-Label 'From' 145 14 | Out-Null
$numFrom=New-Object System.Windows.Forms.NumericUpDown; $numFrom.Location=New-Object System.Drawing.Point(145,36); $numFrom.Minimum=1;$numFrom.Maximum=254;$numFrom.Value=1;$numFrom.Width=58; $top.Controls.Add($numFrom)
New-Label 'To' 215 14 | Out-Null
$numTo=New-Object System.Windows.Forms.NumericUpDown; $numTo.Location=New-Object System.Drawing.Point(215,36); $numTo.Minimum=1;$numTo.Maximum=254;$numTo.Value=254;$numTo.Width=58; $top.Controls.Add($numTo)
New-Label 'Workers' 285 14 | Out-Null
$numThreads=New-Object System.Windows.Forms.NumericUpDown; $numThreads.Location=New-Object System.Drawing.Point(285,36); $numThreads.Minimum=5;$numThreads.Maximum=200;$numThreads.Value=80;$numThreads.Width=68; $top.Controls.Add($numThreads)
$chkWmi=New-Object System.Windows.Forms.CheckBox; $chkWmi.Text='Deep Windows PC info'; $chkWmi.Location=New-Object System.Drawing.Point(370,37); $chkWmi.Width=155; $chkWmi.Checked=$false; $top.Controls.Add($chkWmi)

$btnStart=New-Object System.Windows.Forms.Button; $btnStart.Text='Start Scan'; $btnStart.Location=New-Object System.Drawing.Point(545,30); $btnStart.Size=New-Object System.Drawing.Size(100,32); $btnStart.BackColor=[System.Drawing.Color]::FromArgb(46,125,50); $btnStart.ForeColor=[System.Drawing.Color]::White; $btnStart.FlatStyle='Flat'; $top.Controls.Add($btnStart)
$btnStop=New-Object System.Windows.Forms.Button; $btnStop.Text='Stop'; $btnStop.Location=New-Object System.Drawing.Point(655,30); $btnStop.Size=New-Object System.Drawing.Size(80,32); $btnStop.BackColor=[System.Drawing.Color]::FromArgb(183,28,28); $btnStop.ForeColor=[System.Drawing.Color]::White; $btnStop.FlatStyle='Flat'; $btnStop.Enabled=$false; $top.Controls.Add($btnStop)
$btnExport=New-Object System.Windows.Forms.Button; $btnExport.Text='Export CSV'; $btnExport.Location=New-Object System.Drawing.Point(745,30); $btnExport.Size=New-Object System.Drawing.Size(95,32); $btnExport.FlatStyle='Flat'; $top.Controls.Add($btnExport)
$btnHtml=New-Object System.Windows.Forms.Button; $btnHtml.Text='HTML Report'; $btnHtml.Location=New-Object System.Drawing.Point(850,30); $btnHtml.Size=New-Object System.Drawing.Size(100,32); $btnHtml.FlatStyle='Flat'; $top.Controls.Add($btnHtml)
$btnOui=New-Object System.Windows.Forms.Button; $btnOui.Text='Download OUI'; $btnOui.Location=New-Object System.Drawing.Point(960,30); $btnOui.Size=New-Object System.Drawing.Size(105,32); $btnOui.FlatStyle='Flat'; $top.Controls.Add($btnOui)
$btnOpen=New-Object System.Windows.Forms.Button; $btnOpen.Text='Open Output'; $btnOpen.Location=New-Object System.Drawing.Point(1075,30); $btnOpen.Size=New-Object System.Drawing.Size(100,32); $btnOpen.FlatStyle='Flat'; $top.Controls.Add($btnOpen)

$statusPanel=New-Object System.Windows.Forms.Panel; $statusPanel.Dock='Bottom'; $statusPanel.Height=34; $statusPanel.BackColor=[System.Drawing.Color]::WhiteSmoke; $form.Controls.Add($statusPanel)
$lblStatus=New-Object System.Windows.Forms.Label; $lblStatus.Text='Status: Ready'; $lblStatus.Location=New-Object System.Drawing.Point(12,8); $lblStatus.AutoSize=$true; $statusPanel.Controls.Add($lblStatus)
$lblCount=New-Object System.Windows.Forms.Label; $lblCount.Text='Devices: 0'; $lblCount.Location=New-Object System.Drawing.Point(160,8); $lblCount.AutoSize=$true; $statusPanel.Controls.Add($lblCount)
$progress=New-Object System.Windows.Forms.ProgressBar; $progress.Location=New-Object System.Drawing.Point(260,7); $progress.Size=New-Object System.Drawing.Size(340,18); $statusPanel.Controls.Add($progress)

$split=New-Object System.Windows.Forms.SplitContainer; $split.Dock='Fill'; $split.Orientation='Horizontal'; $split.SplitterDistance=480; $form.Controls.Add($split)

grid= $null
$grid=New-Object System.Windows.Forms.DataGridView
$grid.Dock='Fill'; $grid.ReadOnly=$true; $grid.AllowUserToAddRows=$false; $grid.AllowUserToDeleteRows=$false; $grid.SelectionMode='FullRowSelect'; $grid.MultiSelect=$false; $grid.AutoSizeColumnsMode='DisplayedCells'; $grid.BackgroundColor=[System.Drawing.Color]::White; $grid.RowHeadersVisible=$false
$cols='IPAddress','MACAddress','HostName','Workgroup','Manufacturer','PCBrand','Model','DeviceType','Status','ResponseMS','OpenPorts','WebTitle','URL','CompareResult','LastSeen'
foreach($c in $cols){ [void]$grid.Columns.Add($c,$c) }
$split.Panel1.Controls.Add($grid)

$txtLog=New-Object System.Windows.Forms.TextBox; $txtLog.Dock='Fill'; $txtLog.Multiline=$true; $txtLog.ScrollBars='Vertical'; $txtLog.ReadOnly=$true; $txtLog.BackColor=[System.Drawing.Color]::FromArgb(250,250,250); $split.Panel2.Controls.Add($txtLog)

$menu=New-Object System.Windows.Forms.ContextMenuStrip
$miOpen=New-Object System.Windows.Forms.ToolStripMenuItem('Open Web URL')
$miCopy=New-Object System.Windows.Forms.ToolStripMenuItem('Copy IP')
$menu.Items.AddRange(@($miOpen,$miCopy))
$grid.ContextMenuStrip=$menu

$btnStart.Add_Click({ Start-ScanAsync })
$btnStop.Add_Click({ $script:CancelScan=$true; Add-Log 'Stop requested...' })
$btnExport.Add_Click({ Export-ResultsCsv })
$btnHtml.Add_Click({ Export-ResultsHtml })
$btnOui.Add_Click({ Download-OuiDatabase })
$btnOpen.Add_Click({ Start-Process explorer.exe $OutputRoot })
$grid.Add_CellDoubleClick({ if($grid.CurrentRow){ $url=$grid.CurrentRow.Cells['URL'].Value; if($url){ Start-Process $url } } })
$miOpen.Add_Click({ if($grid.CurrentRow){ $url=$grid.CurrentRow.Cells['URL'].Value; if($url){ Start-Process $url } } })
$miCopy.Add_Click({ if($grid.CurrentRow){ [Windows.Forms.Clipboard]::SetText($grid.CurrentRow.Cells['IPAddress'].Value) } })

Load-OuiDatabase
Add-Log 'Ready. Use Download OUI once for full manufacturer accuracy.'
[void]$form.ShowDialog()
