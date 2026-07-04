# ============================================================
# WEB DEVICE IP + MAC SCANNER APPLICATION
# Purpose:
#   Enter one IP segment, click Scan,
#   find all web devices,
#   fetch each IP related MAC using:
#       ping <ip> -n 1
#       arp -a <ip>
#   export report to D:\fanvil as CSV and XLSX if Excel is installed.
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

    return [PSCustomObject]@{ IsValid=$false; Message="Enter IP segment like 10.209.110 or range like 10.209.110.1-254" }
}

function Add-Log {
    param([string]$Text)
    $time = Get-Date -Format "HH:mm:ss"
    $txtLog.AppendText("[$time] $Text`r`n")
    $txtLog.SelectionStart = $txtLog.Text.Length
    $txtLog.ScrollToCaret()
    [System.Windows.Forms.Application]::DoEvents()
}

function Get-PageTitle {
    param([string]$Html)

    if (!$Html) { return "N/A" }

    try {
        $m = [regex]::Match($Html, "<title[^>]*>(.*?)</title>", "IgnoreCase,Singleline")
        if ($m.Success) {
            $title = $m.Groups[1].Value
            $title = $title -replace "<[^>]+>", ""
            $title = $title -replace "&nbsp;", " "
            $title = $title -replace "&amp;", "&"
            $title = [regex]::Replace($title, "\s+", " ").Trim()
            if ($title) { return $title }
        }
    }
    catch {}

    return "N/A"
}

function Detect-DeviceType {
    param([string]$Html)

    if (!$Html) { return "Web Device" }
    if ($Html -match "X301G") { return "Fanvil X301G" }
    if ($Html -match "X3SG") { return "Fanvil X3SG" }
    if ($Html -match "Fanvil") { return "Fanvil Phone" }
    if ($Html -match "SIP|VoIP|IP Phone") { return "IP Phone / VoIP" }
    if ($Html -match "Printer|LaserJet|Epson|Canon|Brother") { return "Printer" }
    if ($Html -match "Router|Switch|Gateway|Cisco|D-Link|TP-Link|Netgear") { return "Network Device" }
    return "Web Device"
}

function Test-WebDevice {
    param([string]$IP, [int]$TimeoutSec)

    $obj = [PSCustomObject]@{
        IsWebDevice = $false
        Status      = "No Response"
        Title       = "N/A"
        DeviceType  = "N/A"
        Server      = "N/A"
        URL         = "http://$IP"
    }

    try {
        $response = Invoke-WebRequest -Uri "http://$IP" -UseBasicParsing -TimeoutSec $TimeoutSec -ErrorAction Stop

        $html = ""
        if ($response.Content) { $html = [string]$response.Content }

        $server = "N/A"
        try {
            if ($response.Headers["Server"]) { $server = [string]$response.Headers["Server"] }
        } catch {}

        $obj.IsWebDevice = $true
        $obj.Status = "HTTP OK"
        $obj.Title = Get-PageTitle $html
        $obj.DeviceType = Detect-DeviceType $html
        $obj.Server = $server
    }
    catch {}

    return $obj
}

function Get-MacFromArp {
    param([string]$IP)

    try { ping $IP -n 1 | Out-Null } catch {}
    Start-Sleep -Milliseconds 100

    try {
        $arp = arp -a $IP
        $mac = ($arp | Select-String "([0-9a-fA-F]{2}-){5}[0-9a-fA-F]{2}").Matches.Value | Select-Object -First 1
        if ($mac) { return ($mac.ToUpper() -replace "-", ":") }
    }
    catch {}

    return "N/A"
}

function Export-Results {
    if ($script:Results.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("No result available to export.", "Export", "OK", "Information") | Out-Null
        return
    }

    if (!(Test-Path $script:OutputFolder)) {
        New-Item -ItemType Directory -Path $script:OutputFolder -Force | Out-Null
    }

    $stamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $csvPath = Join-Path $script:OutputFolder "Web_Device_IP_MAC_Report_$stamp.csv"
    $xlsxPath = Join-Path $script:OutputFolder "Web_Device_IP_MAC_Report_$stamp.xlsx"

    $script:Results | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8

    $excelCreated = $false
    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false

        $workbook = $excel.Workbooks.Open($csvPath)
        $sheet = $workbook.Worksheets.Item(1)
        $sheet.Name = "Web Devices"
        $sheet.UsedRange.EntireColumn.AutoFit() | Out-Null
        $sheet.Range("A1:J1").Font.Bold = $true
        $sheet.Range("A1:J1").Interior.ColorIndex = 15
        $workbook.SaveAs($xlsxPath, 51)
        $workbook.Close($true)
        $excel.Quit()

        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($sheet) | Out-Null
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) | Out-Null
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
        $excelCreated = $true
    }
    catch { $excelCreated = $false }

    if ($excelCreated) {
        Add-Log "Excel saved: $xlsxPath"
        [System.Windows.Forms.MessageBox]::Show("Excel saved:`r`n$xlsxPath`r`n`r`nCSV also saved:`r`n$csvPath", "Export Completed", "OK", "Information") | Out-Null
    }
    else {
        Add-Log "CSV saved: $csvPath"
        [System.Windows.Forms.MessageBox]::Show("CSV saved:`r`n$csvPath`r`n`r`nExcel not found, so XLSX not created.", "Export Completed", "OK", "Information") | Out-Null
    }
}

$form = New-Object System.Windows.Forms.Form
$form.Text = "Web Device IP MAC Scanner"
$form.Size = New-Object System.Drawing.Size(1100, 700)
$form.StartPosition = "CenterScreen"
$form.Font = New-Object System.Drawing.Font("Segoe UI", 9)

$lblTitle = New-Object System.Windows.Forms.Label
$lblTitle.Text = "Web Device IP MAC Scanner"
$lblTitle.Font = New-Object System.Drawing.Font("Segoe UI", 16, [System.Drawing.FontStyle]::Bold)
$lblTitle.AutoSize = $true
$lblTitle.Location = New-Object System.Drawing.Point(15, 15)
$form.Controls.Add($lblTitle)

$lblSub = New-Object System.Windows.Forms.Label
$lblSub.Text = "Enter one IP segment, scan all web devices, fetch MAC using ping + arp -a, and export to D:\fanvil"
$lblSub.AutoSize = $true
$lblSub.Location = New-Object System.Drawing.Point(18, 48)
$form.Controls.Add($lblSub)

$lblSegment = New-Object System.Windows.Forms.Label
$lblSegment.Text = "IP Segment:"
$lblSegment.AutoSize = $true
$lblSegment.Location = New-Object System.Drawing.Point(20, 85)
$form.Controls.Add($lblSegment)

$txtSegment = New-Object System.Windows.Forms.TextBox
$txtSegment.Location = New-Object System.Drawing.Point(110, 81)
$txtSegment.Size = New-Object System.Drawing.Size(190, 25)
$form.Controls.Add($txtSegment)

$localIP = Get-LocalIPv4
if ($localIP) { $txtSegment.Text = Get-IPBase $localIP }

$lblHint = New-Object System.Windows.Forms.Label
$lblHint.Text = "Example: 10.209.110 or 10.209.110.1-254"
$lblHint.AutoSize = $true
$lblHint.Location = New-Object System.Drawing.Point(310, 85)
$form.Controls.Add($lblHint)

$lblTimeout = New-Object System.Windows.Forms.Label
$lblTimeout.Text = "Timeout:"
$lblTimeout.AutoSize = $true
$lblTimeout.Location = New-Object System.Drawing.Point(20, 122)
$form.Controls.Add($lblTimeout)

$numTimeout = New-Object System.Windows.Forms.NumericUpDown
$numTimeout.Location = New-Object System.Drawing.Point(110, 118)
$numTimeout.Size = New-Object System.Drawing.Size(60, 25)
$numTimeout.Minimum = 1
$numTimeout.Maximum = 10
$numTimeout.Value = 2
$form.Controls.Add($numTimeout)

$lblTimeoutText = New-Object System.Windows.Forms.Label
$lblTimeoutText.Text = "seconds per IP"
$lblTimeoutText.AutoSize = $true
$lblTimeoutText.Location = New-Object System.Drawing.Point(180, 122)
$form.Controls.Add($lblTimeoutText)

$btnScan = New-Object System.Windows.Forms.Button
$btnScan.Text = "Scan"
$btnScan.Location = New-Object System.Drawing.Point(745, 78)
$btnScan.Size = New-Object System.Drawing.Size(90, 32)
$form.Controls.Add($btnScan)

$btnCancel = New-Object System.Windows.Forms.Button
$btnCancel.Text = "Cancel"
$btnCancel.Location = New-Object System.Drawing.Point(845, 78)
$btnCancel.Size = New-Object System.Drawing.Size(90, 32)
$btnCancel.Enabled = $false
$form.Controls.Add($btnCancel)

$btnExport = New-Object System.Windows.Forms.Button
$btnExport.Text = "Export"
$btnExport.Location = New-Object System.Drawing.Point(945, 78)
$btnExport.Size = New-Object System.Drawing.Size(90, 32)
$form.Controls.Add($btnExport)

$progress = New-Object System.Windows.Forms.ProgressBar
$progress.Location = New-Object System.Drawing.Point(20, 155)
$progress.Size = New-Object System.Drawing.Size(1015, 22)
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
$grid.Size = New-Object System.Drawing.Size(1015, 330)
$grid.AllowUserToAddRows = $false
$grid.AllowUserToDeleteRows = $false
$grid.ReadOnly = $true
$grid.SelectionMode = "FullRowSelect"
$grid.AutoSizeColumnsMode = "Fill"
$form.Controls.Add($grid)

$table = New-Object System.Data.DataTable
@("S.No","Device IP","MAC Address","URL","Status","Device Type","Title","Server","Scan Time") | ForEach-Object {
    [void]$table.Columns.Add($_)
}
$grid.DataSource = $table

$txtLog = New-Object System.Windows.Forms.TextBox
$txtLog.Location = New-Object System.Drawing.Point(20, 555)
$txtLog.Size = New-Object System.Drawing.Size(1015, 90)
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

    $base = $parsed.Base
    $start = [int]$parsed.Start
    $end = [int]$parsed.End
    $timeout = [int]$numTimeout.Value
    $total = ($end - $start) + 1
    $count = 0
    $found = 0

    Add-Log "Scan started: $base.$start-$end"
    Add-Log "Output folder: $script:OutputFolder"
    Add-Log "MAC method: ping IP -n 1, then arp -a IP"

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
        $lblStatus.Text = "Checking $ip ($count / $total)"
        [System.Windows.Forms.Application]::DoEvents()

        $web = Test-WebDevice -IP $ip -TimeoutSec $timeout
        if (!$web.IsWebDevice) { continue }

        $mac = Get-MacFromArp -IP $ip
        $found++
        $scanTime = Get-Date -Format "yyyy-MM-dd HH:mm:ss"

        $record = [PSCustomObject]@{
            ScanTime    = $scanTime
            Segment     = "$base.$start-$end"
            DeviceIP    = $ip
            MACAddress  = $mac
            URL         = $web.URL
            Status      = $web.Status
            DeviceType  = $web.DeviceType
            Title       = $web.Title
            Server      = $web.Server
            PCName      = $env:COMPUTERNAME
        }

        [void]$script:Results.Add($record)

        $row = $table.NewRow()
        $row["S.No"] = $found
        $row["Device IP"] = $ip
        $row["MAC Address"] = $mac
        $row["URL"] = $web.URL
        $row["Status"] = $web.Status
        $row["Device Type"] = $web.DeviceType
        $row["Title"] = $web.Title
        $row["Server"] = $web.Server
        $row["Scan Time"] = $scanTime
        $table.Rows.Add($row)

        Add-Log "Web device found: $ip | MAC: $mac | $($web.DeviceType)"
    }

    $progress.Value = 100
    $lblStatus.Text = "Completed. Web devices found: $found"
    $script:IsScanning = $false
    $btnScan.Enabled = $true
    $btnCancel.Enabled = $false
    $btnExport.Enabled = $true
    Add-Log "Scan completed. Web devices found: $found"

    if ($found -gt 0) { Export-Results }
})

$btnCancel.Add_Click({
    $script:CancelScan = $true
    $lblStatus.Text = "Cancelling..."
})

$btnExport.Add_Click({ Export-Results })

Add-Log "Application ready."
Add-Log "Enter IP segment and click Scan."

[void]$form.ShowDialog()
