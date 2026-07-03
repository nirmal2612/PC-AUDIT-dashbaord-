# ============================================================
# FANVIL WEB DEVICE SCANNER APPLICATION
# Purpose : Enter IP segment, click Scan, fetch all HTTP/Web device IPs,
#           show result in GUI, and export report to D:\fanvil
# Models  : Useful for Fanvil X301G / X3SG segments
# Output  : D:\fanvil\Fanvil_WebDevice_Report_yyyyMMdd_HHmmss.csv
#           D:\fanvil\Fanvil_WebDevice_Report_yyyyMMdd_HHmmss.xlsx if Excel installed
# ============================================================

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# ---------------- GLOBAL SETTINGS ----------------

$script:OutputFolder = "D:\fanvil"
$script:ScanResults = New-Object System.Collections.ArrayList
$script:IsScanning = $false
$script:CancelScan = $false

if (!(Test-Path $script:OutputFolder)) {
    New-Item -ItemType Directory -Path $script:OutputFolder -Force | Out-Null
}

# ---------------- HELPER FUNCTIONS ----------------

function Get-LocalIPv4 {
    try {
        $ip = Get-NetIPAddress -AddressFamily IPv4 |
            Where-Object {
                $_.IPAddress -notlike "127.*" -and
                $_.IPAddress -notlike "169.254.*" -and
                $_.IPAddress -match "^\d{1,3}(\.\d{1,3}){3}$"
            } |
            Select-Object -First 1 -ExpandProperty IPAddress
        return $ip
    }
    catch {
        try {
            $ip = Get-WmiObject Win32_NetworkAdapterConfiguration |
                Where-Object { $_.IPEnabled -eq $true -and $_.IPAddress } |
                ForEach-Object { $_.IPAddress } |
                Where-Object {
                    $_ -notlike "127.*" -and
                    $_ -notlike "169.254.*" -and
                    $_ -match "^\d{1,3}(\.\d{1,3}){3}$"
                } |
                Select-Object -First 1
            return $ip
        }
        catch {
            return $null
        }
    }
}

function Get-IPBase {
    param([string]$IP)
    $p = $IP.Split(".")
    return "$($p[0]).$($p[1]).$($p[2])"
}

function Get-PCMac {
    try {
        $mac = Get-NetAdapter |
            Where-Object { $_.Status -eq "Up" } |
            Select-Object -First 1 -ExpandProperty MacAddress
        if ($mac) { return $mac }
    }
    catch {}

    try {
        $mac = Get-WmiObject Win32_NetworkAdapterConfiguration |
            Where-Object { $_.IPEnabled -eq $true -and $_.MACAddress } |
            Select-Object -First 1 -ExpandProperty MACAddress
        if ($mac) { return $mac }
    }
    catch {}

    return "N/A"
}

function Add-Log {
    param([string]$Text)

    $time = Get-Date -Format "HH:mm:ss"
    $txtLog.AppendText("[$time] $Text`r`n")
    $txtLog.SelectionStart = $txtLog.Text.Length
    $txtLog.ScrollToCaret()
    [System.Windows.Forms.Application]::DoEvents()
}

function Get-TitleFromHtml {
    param([string]$Html)

    if (!$Html) { return "N/A" }

    $m = [regex]::Match($Html, "<title[^>]*>(.*?)</title>", "IgnoreCase,Singleline")
    if ($m.Success) {
        $title = $m.Groups[1].Value
        $title = $title -replace "<.*?>", ""
        $title = $title -replace "&nbsp;", " "
        $title = $title -replace "&amp;", "&"
        $title = $title.Trim()
        if ($title) { return $title }
    }

    return "N/A"
}

function Detect-DeviceType {
    param([string]$Html)

    if (!$Html) { return "Web Device" }

    if ($Html -match "X301G") { return "Fanvil X301G" }
    if ($Html -match "X3SG")  { return "Fanvil X3SG" }
    if ($Html -match "Fanvil") { return "Fanvil Phone" }
    if ($Html -match "VoIP|SIP Phone|IP Phone") { return "IP Phone / VoIP Device" }
    if ($Html -match "printer|Print|JetDirect|LaserJet|Epson|Canon|Brother") { return "Printer" }
    if ($Html -match "Router|Gateway|Switch|Cisco|D-Link|TP-Link|Netgear") { return "Network Device" }

    return "Web Device"
}

function Test-WebDevice {
    param(
        [string]$IP,
        [int]$TimeoutSec
    )

    $result = [PSCustomObject]@{
        IsWebDevice = $false
        IPAddress   = $IP
        Port        = "80"
        Status      = "No Response"
        Title       = "N/A"
        DeviceType  = "N/A"
        Server      = "N/A"
        URL         = "http://$IP"
        Error       = ""
        Html        = ""
    }

    try {
        $response = Invoke-WebRequest `
            -Uri "http://$IP" `
            -UseBasicParsing `
            -TimeoutSec $TimeoutSec `
            -ErrorAction Stop

        $server = "N/A"
        try {
            if ($response.Headers["Server"]) {
                $server = $response.Headers["Server"]
            }
        }
        catch {}

        $html = ""
        if ($response.Content) {
            $html = [string]$response.Content
        }

        $result.IsWebDevice = $true
        $result.Status = "HTTP OK"
        $result.Title = Get-TitleFromHtml $html
        $result.DeviceType = Detect-DeviceType $html
        $result.Server = $server
        $result.Html = $html
    }
    catch {
        $result.Error = $_.Exception.Message
    }

    return $result
}

function Parse-SegmentInput {
    param([string]$InputText)

    $input = $InputText.Trim()

    if ($input -match "^(\d{1,3}\.){2}\d{1,3}$") {
        return [PSCustomObject]@{
            IsValid = $true
            Base    = $input
            Start   = 1
            End     = 254
            Message = ""
        }
    }

    if ($input -match "^((\d{1,3}\.){3})(\d{1,3})\s*-\s*(\d{1,3})$") {
        $base = $matches[1].TrimEnd(".")
        $start = [int]$matches[3]
        $end = [int]$matches[4]

        if ($start -lt 1 -or $end -gt 254 -or $start -gt $end) {
            return [PSCustomObject]@{
                IsValid = $false
                Message = "Invalid IP range. Use 1-254."
            }
        }

        return [PSCustomObject]@{
            IsValid = $true
            Base    = $base
            Start   = $start
            End     = $end
            Message = ""
        }
    }

    if ($input -match "^((\d{1,3}\.){3})(\d{1,3})$") {
        $p = $input.Split(".")
        $base = "$($p[0]).$($p[1]).$($p[2])"
        return [PSCustomObject]@{
            IsValid = $true
            Base    = $base
            Start   = 1
            End     = 254
            Message = ""
        }
    }

    return [PSCustomObject]@{
        IsValid = $false
        Message = "Enter segment like 172.29.10 or range like 172.29.10.1-254"
    }
}

function Export-Results {
    if ($script:ScanResults.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show("No results to export.", "Export", "OK", "Information") | Out-Null
        return
    }

    if (!(Test-Path $script:OutputFolder)) {
        New-Item -ItemType Directory -Path $script:OutputFolder -Force | Out-Null
    }

    $stamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $csvPath = Join-Path $script:OutputFolder "Fanvil_WebDevice_Report_$stamp.csv"
    $xlsxPath = Join-Path $script:OutputFolder "Fanvil_WebDevice_Report_$stamp.xlsx"

    $exportRows = foreach ($r in $script:ScanResults) {
        [PSCustomObject]@{
            ScanDate   = $r.ScanDate
            PCName     = $r.PCName
            PCIP       = $r.PCIP
            PCMAC      = $r.PCMAC
            Segment    = $r.Segment
            DeviceIP   = $r.DeviceIP
            URL        = $r.URL
            Status     = $r.Status
            DeviceType = $r.DeviceType
            Title      = $r.Title
            Server     = $r.Server
        }
    }

    $exportRows | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8

    $excelCreated = $false

    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false

        $workbook = $excel.Workbooks.Open($csvPath)
        $sheet = $workbook.Worksheets.Item(1)
        $sheet.Name = "Web Devices"

        $used = $sheet.UsedRange
        $used.EntireColumn.AutoFit() | Out-Null

        $header = $sheet.Range("A1:K1")
        $header.Font.Bold = $true
        $header.Interior.ColorIndex = 15

        $workbook.SaveAs($xlsxPath, 51)
        $workbook.Close($true)
        $excel.Quit()

        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($sheet) | Out-Null
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) | Out-Null
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null

        $excelCreated = $true
    }
    catch {
        $excelCreated = $false
    }

    if ($excelCreated) {
        Add-Log "Export completed: $xlsxPath"
        [System.Windows.Forms.MessageBox]::Show("Excel report saved:`r`n$xlsxPath`r`n`r`nCSV also saved:`r`n$csvPath", "Export Completed", "OK", "Information") | Out-Null
    }
    else {
        Add-Log "CSV export completed: $csvPath"
        Add-Log "Excel app not available. CSV file can be opened in Excel."
        [System.Windows.Forms.MessageBox]::Show("CSV report saved:`r`n$csvPath`r`n`r`nExcel application was not available, so XLSX was not created.", "Export Completed", "OK", "Information") | Out-Null
    }
}

# ---------------- FORM DESIGN ----------------

$form = New-Object System.Windows.Forms.Form
$form.Text = "Fanvil Web Device Scanner"
$form.Size = New-Object System.Drawing.Size(980, 650)
$form.StartPosition = "CenterScreen"
$form.MinimumSize = New-Object System.Drawing.Size(920, 600)
$form.Font = New-Object System.Drawing.Font("Segoe UI", 9)

$lblTitle = New-Object System.Windows.Forms.Label
$lblTitle.Text = "Fanvil Web Device Scanner"
$lblTitle.Font = New-Object System.Drawing.Font("Segoe UI", 16, [System.Drawing.FontStyle]::Bold)
$lblTitle.AutoSize = $true
$lblTitle.Location = New-Object System.Drawing.Point(15, 15)
$form.Controls.Add($lblTitle)

$lblSub = New-Object System.Windows.Forms.Label
$lblSub.Text = "Enter IP segment on one PC, scan all web devices, and export report to D:\fanvil"
$lblSub.AutoSize = $true
$lblSub.Location = New-Object System.Drawing.Point(18, 48)
$form.Controls.Add($lblSub)

$lblSegment = New-Object System.Windows.Forms.Label
$lblSegment.Text = "IP Segment / Range:"
$lblSegment.AutoSize = $true
$lblSegment.Location = New-Object System.Drawing.Point(18, 86)
$form.Controls.Add($lblSegment)

$txtSegment = New-Object System.Windows.Forms.TextBox
$txtSegment.Location = New-Object System.Drawing.Point(130, 82)
$txtSegment.Size = New-Object System.Drawing.Size(230, 25)
$form.Controls.Add($txtSegment)

$localIp = Get-LocalIPv4
if ($localIp) {
    $txtSegment.Text = Get-IPBase $localIp
}

$lblHint = New-Object System.Windows.Forms.Label
$lblHint.Text = "Example: 172.29.10 or 172.29.10.1-254"
$lblHint.AutoSize = $true
$lblHint.Location = New-Object System.Drawing.Point(370, 86)
$form.Controls.Add($lblHint)

$lblTimeout = New-Object System.Windows.Forms.Label
$lblTimeout.Text = "Timeout:"
$lblTimeout.AutoSize = $true
$lblTimeout.Location = New-Object System.Drawing.Point(18, 122)
$form.Controls.Add($lblTimeout)

$numTimeout = New-Object System.Windows.Forms.NumericUpDown
$numTimeout.Location = New-Object System.Drawing.Point(130, 118)
$numTimeout.Size = New-Object System.Drawing.Size(70, 25)
$numTimeout.Minimum = 1
$numTimeout.Maximum = 10
$numTimeout.Value = 2
$form.Controls.Add($numTimeout)

$lblTimeoutSec = New-Object System.Windows.Forms.Label
$lblTimeoutSec.Text = "seconds per IP"
$lblTimeoutSec.AutoSize = $true
$lblTimeoutSec.Location = New-Object System.Drawing.Point(207, 122)
$form.Controls.Add($lblTimeoutSec)

$btnScan = New-Object System.Windows.Forms.Button
$btnScan.Text = "Scan"
$btnScan.Location = New-Object System.Drawing.Point(620, 78)
$btnScan.Size = New-Object System.Drawing.Size(100, 32)
$form.Controls.Add($btnScan)

$btnCancel = New-Object System.Windows.Forms.Button
$btnCancel.Text = "Cancel"
$btnCancel.Location = New-Object System.Drawing.Point(730, 78)
$btnCancel.Size = New-Object System.Drawing.Size(100, 32)
$btnCancel.Enabled = $false
$form.Controls.Add($btnCancel)

$btnExport = New-Object System.Windows.Forms.Button
$btnExport.Text = "Export Excel"
$btnExport.Location = New-Object System.Drawing.Point(840, 78)
$btnExport.Size = New-Object System.Drawing.Size(110, 32)
$form.Controls.Add($btnExport)

$progress = New-Object System.Windows.Forms.ProgressBar
$progress.Location = New-Object System.Drawing.Point(20, 155)
$progress.Size = New-Object System.Drawing.Size(930, 22)
$progress.Minimum = 0
$progress.Maximum = 100
$progress.Value = 0
$form.Controls.Add($progress)

$lblStatus = New-Object System.Windows.Forms.Label
$lblStatus.Text = "Ready"
$lblStatus.AutoSize = $true
$lblStatus.Location = New-Object System.Drawing.Point(20, 183)
$form.Controls.Add($lblStatus)

$grid = New-Object System.Windows.Forms.DataGridView
$grid.Location = New-Object System.Drawing.Point(20, 210)
$grid.Size = New-Object System.Drawing.Size(930, 285)
$grid.AllowUserToAddRows = $false
$grid.AllowUserToDeleteRows = $false
$grid.ReadOnly = $true
$grid.SelectionMode = "FullRowSelect"
$grid.AutoSizeColumnsMode = "Fill"
$form.Controls.Add($grid)

$table = New-Object System.Data.DataTable
$table.Columns.Add("S.No") | Out-Null
$table.Columns.Add("Device IP") | Out-Null
$table.Columns.Add("URL") | Out-Null
$table.Columns.Add("Status") | Out-Null
$table.Columns.Add("Device Type") | Out-Null
$table.Columns.Add("Title") | Out-Null
$table.Columns.Add("Server") | Out-Null
$grid.DataSource = $table

$txtLog = New-Object System.Windows.Forms.TextBox
$txtLog.Location = New-Object System.Drawing.Point(20, 510)
$txtLog.Size = New-Object System.Drawing.Size(930, 85)
$txtLog.Multiline = $true
$txtLog.ScrollBars = "Vertical"
$txtLog.ReadOnly = $true
$form.Controls.Add($txtLog)

# ---------------- BUTTON EVENTS ----------------

$btnScan.Add_Click({

    if ($script:IsScanning) {
        return
    }

    $parsed = Parse-SegmentInput $txtSegment.Text

    if (!$parsed.IsValid) {
        [System.Windows.Forms.MessageBox]::Show($parsed.Message, "Invalid Segment", "OK", "Warning") | Out-Null
        return
    }

    $script:ScanResults.Clear()
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

    $pcIp = Get-LocalIPv4
    $pcMac = Get-PCMac
    $segmentText = "$base.$start-$end"

    Add-Log "Scan started for $segmentText"
    Add-Log "Output folder: $script:OutputFolder"

    for ($i = $start; $i -le $end; $i++) {

        if ($script:CancelScan) {
            Add-Log "Scan cancelled by user."
            break
        }

        $ip = "$base.$i"
        $count++

        $lblStatus.Text = "Scanning $ip ($count / $total)"
        $percent = [int](($count / $total) * 100)
        if ($percent -gt 100) { $percent = 100 }
        $progress.Value = $percent

        [System.Windows.Forms.Application]::DoEvents()

        $web = Test-WebDevice -IP $ip -TimeoutSec $timeout

        if ($web.IsWebDevice) {
            $found++

            $row = $table.NewRow()
            $row["S.No"] = $found
            $row["Device IP"] = $web.IPAddress
            $row["URL"] = $web.URL
            $row["Status"] = $web.Status
            $row["Device Type"] = $web.DeviceType
            $row["Title"] = $web.Title
            $row["Server"] = $web.Server
            $table.Rows.Add($row)

            $record = [PSCustomObject]@{
                ScanDate   = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
                PCName     = $env:COMPUTERNAME
                PCIP       = $pcIp
                PCMAC      = $pcMac
                Segment    = $segmentText
                DeviceIP   = $web.IPAddress
                URL        = $web.URL
                Status     = $web.Status
                DeviceType = $web.DeviceType
                Title      = $web.Title
                Server     = $web.Server
            }

            [void]$script:ScanResults.Add($record)

            Add-Log "Web device found: $ip | $($web.DeviceType) | $($web.Title)"
        }
    }

    $lblStatus.Text = "Completed. Web devices found: $found"
    $progress.Value = 100

    $script:IsScanning = $false
    $btnScan.Enabled = $true
    $btnCancel.Enabled = $false
    $btnExport.Enabled = $true

    Add-Log "Scan completed. Total web devices found: $found"

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

# ---------------- START APP ----------------

Add-Log "Application ready."
Add-Log "Enter segment and click Scan."

[void]$form.ShowDialog()
