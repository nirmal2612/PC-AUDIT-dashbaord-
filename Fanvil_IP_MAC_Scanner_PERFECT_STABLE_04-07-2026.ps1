# ============================================================
# 📞 FANVIL WEB DEVICE IP + MAC SCANNER
# Fresh Stable Version - 04-07-2026
# ------------------------------------------------------------
# Features:
#   ✅ Enter IP segment
#   ✅ Start Scan / Stop Scan / Open Output
#   ✅ Finds all web devices
#   ✅ Gets MAC using: ping IP -n 1 + arp -a IP
#   ✅ Strict Fanvil detection only if page contains Fanvil/X301G/X3SG
#   ✅ Brand = Fanvil only when confirmed
#   ✅ Login Status = Confirmed only for Fanvil
#   ✅ Model / Extension Number / SIP Username left empty
#   ✅ URL column contains exact scanned URL
#   ✅ Exports CSV + XLSX to D:\fanvil
#   ✅ Errors shown in log, app will not auto-close
# ============================================================

try {
    Add-Type -AssemblyName System.Windows.Forms
    Add-Type -AssemblyName System.Drawing
}
catch {
    Write-Host "Unable to load WinForms: $($_.Exception.Message)"
    pause
    exit
}

$ErrorActionPreference = "Continue"
$script:OutputFolder = "D:\fanvil"
$script:CancelScan = $false
$script:IsScanning = $false
$script:Results = New-Object System.Collections.ArrayList

try {
    if (!(Test-Path $script:OutputFolder)) {
        New-Item -ItemType Directory -Path $script:OutputFolder -Force | Out-Null
    }
}
catch {
    [System.Windows.Forms.MessageBox]::Show("Cannot create output folder D:\fanvil`r`n$($_.Exception.Message)", "Folder Error")
}

function Write-Log {
    param([string]$Message)

    try {
        $time = Get-Date -Format "HH:mm:ss"
        $script:txtLog.AppendText("[$time] $Message`r`n")
        $script:txtLog.SelectionStart = $script:txtLog.Text.Length
        $script:txtLog.ScrollToCaret()
        [System.Windows.Forms.Application]::DoEvents()
    }
    catch {}
}

function Get-LocalIPv4 {
    try {
        $ip = Get-NetIPAddress -AddressFamily IPv4 -ErrorAction Stop |
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
            return ""
        }
    }
}

function Get-IPBase {
    param([string]$IP)
    try {
        $p = $IP.Split(".")
        return "$($p[0]).$($p[1]).$($p[2])"
    }
    catch {
        return ""
    }
}

function Get-MacFromArp {
    param([string]$IP)

    try {
        ping $IP -n 1 -w 700 | Out-Null
        Start-Sleep -Milliseconds 150

        $arpOutput = arp -a $IP 2>$null | Out-String

        foreach ($line in ($arpOutput -split "`r?`n")) {
            if ($line -match [regex]::Escape($IP) -and $line -match "([0-9a-fA-F]{2}[-:]){5}[0-9a-fA-F]{2}") {
                return ($Matches[0].ToUpper() -replace "-", ":")
            }
        }
    }
    catch {}

    return ""
}

function Get-PageTitle {
    param([string]$Html)

    try {
        if (!$Html) { return "" }

        $m = [regex]::Match($Html, "<title[^>]*>(.*?)</title>", "IgnoreCase,Singleline")
        if ($m.Success) {
            $title = $m.Groups[1].Value
            $title = $title -replace "<[^>]+>", ""
            $title = $title -replace "&nbsp;", " "
            $title = $title -replace "&amp;", "&"
            $title = [regex]::Replace($title, "\s+", " ").Trim()
            return $title
        }
    }
    catch {}

    return ""
}

function Test-WebDevice {
    param(
        [string]$IP,
        [int]$TimeoutSec
    )

    $result = [PSCustomObject]@{
        IsWebDevice = $false
        Url         = "http://$IP"
        Status      = ""
        Brand       = "Web Device"
        LoginStatus = ""
        Title       = ""
        Server      = ""
    }

    try {
        $response = Invoke-WebRequest -Uri "http://$IP" -UseBasicParsing -TimeoutSec $TimeoutSec -ErrorAction Stop

        $html = ""
        if ($response.Content) {
            $html = [string]$response.Content
        }

        $server = ""
        try {
            if ($response.Headers["Server"]) {
                $server = [string]$response.Headers["Server"]
            }
        }
        catch {}

        $title = Get-PageTitle $html
        $checkText = "$html $title $server"

        $isFanvil = $false

        # Strict detection only.
        # Generic SIP / VoIP / GoAhead / Boa will NOT be Fanvil.
        if ($checkText -match "(?i)\bFanvil\b") {
            $isFanvil = $true
        }
        elseif ($checkText -match "(?i)\bX301G\b|\bX3SG\b|\bX301\b|\bX3S\b") {
            $isFanvil = $true
        }

        $result.IsWebDevice = $true
        $result.Status = "HTTP OK"
        $result.Title = $title
        $result.Server = $server

        if ($isFanvil) {
            $result.Brand = "Fanvil"
            $result.LoginStatus = "Confirmed"
        }
        else {
            $result.Brand = "Web Device"
            $result.LoginStatus = ""
        }
    }
    catch {}

    return $result
}

function Export-Results {
    try {
        if ($script:Results.Count -eq 0) {
            [System.Windows.Forms.MessageBox]::Show("No results to export.", "Export")
            return
        }

        if (!(Test-Path $script:OutputFolder)) {
            New-Item -ItemType Directory -Path $script:OutputFolder -Force | Out-Null
        }

        $stamp = Get-Date -Format "yyyyMMdd_HHmmss"
        $csvPath = Join-Path $script:OutputFolder "Fanvil_Web_Device_Report_$stamp.csv"
        $xlsxPath = Join-Path $script:OutputFolder "Fanvil_Web_Device_Report_$stamp.xlsx"

        $script:Results | Export-Csv -Path $csvPath -NoTypeInformation -Encoding UTF8
        Write-Log "CSV saved: $csvPath"

        $excelDone = $false

        try {
            $excel = New-Object -ComObject Excel.Application
            $excel.Visible = $false
            $excel.DisplayAlerts = $false

            $wb = $excel.Workbooks.Open($csvPath)
            $ws = $wb.Worksheets.Item(1)
            $ws.Name = "Fanvil Scan"

            $ws.UsedRange.EntireColumn.AutoFit() | Out-Null
            $ws.Range("A1:M1").Font.Bold = $true
            $ws.Range("A1:M1").Interior.ColorIndex = 15

            $wb.SaveAs($xlsxPath, 51)
            $wb.Close($true)
            $excel.Quit()

            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($ws) | Out-Null
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($wb) | Out-Null
            [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null

            $excelDone = $true
        }
        catch {
            Write-Log "Excel export skipped. CSV is available."
        }

        if ($excelDone) {
            Write-Log "Excel saved: $xlsxPath"
            [System.Windows.Forms.MessageBox]::Show("Export completed.`r`n`r`nExcel:`r`n$xlsxPath`r`n`r`nCSV:`r`n$csvPath", "Export Completed")
        }
        else {
            [System.Windows.Forms.MessageBox]::Show("CSV export completed.`r`n`r`n$csvPath", "Export Completed")
        }
    }
    catch {
        Write-Log "Export error: $($_.Exception.Message)"
        [System.Windows.Forms.MessageBox]::Show("Export error:`r`n$($_.Exception.Message)", "Export Error")
    }
}

function Start-Scan {
    try {
        if ($script:IsScanning) { return }

        $segment = $script:txtSegment.Text.Trim()
        $start = [int]$script:numStart.Value
        $end = [int]$script:numEnd.Value
        $timeout = [int]$script:numTimeout.Value

        if ($segment -notmatch "^\d{1,3}\.\d{1,3}\.\d{1,3}$") {
            [System.Windows.Forms.MessageBox]::Show("Enter segment like 10.209.110", "Invalid Segment")
            return
        }

        if ($start -gt $end) {
            [System.Windows.Forms.MessageBox]::Show("Start IP must be less than End IP.", "Invalid Range")
            return
        }

        $script:CancelScan = $false
        $script:IsScanning = $true
        $script:Results.Clear()
        $script:grid.Rows.Clear()
        $script:progress.Value = 0

        $script:btnStart.Enabled = $false
        $script:btnStop.Enabled = $true
        $script:btnExport.Enabled = $false

        $total = ($end - $start) + 1
        $checked = 0
        $found = 0

        Write-Log "🚀 Scan started: $segment.$start - $segment.$end"
        Write-Log "📌 MAC method: ping IP -n 1 + arp -a IP"
        Write-Log "📁 Output folder: $script:OutputFolder"

        for ($i = $start; $i -le $end; $i++) {
            if ($script:CancelScan) {
                Write-Log "⛔ Scan stopped."
                break
            }

            $ip = "$segment.$i"
            $checked++

            $percent = [int](($checked / $total) * 100)
            if ($percent -gt 100) { $percent = 100 }
            $script:progress.Value = $percent
            $script:lblStatus.Text = "Checking $ip ($checked / $total)"
            [System.Windows.Forms.Application]::DoEvents()

            $web = Test-WebDevice -IP $ip -TimeoutSec $timeout

            if (-not $web.IsWebDevice) {
                continue
            }

            $mac = Get-MacFromArp -IP $ip

            $found++
            $scanTime = Get-Date -Format "yyyy-MM-dd HH:mm:ss"

            $record = [PSCustomObject]@{
                "S.No"             = $found
                "IP Address"       = $ip
                "MAC Address"      = $mac
                "Brand"            = $web.Brand
                "Model"            = ""
                "Extension Number" = ""
                "SIP Username"     = ""
                "URL"              = $web.Url
                "Login Status"     = $web.LoginStatus
                "Web Status"       = $web.Status
                "Title"            = $web.Title
                "Server"           = $web.Server
                "Scan Time"        = $scanTime
            }

            [void]$script:Results.Add($record)

            [void]$script:grid.Rows.Add(
                $found,
                $ip,
                $mac,
                $web.Brand,
                "",
                "",
                "",
                $web.Url,
                $web.LoginStatus,
                $web.Status,
                $scanTime
            )

            Write-Log "✅ Found: $ip | MAC=$mac | Brand=$($web.Brand) | URL=$($web.Url)"
        }

        $script:lblStatus.Text = "Completed. Found: $found"
        $script:progress.Value = 100

        Write-Log "✅ Scan completed. Web devices found: $found"

        if ($found -gt 0) {
            Export-Results
        }
    }
    catch {
        Write-Log "Scan error: $($_.Exception.Message)"
        [System.Windows.Forms.MessageBox]::Show("Scan error:`r`n$($_.Exception.Message)", "Scan Error")
    }
    finally {
        $script:IsScanning = $false
        $script:btnStart.Enabled = $true
        $script:btnStop.Enabled = $false
        $script:btnExport.Enabled = $true
    }
}

function Stop-Scan {
    $script:CancelScan = $true
    Write-Log "Stop requested. Please wait..."
}

function Open-Output {
    try {
        if (!(Test-Path $script:OutputFolder)) {
            New-Item -ItemType Directory -Path $script:OutputFolder -Force | Out-Null
        }
        Invoke-Item $script:OutputFolder
    }
    catch {
        Write-Log "Open output error: $($_.Exception.Message)"
    }
}

# ================= UI =================

$form = New-Object System.Windows.Forms.Form
$form.Text = "📞 Fanvil IP MAC Scanner"
$form.Size = New-Object System.Drawing.Size(1120, 700)
$form.StartPosition = "CenterScreen"
$form.MinimumSize = New-Object System.Drawing.Size(1030, 620)
$form.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$form.ShowIcon = $false

$lblTitle = New-Object System.Windows.Forms.Label
$lblTitle.Text = "📞 Fanvil IP MAC Scanner"
$lblTitle.Font = New-Object System.Drawing.Font("Segoe UI", 16, [System.Drawing.FontStyle]::Bold)
$lblTitle.Location = New-Object System.Drawing.Point(20, 15)
$lblTitle.Size = New-Object System.Drawing.Size(600, 35)
$form.Controls.Add($lblTitle)

$lblSub = New-Object System.Windows.Forms.Label
$lblSub.Text = "Scan web devices, fetch IP related MAC using ARP, and export Excel/CSV to D:\fanvil"
$lblSub.Location = New-Object System.Drawing.Point(23, 48)
$lblSub.Size = New-Object System.Drawing.Size(850, 22)
$form.Controls.Add($lblSub)

$lblSegment = New-Object System.Windows.Forms.Label
$lblSegment.Text = "IP Segment"
$lblSegment.Location = New-Object System.Drawing.Point(22, 85)
$lblSegment.Size = New-Object System.Drawing.Size(80, 25)
$form.Controls.Add($lblSegment)

$txtSegment = New-Object System.Windows.Forms.TextBox
$txtSegment.Location = New-Object System.Drawing.Point(105, 81)
$txtSegment.Size = New-Object System.Drawing.Size(150, 25)

$local = Get-LocalIPv4
if ($local) {
    $txtSegment.Text = Get-IPBase $local
}
else {
    $txtSegment.Text = "10.209.110"
}

$form.Controls.Add($txtSegment)
$script:txtSegment = $txtSegment

$lblRange = New-Object System.Windows.Forms.Label
$lblRange.Text = "Range"
$lblRange.Location = New-Object System.Drawing.Point(275, 85)
$lblRange.Size = New-Object System.Drawing.Size(45, 25)
$form.Controls.Add($lblRange)

$numStart = New-Object System.Windows.Forms.NumericUpDown
$numStart.Location = New-Object System.Drawing.Point(325, 81)
$numStart.Size = New-Object System.Drawing.Size(60, 25)
$numStart.Minimum = 1
$numStart.Maximum = 254
$numStart.Value = 1
$form.Controls.Add($numStart)
$script:numStart = $numStart

$numEnd = New-Object System.Windows.Forms.NumericUpDown
$numEnd.Location = New-Object System.Drawing.Point(395, 81)
$numEnd.Size = New-Object System.Drawing.Size(60, 25)
$numEnd.Minimum = 1
$numEnd.Maximum = 254
$numEnd.Value = 254
$form.Controls.Add($numEnd)
$script:numEnd = $numEnd

$lblTimeout = New-Object System.Windows.Forms.Label
$lblTimeout.Text = "Timeout"
$lblTimeout.Location = New-Object System.Drawing.Point(475, 85)
$lblTimeout.Size = New-Object System.Drawing.Size(55, 25)
$form.Controls.Add($lblTimeout)

$numTimeout = New-Object System.Windows.Forms.NumericUpDown
$numTimeout.Location = New-Object System.Drawing.Point(535, 81)
$numTimeout.Size = New-Object System.Drawing.Size(55, 25)
$numTimeout.Minimum = 1
$numTimeout.Maximum = 10
$numTimeout.Value = 2
$form.Controls.Add($numTimeout)
$script:numTimeout = $numTimeout

$btnStart = New-Object System.Windows.Forms.Button
$btnStart.Text = "▶ Start Scan"
$btnStart.Location = New-Object System.Drawing.Point(620, 76)
$btnStart.Size = New-Object System.Drawing.Size(120, 34)
$btnStart.Add_Click({ Start-Scan })
$form.Controls.Add($btnStart)
$script:btnStart = $btnStart

$btnStop = New-Object System.Windows.Forms.Button
$btnStop.Text = "⏹ Stop Scan"
$btnStop.Location = New-Object System.Drawing.Point(750, 76)
$btnStop.Size = New-Object System.Drawing.Size(120, 34)
$btnStop.Enabled = $false
$btnStop.Add_Click({ Stop-Scan })
$form.Controls.Add($btnStop)
$script:btnStop = $btnStop

$btnExport = New-Object System.Windows.Forms.Button
$btnExport.Text = "📤 Export"
$btnExport.Location = New-Object System.Drawing.Point(880, 76)
$btnExport.Size = New-Object System.Drawing.Size(100, 34)
$btnExport.Add_Click({ Export-Results })
$form.Controls.Add($btnExport)
$script:btnExport = $btnExport

$btnOutput = New-Object System.Windows.Forms.Button
$btnOutput.Text = "📂 Open Output"
$btnOutput.Location = New-Object System.Drawing.Point(990, 76)
$btnOutput.Size = New-Object System.Drawing.Size(110, 34)
$btnOutput.Add_Click({ Open-Output })
$form.Controls.Add($btnOutput)

$progress = New-Object System.Windows.Forms.ProgressBar
$progress.Location = New-Object System.Drawing.Point(22, 125)
$progress.Size = New-Object System.Drawing.Size(1075, 18)
$form.Controls.Add($progress)
$script:progress = $progress

$lblStatus = New-Object System.Windows.Forms.Label
$lblStatus.Text = "Ready"
$lblStatus.Location = New-Object System.Drawing.Point(22, 148)
$lblStatus.Size = New-Object System.Drawing.Size(600, 20)
$form.Controls.Add($lblStatus)
$script:lblStatus = $lblStatus

$grid = New-Object System.Windows.Forms.DataGridView
$grid.Location = New-Object System.Drawing.Point(22, 175)
$grid.Size = New-Object System.Drawing.Size(1075, 340)
$grid.Anchor = "Top,Left,Right,Bottom"
$grid.AllowUserToAddRows = $false
$grid.AllowUserToDeleteRows = $false
$grid.ReadOnly = $true
$grid.SelectionMode = "FullRowSelect"
$grid.AutoSizeColumnsMode = "Fill"
$grid.RowHeadersVisible = $false

$columns = @(
    "S.No",
    "IP Address",
    "MAC Address",
    "Brand",
    "Model",
    "Extension Number",
    "SIP Username",
    "URL",
    "Login Status",
    "Web Status",
    "Scan Time"
)

foreach ($c in $columns) {
    [void]$grid.Columns.Add($c, $c)
}

$form.Controls.Add($grid)
$script:grid = $grid

$txtLog = New-Object System.Windows.Forms.TextBox
$txtLog.Location = New-Object System.Drawing.Point(22, 530)
$txtLog.Size = New-Object System.Drawing.Size(1075, 95)
$txtLog.Anchor = "Left,Right,Bottom"
$txtLog.Multiline = $true
$txtLog.ScrollBars = "Vertical"
$txtLog.ReadOnly = $true
$form.Controls.Add($txtLog)
$script:txtLog = $txtLog

$lblNote = New-Object System.Windows.Forms.Label
$lblNote.Text = "📁 Output: D:\fanvil | Brand becomes Fanvil only when Fanvil/X301G/X3SG is confirmed."
$lblNote.Location = New-Object System.Drawing.Point(22, 635)
$lblNote.Size = New-Object System.Drawing.Size(1050, 22)
$lblNote.Anchor = "Left,Right,Bottom"
$form.Controls.Add($lblNote)

Write-Log "Ready. Enter segment and click Start Scan."

try {
    [void]$form.ShowDialog()
}
catch {
    Write-Host "Application error: $($_.Exception.Message)"
    pause
}
