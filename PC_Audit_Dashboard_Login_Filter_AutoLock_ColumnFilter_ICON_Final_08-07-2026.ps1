#Requires -Version 5.1
# ====================================================================================================
# 🖥️ PC AUDIT DASHBOARD - FULLY DOCUMENTED POWERSHELL SCRIPT
# ====================================================================================================
# 👤 Creator / Author : Nirmal kumar
# 🧾 Script Type      : PowerShell HTTP Server Dashboard
# 🎯 Main Purpose     : Read PC audit CSV data and show it in a browser dashboard
# 📁 Output Type      : Local/Intranet web dashboard with summary and details pages
# ⚠️ Important Note   : Original script functionality is kept intact.
#                       Only beginner-friendly comments, section banners, and explanations are added.
#
# ----------------------------------------------------------------------------------------------------
# ✅ What this script does
# ----------------------------------------------------------------------------------------------------
# • 📥 Reads PC audit data from master_audit.csv.
# • 🧹 Converts CSV rows into clean PowerShell objects.
# • 🖧 Identifies network segment based on IP address:
#      - 172.29.x.x              → SFC
#      - 10.209.x.x / 10.208.x.x → Office
# • 🔁 Keeps latest unique PC data for issue cards, and keeps duplicate rows for Daily/Weekly/Monthly/Total.
# • 📊 Calculates Daily, Weekly, Monthly, Total, SFC, and Office counts.
# • 🏆 Finds Top 5 auditors based on audit count.
# • ⚠️ Finds status-based records such as BarTender, Printer, Windows, Symantec, Watermark, and GPO.
# • 🌐 Starts a small PowerShell HTTP server and opens the dashboard in a browser.
# • 📝 Writes logs to C:\Auditdata for troubleshooting.
#
# ----------------------------------------------------------------------------------------------------
# 🧑‍🎓 Beginner Reading Guide
# ----------------------------------------------------------------------------------------------------
# • Lines starting with # are comments. PowerShell ignores them while running.
# • Comments explain WHY each section exists and WHAT each important part does.
# • Do not remove required code lines unless you understand the impact.
# • To change the CSV/server path, check SECTION 01 below.
#
# ----------------------------------------------------------------------------------------------------
# 🧩 Main Sections
# ----------------------------------------------------------------------------------------------------
# 01. Input Parameters and Basic Settings
# 02. Global Variables and Logging Setup
# 03. Write-Log Function
# 04. CSV Import and Data Normalization
# 05. Statistics and Issue Calculation
# 06. HTML Encoding Helper
# 07. Main Dashboard HTML Page
# 08. Details HTML Page
# 09. HTTP Server Engine
# 10. Script Execution and Cleanup
# ====================================================================================================

# ====================================================================================================
# 📌 SECTION 01: INPUT PARAMETERS AND BASIC SETTINGS
# ====================================================================================================
# Purpose:
# • Allows beginner/admin to change dashboard port and CSV base folder without editing code logic.
# • $Port controls the dashboard port number.
# • $BasePath points to the folder where master_audit.csv is stored.
# ----------------------------------------------------------------------------------------------------
param(
   [int]$Port = 8080,
   [string]$BasePath = "\\10.208.193.241\pc_audit$\PCAUDIT-IT"
)

# ====================================================================================================
# EXE CONVERSION FIX - SAFE PATH + SILENT CONSOLE MODE
# ====================================================================================================
# NOTE:
# • PS2EXE / EXE mode can make $PSScriptRoot empty and can show Write-Host text as popup boxes.
# • These safe variables prevent: Cannot bind argument to parameter 'Path' because it is an empty string.
# • Dashboard HTML / CSV calculation logic is not changed.
# ====================================================================================================
$script:ExeSafeScriptRoot = $PSScriptRoot
if ([string]::IsNullOrWhiteSpace($script:ExeSafeScriptRoot)) {
    try {
        $script:ExeSafeScriptRoot = Split-Path -Parent ([System.Diagnostics.Process]::GetCurrentProcess().MainModule.FileName)
    } catch { }
}
if ([string]::IsNullOrWhiteSpace($script:ExeSafeScriptRoot)) {
    $script:ExeSafeScriptRoot = [Environment]::GetFolderPath('ApplicationData')
}
if ([string]::IsNullOrWhiteSpace($BasePath)) {
    # EXE FIX: Keep default CSV base path if converter passes empty BasePath.
    $BasePath = "\\10.208.193.241\pc_audit$\PCAUDIT-IT"
}
$script:ExeSilentConsole = $true

# ====================================================================================================
# GUI MODE: HIDE POWERSHELL CONSOLE EARLY
# ====================================================================================================
# Purpose:
# • The original dashboard wrote status to the PowerShell console.
# • GUI mode hides the console before startup messages and mirrors activity into the WinForms log box.
# • Dashboard HTML / CSV / calculation logic below is not changed.
# ----------------------------------------------------------------------------------------------------
try {
    Add-Type @"
using System;
using System.Runtime.InteropServices;
public class PcAuditEarlyConsole {
    [DllImport("kernel32.dll")] public static extern IntPtr GetConsoleWindow();
    [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
}
"@
    $earlyConsoleHandle = [PcAuditEarlyConsole]::GetConsoleWindow()
    if ($earlyConsoleHandle -ne [IntPtr]::Zero) { [PcAuditEarlyConsole]::ShowWindow($earlyConsoleHandle, 0) | Out-Null }
} catch { }

# ====================================================================================================
# 🛑 ERROR HANDLING MODE
# ====================================================================================================
# Purpose:
# • Stops script immediately when a serious error happens.
# • Prevents dashboard from continuing with incomplete/wrong data.
# ----------------------------------------------------------------------------------------------------
$ErrorActionPreference = "Stop"

# ====================================================================================================
# 🚨 GLOBAL EMERGENCY ERROR HANDLER
# ====================================================================================================
# Purpose:
# • Catches unexpected script errors.
# • Shows the error in red and keeps the window open for reading.
# ----------------------------------------------------------------------------------------------------
trap {
   # EXE FIX: Do not use Write-Host / ReadKey here. In converted EXE mode these can show unwanted popups.
   try { Write-Log -Message "Unhandled error: $_" -Level 'ERROR' } catch { }
   try {
       Add-Type -AssemblyName System.Windows.Forms -ErrorAction SilentlyContinue
       [System.Windows.Forms.MessageBox]::Show("Unexpected error:`n`n$_", "PC Audit Dashboard Error", "OK", "Error") | Out-Null
   } catch { }
   exit 1
}

# ====================================================================================================
# 📌 SECTION 02: GLOBAL VARIABLES AND LOGGING SETUP
# ====================================================================================================
# Purpose:
# • Stores important values used by multiple functions.
# • script: scope means the variable can be accessed anywhere in this script.
# ----------------------------------------------------------------------------------------------------
$script:MasterFile = Join-Path $BasePath "master_audit.csv"
$script:Listener = $null
$script:RunServer = $true

# CACHE FIX - prevents dashboard buffering by avoiding repeated CSV calculation
$script:CachedStats = $null
$script:CachedMasterFileLastWriteTime = $null
$script:CachedStatsDate = $null
$script:CachedWeekStart = $null
$script:CachedMonthStart = $null


# WEB LOGO / FAVICON ICO Base64 code
# Purpose:
# • Browser tab favicon and dashboard title logo use this Base64 code.
# • Keep this file available on the server machine when running PS1/EXE.
$script:DashboardLogoBase64 = @"
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
"@

[string]$script:LogPath = "D:\PC_AUDIT_HOST\PC_AUDIT_LOG\LOGS"
function Get-DailyLogFile {
    return (Join-Path $script:LogPath "audit_dashboard_$(Get-Date -Format 'yyyyMMdd').log")
}

[string]$script:LogFile = Get-DailyLogFile

# ====================================================================================================
# 📌 SECTION 03: WRITE-LOG FUNCTION
# ====================================================================================================
# Function Name : Write-Log
# Purpose       : Save dashboard activity/errors into a daily log file.
# Why used      : Helps troubleshooting without depending only on console messages.
# ----------------------------------------------------------------------------------------------------
function Write-Log {
    param(
        [Parameter(Mandatory=$true)]
        [string]$Message,
        [ValidateSet('INFO', 'WARN', 'ERROR', 'DEBUG')]
        [string]$Level = 'INFO'
    )
    
    try {
        $script:LogFile = Get-DailyLogFile
        if (-not (Test-Path $script:LogPath)) {
            New-Item -ItemType Directory -Path $script:LogPath -Force | Out-Null
        }
        
        $timestamp = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'
        $logEntry = "[$timestamp] [$Level] $Message"
        
        Add-Content -Path $script:LogFile -Value $logEntry -ErrorAction SilentlyContinue
        
        # EXE FIX: Console output disabled because PS2EXE GUI mode can show Write-Host as popup boxes.
        # Original console color output removed; log file + GUI activity box are used instead.
    }
    catch {
        # EXE FIX COMMENTED: Console Write-Host can become popup in converted EXE.
        # Write-Host "Logging failed: $_" -ForegroundColor DarkRed
    }
}

# ====================================================================================================
# 🚀 STARTUP MESSAGES
# ====================================================================================================
# Purpose:
# • Writes dashboard startup information to console and log file.
# • Confirms where the log file is saved.
# ----------------------------------------------------------------------------------------------------
Write-Log -Message "=== PC Audit Dashboard Starting ===" -Level 'INFO'
# EXE FIX COMMENTED: Console Write-Host can become popup in converted EXE.
# Write-Host "PC Audit Dashboard" -ForegroundColor Cyan
# EXE FIX COMMENTED: Console Write-Host can become popup in converted EXE.
# Write-Host "==================" -ForegroundColor Cyan
# EXE FIX COMMENTED: Console Write-Host can become popup in converted EXE.
# Write-Host "Log location: $script:LogFile" -ForegroundColor Gray

# ====================================================================================================
# 📌 SECTION 04: CSV IMPORT AND DATA NORMALIZATION
# ====================================================================================================
# Function Name : Import-AuditData
# Purpose       : Read master_audit.csv and convert each row into a clean object.
# Why used      : Dashboard calculations need consistent property names, even if CSV columns vary slightly.
# Beginner notes:
# • Import-Csv reads rows from CSV.
# • Date parsing converts text dates into DateTime values.
# • Network segment detection uses IP address pattern matching.
# • Duplicate handling keeps latest record per ComputerName + SerialNumber.
# ----------------------------------------------------------------------------------------------------
function Import-AuditData {
   param(
       # ✅ When enabled, this returns ALL CSV rows including duplicate audits.
       # Why: Daily / Weekly / Monthly / Total dashboard cards must show duplicate audit entries also.
       [switch]$KeepDuplicates
   )

   Write-Log -Message "Importing audit data from: $script:MasterFile" -Level 'INFO'
   
   # 📥 Read the CSV file into memory.
   # Why: All dashboard stats and tables are calculated from this CSV data.
   $csvData = Import-Csv $script:MasterFile -ErrorAction Stop
   
   if ($csvData.Count -eq 0) {
       Write-Log -Message "CSV file is empty!" -Level 'ERROR'
       throw "CSV file is empty!"
   }

   # EXE FIX COMMENTED: Console Write-Host can become popup in converted EXE.
   # Write-Host "CSV Columns found:" -ForegroundColor Yellow
   # EXE FIX COMMENTED: CSV column console listing can become many popup boxes after conversion.
   # ($csvData[0].PSObject.Properties.Name) | ForEach-Object { Write-Host "  - $_" }
   # EXE FIX COMMENTED: Console Write-Host can become popup in converted EXE.
   # Write-Host "Total rows: $($csvData.Count)" -ForegroundColor Yellow
   Write-Log -Message "CSV loaded: $($csvData.Count) rows, $(($csvData[0].PSObject.Properties.Name).Count) columns" -Level 'INFO'

   # 🧾 Save original CSV column names.
   # Why: Detail pages display the same columns as master sheet.
   $script:OriginalColumns = $csvData[0].PSObject.Properties.Name

   # 🔍 Use first row to understand available CSV column names.
   # Why: CSV files may have spaces, hyphens, or slightly different column names.
   $firstRow = $csvData[0]
   $colMap = @{}
   foreach ($prop in $firstRow.PSObject.Properties.Name) {
       $colMap[$prop.ToLower().Replace(' ', '').Replace('-', '')] = $prop
   }

   # 👤 Detect auditor/employee ID column automatically.
   # Why: Different CSV versions may use Auditor EMP ID, EmpID, AuditorID, etc.
   $auditorCol = $null
   foreach ($key in @('auditorempid', 'auditorid', 'empid', 'auditor', 'employeeid', 'auditor_emp_id')) {
       if ($colMap.ContainsKey($key)) {
           $auditorCol = $colMap[$key]
           # EXE FIX COMMENTED: Console Write-Host can become popup in converted EXE.
           # Write-Host "Found auditor column: '$auditorCol'" -ForegroundColor Green
           break
       }
   }
   if (-not $auditorCol) {
       # EXE FIX COMMENTED: Console Write-Warning can become popup/noise in converted EXE.
       # Write-Warning "Could not find auditor column! Available: $($colMap.Keys -join ', ')"
       Write-Log -Message "Auditor column not found, using fallback. Available: $($colMap.Keys -join ', ')" -Level 'WARN'
       $auditorCol = 'Auditor EMP ID'
   }

   $dateCol = 'Date'
   foreach ($key in @('date', 'auditdate', 'datetime', 'createdate')) {
       if ($colMap.ContainsKey($key)) {
           $dateCol = $colMap[$key]
           break
       }
   }

   $printerCol = 'Printer Available'
   foreach ($key in @('printeravailable', 'printer', 'printerstatus')) {
       if ($colMap.ContainsKey($key)) {
           $printerCol = $colMap[$key]
           break
       }
   }

   # 📦 This array will store cleaned audit records.
   # Why: Each CSV row is converted into a structured PowerShell object.
   $result = @()
   $rowNum = 0

   # 🔁 Process every CSV row one by one.
   # Why: Each row represents one PC audit entry.
   foreach ($row in $csvData) {
       $rowNum++
       
       $dateVal = $row.$dateCol
       $auditDate = $null
       
       foreach ($fmt in @("MM-dd-yyyy", "M/d/yyyy", "yyyy-MM-dd", "dd-MM-yyyy", "MM/dd/yyyy", "dd/MM/yyyy")) {
           try {
               $auditDate = [DateTime]::ParseExact($dateVal, $fmt, $null)
               break
           } catch { }
       }
       
       if (-not $auditDate) {
           try { $auditDate = [DateTime]::Parse($dateVal) } catch { }
       }

       if ($rowNum -le 3) {
           # EXE FIX COMMENTED: Console Write-Host can become popup in converted EXE.
           # Write-Host "Row $rowNum : Date='$dateVal' -> $($auditDate.ToString('yyyy-MM-dd')) | Auditor='$($row.$auditorCol)'" -ForegroundColor DarkGray
       }

       # 🧱 Create one clean object for dashboard use.
       # Why: This gives dashboard fixed property names even if CSV column names differ.
       $obj = New-Object PSObject -Property @{
           SNo = 0
           Date = $auditDate
           DateString = $dateVal
           ComputerName = ""
           MACAddress = ""
           IPAddress = ""
           SerialNumber = ""
           Brand = ""
           Model = ""
           BarTender = ""
           Printer = ""
           PrinterName = ""
           WindowsActivation = ""
           SymantecStatus = ""
           WatermarkStatus = ""
           GPOStatus = ""
           AuditorID = ""
           NetworkSegment = "Unknown"
           
           _OriginalRow = $row
           _IsDuplicate = $false
       }

       try { $obj.SNo = [int]$row.SNo } catch { }
       if ($row.'Computer Name') { $obj.ComputerName = $row.'Computer Name' }
       if ($row.'MAC Address') { $obj.MACAddress = $row.'MAC Address' }
       # 🖧 Read IP address and identify network segment.
       # Why: Dashboard separates SFC and Office records using IP ranges.
       if ($row.'IP Address') { 
           $obj.IPAddress = $row.'IP Address'
           $ip = $row.'IP Address'
           if ($ip -match '^172\.29\.') {
               $obj.NetworkSegment = 'SFC'
           }
           elseif ($ip -match '^10\.(209|208)\.') {
               $obj.NetworkSegment = 'Office'
           }
       }
       if ($row.'Serial Number') { $obj.SerialNumber = $row.'Serial Number' }
       if ($row.Brand) { $obj.Brand = $row.Brand }
       if ($row.Model) { $obj.Model = $row.Model }
       if ($row.'BarTender Available') { $obj.BarTender = $row.'BarTender Available' }
       if ($row.$printerCol) { $obj.Printer = $row.$printerCol }
       if ($row.'Printer Name') { $obj.PrinterName = $row.'Printer Name' }
       if ($row.'Windows Activation') { $obj.WindowsActivation = $row.'Windows Activation' }
       if ($row.'Symantec Status') { $obj.SymantecStatus = $row.'Symantec Status' }
       if ($row.'Watermark Status') { $obj.WatermarkStatus = $row.'Watermark Status' }
       if ($row.'GPO Update Status') { $obj.GPOStatus = $row.'GPO Update Status' }
       
       $auditorVal = $row.$auditorCol
       if ($auditorVal) { 
           $obj.AuditorID = $auditorVal.ToString().Trim()
       }

       $result += $obj
   }

   # ====================================================================================================
   # 🔁 DUPLICATE HANDLING - LATEST UNIQUE DATA + OPTIONAL FULL DUPLICATE DATA
   # ====================================================================================================
   # Purpose:
   # • Groups records by ComputerName + SerialNumber.
   # • Marks older repeated audit rows as DUPLICATE.
   # • Default return keeps only newest audit entry per PC for issue cards.
   # • -KeepDuplicates return keeps ALL rows for Daily / Weekly / Monthly / Total cards and details.
   # Important:
   # • This protects the 6 issue cards from duplicate inflation.
   # • Period cards can now show duplicate audit rows exactly as entered in master_audit.csv.
   # ----------------------------------------------------------------------------------------------------
   $grouped = $result | Where-Object { $_.ComputerName -and $_.SerialNumber } | 
                    Group-Object { "$($_.ComputerName)|$($_.SerialNumber)" }

   $uniqueResult = @()

   foreach ($group in $grouped) {
       # Sort group by Date descending. Newest record is treated as current/latest PC status.
       $sortedGroup = @($group.Group | Sort-Object { if ($_.Date) { $_.Date } else { [DateTime]::MinValue } } -Descending)

       if ($sortedGroup.Count -gt 0) {
           # Latest row stays normal.
           $sortedGroup[0]._IsDuplicate = $false
           $uniqueResult += $sortedGroup[0]

           # Older rows are marked as duplicate for period detail pages.
           if ($sortedGroup.Count -gt 1) {
               for ($i = 1; $i -lt $sortedGroup.Count; $i++) {
                   $sortedGroup[$i]._IsDuplicate = $true
               }
           }
       }
   }

   # Handle items without ComputerName or SerialNumber.
   # Why: They cannot be safely grouped, so they are kept as individual rows.
   $incompleteItems = @($result | Where-Object { -not $_.ComputerName -or -not $_.SerialNumber })
   foreach ($item in $incompleteItems) {
       $item._IsDuplicate = $false
   }

   # ✅ FULL DATA MODE:
   # Used only for Daily / Weekly / Monthly / Total cards and period detail pages.
   if ($KeepDuplicates) {
       Write-Log -Message "Duplicate handling: Returning ALL $($result.Count) records for period counts/details. Duplicate rows are marked." -Level 'INFO'
       return $result
   }

   # ✅ UNIQUE DATA MODE:
   # Used for 6 issue cards so existing issue-card logic/counts are NOT affected by duplicate audits.
   $finalResult = $uniqueResult + $incompleteItems

   Write-Log -Message "Duplicate handling: Original $($result.Count) records → Unique $($finalResult.Count) records after keeping latest per PC" -Level 'INFO'

   return $finalResult
}


# ====================================================================================================
# 📌 SECTION 05: STATISTICS AND ISSUE CALCULATION
# ====================================================================================================
# Function Name : Get-Stats
# Purpose       : Calculate dashboard summary counts and issue lists.
# Why used      : Main page and details page both need counts and filtered records.
# Beginner notes:
# • Daily = records from today.
# • Weekly = current week start to today.
# • Monthly = first day of current month to today.
# • SFC and Office are calculated from NetworkSegment.
# ----------------------------------------------------------------------------------------------------
function Get-Stats {
   # ✅ Period data keeps all CSV rows including duplicate audits.
   # Used by Daily / Weekly / Monthly / Total cards and their detail pages.
   # CACHE FIX - reload CSV only if master_audit.csv changed
   $currentLastWriteTime = (Get-Item $script:MasterFile).LastWriteTime

   $currentLastWriteTime = (Get-Item $script:MasterFile).LastWriteTime

    $currentStatsDate = (Get-Date).Date
    $currentWeekStart = $currentStatsDate.AddDays(-([int]$currentStatsDate.DayOfWeek))
    $currentMonthStart = $currentStatsDate.AddDays(-($currentStatsDate.Day - 1))

if ($script:CachedStats -ne $null -and 
    $script:CachedMasterFileLastWriteTime -eq $currentLastWriteTime -and
    $script:CachedStatsDate -eq $currentStatsDate -and
    $script:CachedWeekStart -eq $currentWeekStart -and
    $script:CachedMonthStart -eq $currentMonthStart) {
    return $script:CachedStats
}

   # ✅ Issue data keeps latest unique PC only.
   # Used by the 6 issue cards so existing issue-card counts are not affected. 
   # FAST MODE - read CSV only once
$allData = Import-AuditData -KeepDuplicates

# Create latest unique data from same imported data
$grouped = $allData | Where-Object { $_.ComputerName -and $_.SerialNumber } |
    Group-Object { "$($_.ComputerName)|$($_.SerialNumber)" }

$data = @()

foreach ($group in $grouped) {
    $latest = $group.Group | Sort-Object { 
        if ($_.Date) { $_.Date } else { [DateTime]::MinValue } 
    } -Descending | Select-Object -First 1

    if ($latest) {
        $data += $latest
    }
}

# Keep incomplete rows also
$incompleteItems = @($allData | Where-Object { -not $_.ComputerName -or -not $_.SerialNumber })
$data += $incompleteItems

   
   # EXE FIX COMMENTED: Console Write-Host can become popup in converted EXE.
   # Write-Host "`nCalculating statistics..." -ForegroundColor Yellow
   Write-Log -Message "Calculating statistics..." -Level 'INFO'

   # 📅 Calculate date boundaries used by dashboard cards.
   # Why: These dates decide Daily, Weekly, and Monthly records.
   $today = (Get-Date).Date
   $weekStart = $today.AddDays(-([int]$today.DayOfWeek))
   $monthStart = $today.AddDays(-($today.Day - 1))

   # EXE FIX COMMENTED: Console Write-Host can become popup in converted EXE.
   # Write-Host "Date range: Today=$today, WeekStart=$weekStart, MonthStart=$monthStart" -ForegroundColor DarkGray

   # ====================================================================================================
   # 🔢 COUNT RECORDS BY PERIOD AND NETWORK SEGMENT
   # ====================================================================================================
   # Purpose:
   # • Counts Daily, Weekly, Monthly totals.
   # • Separately counts SFC and Office records.
   # ----------------------------------------------------------------------------------------------------
   # Calculate counts with network segment breakdown
   $dailyTotal = 0; $dailySFC = 0; $dailyOffice = 0
   $weeklyTotal = 0; $weeklySFC = 0; $weeklyOffice = 0
   $monthlyTotal = 0; $monthlySFC = 0; $monthlyOffice = 0

   foreach ($item in $allData) {
       if ($item.Date) {
           $itemDate = $item.Date.Date
           
           # Daily
           if ($itemDate -eq $today) { 
               $dailyTotal++
               if ($item.NetworkSegment -eq 'SFC') { $dailySFC++ }
               elseif ($item.NetworkSegment -eq 'Office') { $dailyOffice++ }
           }
           
           # Weekly
           if ($itemDate -ge $weekStart) { 
               $weeklyTotal++
               if ($item.NetworkSegment -eq 'SFC') { $weeklySFC++ }
               elseif ($item.NetworkSegment -eq 'Office') { $weeklyOffice++ }
           }
           
           # Monthly
           if ($itemDate -ge $monthStart) { 
               $monthlyTotal++
               if ($item.NetworkSegment -eq 'SFC') { $monthlySFC++ }
               elseif ($item.NetworkSegment -eq 'Office') { $monthlyOffice++ }
           }
       }
   }

   # FIXED: Total counts by segment - Calculate from actual data, not cached values
   $totalSFC = 0
   $totalOffice = 0
   
   foreach ($item in $allData) {
       if ($item.NetworkSegment -eq 'SFC') { $totalSFC++ }
       elseif ($item.NetworkSegment -eq 'Office') { $totalOffice++ }
   }

   # EXE FIX COMMENTED: Console Write-Host can become popup in converted EXE.
   # Write-Host "Counts: Daily=$dailyTotal (SFC:$dailySFC, Office:$dailyOffice), Weekly=$weeklyTotal (SFC:$weeklySFC, Office:$weeklyOffice), Monthly=$monthlyTotal (SFC:$monthlySFC, Office:$monthlyOffice), Total=$($allData.Count) (SFC:$totalSFC, Office:$totalOffice)" -ForegroundColor Green
   Write-Log -Message "Stats calculated with segment breakdown - Total SFC: $totalSFC, Total Office: $totalOffice" -Level 'INFO'

   # 🏆 Prepare auditor data for Top 5 table.
   # Why: Empty auditor IDs are removed so ranking is accurate.
   $auditorData = @($data | Where-Object { 
       $_.AuditorID -and $_.AuditorID -ne "" -and $_.AuditorID -ne " " 
   })

   $recentAuditorData = @($auditorData | Where-Object { $_.Date -and $_.Date -ge $weekStart })
   
   if ($recentAuditorData.Count -gt 0) {
       # EXE FIX COMMENTED: Console Write-Host can become popup in converted EXE.
       # Write-Host "Using weekly auditor data: $($recentAuditorData.Count) records" -ForegroundColor Green
       $auditorSource = $recentAuditorData
   } else {
       # EXE FIX COMMENTED: Console Write-Host can become popup in converted EXE.
       # Write-Host "No weekly data, using all auditor data: $($auditorData.Count) records" -ForegroundColor Yellow
       Write-Log -Message "No weekly auditor data, falling back to all data" -Level 'WARN'
       $auditorSource = $auditorData
   }

   $topAuditors = @()
   if ($auditorSource.Count -gt 0) {
       $grouped = $auditorSource | Group-Object -Property { $_.AuditorID.ToString().Trim() }
       # EXE FIX COMMENTED: Console Write-Host can become popup in converted EXE.
       # Write-Host "Grouped into $($grouped.Count) unique auditors" -ForegroundColor Green
       
       $topAuditors = $grouped | 
           Select-Object @{N='AuditorID';E={$_.Name}}, @{N='Count';E={$_.Count}} |
           Sort-Object Count -Descending |
           Select-Object -First 5

       $topAuditors | ForEach-Object {
           # EXE FIX COMMENTED: Console Write-Host can become popup in converted EXE.
           # Write-Host "  TOP: $($_.AuditorID) = $($_.Count)" -ForegroundColor Cyan
       }
       Write-Log -Message "Top auditors: $($topAuditors.Count) found" -Level 'INFO'
   } else {
       # EXE FIX COMMENTED: Console Write-Warning can become popup/noise in converted EXE.
       # Write-Warning "No auditor data found!"
       Write-Log -Message "No auditor data found!" -Level 'WARN'
   }

   $output = @{
       Data = $allData
       Daily = @{ Total = $dailyTotal; SFC = $dailySFC; Office = $dailyOffice }
       Weekly = @{ Total = $weeklyTotal; SFC = $weeklySFC; Office = $weeklyOffice }
       Monthly = @{ Total = $monthlyTotal; SFC = $monthlySFC; Office = $monthlyOffice }
       Total = @{ Total = $allData.Count; SFC = $totalSFC; Office = $totalOffice }
       TopAuditors = $topAuditors
       Today = $today
       WeekStart = $weekStart
       MonthStart = $monthStart
       OriginalColumns = $script:OriginalColumns
   }

   # ⚠️ Build issue/status collections used by issue cards and detail pages.
   # Why: Each issue stores All, SFC, and Office filtered record lists.
   $issues = @{}
   
   $issues['BarTender'] = @{
       All = @($data | Where-Object { $_.BarTender -notmatch '^(NO|No|no|N)$' })
       SFC = @($data | Where-Object { $_.BarTender -notmatch '^(NO|No|no|N)$' -and $_.NetworkSegment -eq 'SFC' })
       Office = @($data | Where-Object { $_.BarTender -notmatch '^(NO|No|no|N)$' -and $_.NetworkSegment -eq 'Office' })
   }
   
   $issues['PrinterAvailable'] = @{
       All = @($data | Where-Object { $_.Printer -and $_.Printer -notmatch '^(NO|No|no|N)$' })
       SFC = @($data | Where-Object { $_.Printer -and $_.Printer -notmatch '^(NO|No|no|N)$' -and $_.NetworkSegment -eq 'SFC' })
       Office = @($data | Where-Object { $_.Printer -and $_.Printer -notmatch '^(NO|No|no|N)$' -and $_.NetworkSegment -eq 'Office' })
   }
   
   $issues['WindowsActivation'] = @{
       All = @($data | Where-Object { $_.WindowsActivation -notmatch '^(YES|Yes|yes|Y)$' })
       SFC = @($data | Where-Object { $_.WindowsActivation -notmatch '^(YES|Yes|yes|Y)$' -and $_.NetworkSegment -eq 'SFC' })
       Office = @($data | Where-Object { $_.WindowsActivation -notmatch '^(YES|Yes|yes|Y)$' -and $_.NetworkSegment -eq 'Office' })
   }
   
   $issues['Symantec'] = @{
       All = @($data | Where-Object { $_.SymantecStatus -notmatch '^(ACTIVE|Active|active)$' })
       SFC = @($data | Where-Object { $_.SymantecStatus -notmatch '^(ACTIVE|Active|active)$' -and $_.NetworkSegment -eq 'SFC' })
       Office = @($data | Where-Object { $_.SymantecStatus -notmatch '^(ACTIVE|Active|active)$' -and $_.NetworkSegment -eq 'Office' })
   }
   
   $issues['Watermark'] = @{
       All = @($data | Where-Object { $_.WatermarkStatus -notmatch '^(YES|Yes|yes|Y)$' })
       SFC = @($data | Where-Object { $_.WatermarkStatus -notmatch '^(YES|Yes|yes|Y)$' -and $_.NetworkSegment -eq 'SFC' })
       Office = @($data | Where-Object { $_.WatermarkStatus -notmatch '^(YES|Yes|yes|Y)$' -and $_.NetworkSegment -eq 'Office' })
   }
   
   $issues['GPO'] = @{
       All = @($data | Where-Object { $_.GPOStatus -match 'NOTUPDATED' -or $_.GPOStatus -notmatch 'UPDATED' -or $_.GPOStatus -eq '()' -or $_.GPOStatus -eq '( )' })
       SFC = @($data | Where-Object { ($_.GPOStatus -match 'NOTUPDATED' -or $_.GPOStatus -notmatch 'UPDATED' -or $_.GPOStatus -eq '()' -or $_.GPOStatus -eq '( )') -and $_.NetworkSegment -eq 'SFC' })
       Office = @($data | Where-Object { ($_.GPOStatus -match 'NOTUPDATED' -or $_.GPOStatus -notmatch 'UPDATED' -or $_.GPOStatus -eq '()' -or $_.GPOStatus -eq '( )') -and $_.NetworkSegment -eq 'Office' })
   }

   $output['Issues'] = $issues
   
      Write-Log -Message "Issues detected with segment breakdown" -Level 'INFO'

   # CACHE FIX - save latest calculated dashboard data
    $script:CachedStats = $output
    $script:CachedMasterFileLastWriteTime = $currentLastWriteTime
    $script:CachedStatsDate = $currentStatsDate
    $script:CachedWeekStart = $currentWeekStart
    $script:CachedMonthStart = $currentMonthStart

    return $output
}

# ====================================================================================================
# 📌 SECTION 06: HTML ENCODING HELPER
# ====================================================================================================
# Function Name : Get-HtmlEncoded
# Purpose       : Safely display CSV text inside HTML pages.
# Why used      : Prevents special characters like <, >, &, and quotes from breaking HTML layout.
# ----------------------------------------------------------------------------------------------------
function Get-HtmlEncoded {
   param([string]$Text)
   if ([string]::IsNullOrEmpty($Text)) { return "" }
   
   $encoded = $Text -replace '&', '&amp;'
   $encoded = $encoded -replace '<', '&lt;'
   $encoded = $encoded -replace '>', '&gt;'
   $encoded = $encoded -replace '"', '&quot;'
   
   return $encoded
}

# ====================================================================================================
# 📌 SECTION 07: MAIN DASHBOARD HTML PAGE
# ====================================================================================================
# Function Name : Get-MainPage
# Purpose       : Generate the main dashboard web page.
# Why used      : This returns HTML/CSS/JavaScript that browser can display.
# Beginner notes:
# • PowerShell creates dynamic HTML using values from Get-Stats.
# • CSS controls dashboard design.
# • JavaScript openDetail() opens detailed records in popup window.
# ----------------------------------------------------------------------------------------------------
function Get-MainPage {
   $stats = Get-Stats

   # 🏆 Build HTML table rows for Top 5 Auditors.
   # Why: This converts PowerShell data into HTML <tr> rows.
   $auditorRows = ""
   if ($stats.TopAuditors -and $stats.TopAuditors.Count -gt 0) {
       foreach ($aud in $stats.TopAuditors) {
           $name = Get-HtmlEncoded $aud.AuditorID
           $auditorRows = $auditorRows + "<tr><td>$name</td><td>$($aud.Count)</td></tr>`n"
       }
   } else {
       $auditorRows = "<tr><td colspan='2' style='color:#888;'>No data this week (showing all-time if available)</td></tr>"
   }

   # ====================================================================================================
   # ⚠️ ISSUE CARD COUNTS
   # ====================================================================================================
   # Purpose:
   # • Calculates values shown in colored issue cards.
   # • Each card has total count and SFC/Office split.
   # ----------------------------------------------------------------------------------------------------
   # Issue card counts
   $btCountSFC = $stats.Issues['BarTender'].SFC.Count
   $btCountOffice = $stats.Issues['BarTender'].Office.Count
   $btCountTotal = $btCountSFC + $btCountOffice
   
   $prCountSFC = $stats.Issues['PrinterAvailable'].SFC.Count
   $prCountOffice = $stats.Issues['PrinterAvailable'].Office.Count
   $prCountTotal = $prCountSFC + $prCountOffice
   
   $waCountSFC = $stats.Issues['WindowsActivation'].SFC.Count
   $waCountOffice = $stats.Issues['WindowsActivation'].Office.Count
   $waCountTotal = $waCountSFC + $waCountOffice
   
   $syCountSFC = $stats.Issues['Symantec'].SFC.Count
   $syCountOffice = $stats.Issues['Symantec'].Office.Count
   $syCountTotal = $syCountSFC + $syCountOffice
   
   $wmCountSFC = $stats.Issues['Watermark'].SFC.Count
   $wmCountOffice = $stats.Issues['Watermark'].Office.Count
   $wmCountTotal = $wmCountSFC + $wmCountOffice
   
   $gpCountSFC = $stats.Issues['GPO'].SFC.Count
   $gpCountOffice = $stats.Issues['GPO'].Office.Count
   $gpCountTotal = $gpCountSFC + $gpCountOffice

   # Period counts with segment breakdown
   $dailyTotal = $stats.Daily.Total
   $dailySFC = $stats.Daily.SFC
   $dailyOffice = $stats.Daily.Office
   
   $weeklyTotal = $stats.Weekly.Total
   $weeklySFC = $stats.Weekly.SFC
   $weeklyOffice = $stats.Weekly.Office
   
   $monthlyTotal = $stats.Monthly.Total
   $monthlySFC = $stats.Monthly.SFC
   $monthlyOffice = $stats.Monthly.Office
   
   $totalAll = $stats.Total.Total
   $totalSFC = $stats.Total.SFC
   $totalOffice = $stats.Total.Office

   # ====================================================================================================
   # 🔔 NOTIFICATION COUNT SNAPSHOT
   # ====================================================================================================
   # Purpose:
   # • Sends current dashboard counts to browser JavaScript.
   # • Browser compares these counts with previous refresh values saved in localStorage.
   # • If any issue card count changes, popup notification updates and browser beep plays.
   # • If total audited PC count increases, dashboard shows a new PC audited notification.
   # Important:
   # • This does not change existing PowerShell calculation logic.
   # • Notification comparison happens only in browser side JavaScript.
   # ----------------------------------------------------------------------------------------------------
   $notificationGeneratedAt = Get-Date -Format 'yyyy-MM-dd HH:mm:ss'

   # 🌐 Start of HTML here-string.
   # Why: A here-string allows large multi-line HTML content inside PowerShell.
   $html = @"
<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<title>PC Audit Dashboard</title>
<link rel="icon" type="image/png" href="data:image/png;base64,$($script:DashboardLogoBase64)">
<link rel="shortcut icon" type="image/png" href="data:image/png;base64,$($script:DashboardLogoBase64)">
<meta http-equiv="refresh" content="60">

<style>
* { margin: 0; padding: 0; box-sizing: border-box; }

body { 
    font-family: 'Segoe UI', Arial, sans-serif; 
    background: #1a1d21;
    color: #fff;
    padding: 16px; 
    min-height: 100vh; 
}

h1 { 
    text-align: center; 
    margin-bottom: 9px; 
    font-size: 2.8em; 
}

.timestamp { 
    text-align: center; 
    color: #888;
    margin-bottom: 22px; 
    font-size: 0.98em;
}

.stats-container { 
    display: flex; 
    justify-content: center; 
    gap: 16px; 
    margin-bottom: 27px; 
    flex-wrap: wrap; 
}

.stat-box { 
    width: 215px; 
    padding: 18px 14px; 
    border-radius: 11px; 
    text-align: center; 
}

.stat-box.clickable { 
    cursor: pointer; 
    transition: transform 0.2s;
}

.stat-box.clickable:hover { 
    transform: translateY(-4px);
    opacity: 0.9; 
}

.stat-box.daily { background: #2ecc71; }
.stat-box.weekly { background: #3498db; }
.stat-box.monthly { background: #9b59b6; }
.stat-box.total { background: #e67e22; }

.stat-label { 
    font-size: 12.5px; 
    text-transform: uppercase; 
    opacity: 0.9; 
    margin-bottom: 7px;
    letter-spacing: 0.5px;
}

.stat-value { 
    font-size: 37px; 
    font-weight: bold; 
    margin-bottom: 13px;
    line-height: 1;
}

.dual-buttons {
    display: flex;
    gap: 7px;
    justify-content: center;
}

.segment-btn {
    flex: 1;
    padding: 8px 10px;
    border: none;
    border-radius: 5.5px;
    cursor: pointer;
    font-size: 10.5px;
    font-weight: 600;
    transition: all 0.2s;
    text-transform: uppercase;
    line-height: 1.25;
}

.segment-btn.sfc {
    background: rgba(0,0,0,0.3);
    color: #fff;
}

.segment-btn.sfc:hover {
    background: rgba(0,0,0,0.5);
    transform: scale(1.05);
}

.segment-btn.office {
    background: rgba(255,255,255,0.9);
    color: #333;
}

.segment-btn.office:hover {
    background: #fff;
    transform: scale(1.05);
}

.segment-btn .count {
    display: block;
    font-size: 18px;
    font-weight: bold;
    margin-top: 2.5px;
}

.content-grid { 
    display: grid; 
    grid-template-columns: 1fr 1fr;
    gap: 22px; 
    max-width: 1350px; 
    margin: 0 auto; 
}

.section { 
    background: #2d3139;
    border-radius: 11px; 
    padding: 20px; 
}

.section h2 { 
    margin-bottom: 16px; 
    font-size: 1.22em; 
}

table { 
    width: 100%; 
    border-collapse: collapse; 
    font-size: 0.95em;
}

th, td { 
    padding: 11px 13px; 
    text-align: left; 
    border-bottom: 1px solid #444; 
}

th { 
    background: #363b44; 
    font-weight: 600; 
}

tr:hover { 
    background: #363b44;
}

.issues-grid { 
    display: grid; 
    grid-template-columns: 1fr 1fr;
    gap: 13px; 
}

.issue-card { 
    padding: 14px; 
    border-radius: 9px; 
    transition: transform 0.2s; 
}

.issue-card:hover { 
    transform: translateY(-3px); 
    opacity: 0.9; 
}

.issue-card:nth-child(1) { background: #e74c3c; }
.issue-card:nth-child(2) { background: #27ae60; }
.issue-card:nth-child(3) { background: #f1c40f; color: #000; }
.issue-card:nth-child(4) { background: #9b59b6; }
.issue-card:nth-child(5) { background: #3498db; }
.issue-card:nth-child(6) { background: #95a5a6; }

.issue-title { 
    font-size: 12px; 
    margin-bottom: 9px; 
    text-align: center;
    line-height: 1.3;
}

.issue-count { 
    font-size: 25px; 
    font-weight: bold; 
    text-align: center;
    margin-bottom: 11px;
    line-height: 1;
}

.status-bar { 
    position: fixed; 
    bottom: 0; 
    left: 0; 
    right: 0; 
    background: #2d3139; 
    padding: 9px 18px; 
    text-align: center; 
    color: #888; 
    font-size: 11.5px; 
}

/* ====================================================================================================
   🔔 NOTIFICATION UI - NEW ADDITION
   Purpose:
   • Shows an enable button so browser notification and sound can work.
   • Shows small popup toast inside dashboard when issue counts or PC audit counts change.
   • Does not change existing dashboard calculations or existing functions.
   ==================================================================================================== */
.notification-enable {
    position: fixed;
    top: 14px;
    right: 14px;
    z-index: 9999;
    padding: 10px 14px;
    border: none;
    border-radius: 8px;
    background: #f39c12;
    color: #111;
    font-weight: 700;
    cursor: pointer;
    box-shadow: 0 4px 14px rgba(0,0,0,0.35);
}

.notification-toast {
    position: fixed;
    top: 62px;
    right: 14px;
    max-width: 360px;
    z-index: 9999;
    padding: 13px 16px;
    border-radius: 9px;
    background: #111827;
    color: #fff;
    border-left: 5px solid #f39c12;
    box-shadow: 0 6px 18px rgba(0,0,0,0.45);
    display: none;
    line-height: 1.45;
    font-size: 13px;
}

@media (max-width: 900px) { 
    .content-grid { grid-template-columns: 1fr; }
    .issues-grid { grid-template-columns: 1fr; }
    .stats-container { flex-direction: column; align-items: center; }
    .stat-box { width: 100%; max-width: 300px; }
    h1 { font-size: 1.9em; }
}
</style>
</head>
<body>

<!-- 🔔 NEW: Notification enable button and toast popup area -->
<div id="auditNotificationToast" class="notification-toast"></div>

<h1 style="display:flex;align-items:center;justify-content:center;gap:12px;">
    <img src="data:image/png;base64,$($script:DashboardLogoBase64)"
         alt="PC Audit Logo"
         style="width:48px;height:48px;border-radius:8px;object-fit:contain;">
    <span>PC AUDIT DASHBOARD</span>
</h1>

<div class="timestamp">
    Last updated: $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') | 
    Total Records: $($totalAll)
</div>

<div class="stats-container">
    <div class="stat-box daily clickable" onclick="openDetail('Daily_All')">
        <div class="stat-label">Daily</div>
        <div class="stat-value">$dailyTotal</div>
        <div class="dual-buttons">
            <button class="segment-btn sfc" onclick="event.stopPropagation(); openDetail('Daily_SFC')">
                SFC<span class="count">$dailySFC</span>
            </button>
            <button class="segment-btn office" onclick="event.stopPropagation(); openDetail('Daily_Office')">
                Office<span class="count">$dailyOffice</span>
            </button>
        </div>
    </div>
    <div class="stat-box weekly clickable" onclick="openDetail('Weekly_All')">
        <div class="stat-label">Weekly</div>
        <div class="stat-value">$weeklyTotal</div>
        <div class="dual-buttons">
            <button class="segment-btn sfc" onclick="event.stopPropagation(); openDetail('Weekly_SFC')">
                SFC<span class="count">$weeklySFC</span>
            </button>
            <button class="segment-btn office" onclick="event.stopPropagation(); openDetail('Weekly_Office')">
                Office<span class="count">$weeklyOffice</span>
            </button>
        </div>
    </div>
    <div class="stat-box monthly clickable" onclick="openDetail('Monthly_All')">
        <div class="stat-label">Monthly</div>
        <div class="stat-value">$monthlyTotal</div>
        <div class="dual-buttons">
            <button class="segment-btn sfc" onclick="event.stopPropagation(); openDetail('Monthly_SFC')">
                SFC<span class="count">$monthlySFC</span>
            </button>
            <button class="segment-btn office" onclick="event.stopPropagation(); openDetail('Monthly_Office')">
                Office<span class="count">$monthlyOffice</span>
            </button>
        </div>
    </div>
    <div class="stat-box total clickable" onclick="openDetail('Total_All')">
        <div class="stat-label">Total</div>
        <div class="stat-value">$totalAll</div>
        <div class="dual-buttons">
            <button class="segment-btn sfc" onclick="event.stopPropagation(); openDetail('Total_SFC')">
                SFC<span class="count">$totalSFC</span>
            </button>
            <button class="segment-btn office" onclick="event.stopPropagation(); openDetail('Total_Office')">
                Office<span class="count">$totalOffice</span>
            </button>
        </div>
    </div>
</div>

<div class="content-grid">

    <div class="section">
        <h2>🏆 Top 5 Auditors</h2>
        <table>
            <thead>
                <tr>
                    <th>Auditor ID</th>
                    <th>Audit Count</th>
                </tr>
            </thead>
            <tbody>$auditorRows</tbody>
        </table>
    </div>

    <div class="section">
        <h2>📊 System Status (Click for details)</h2>
        <div class="issues-grid">
            
            <div class="issue-card">
                <div class="issue-title">✅ BarTender Available PCs</div>
                <div class="issue-count">$btCountTotal</div>
                <div class="dual-buttons">
                    <button class="segment-btn sfc" onclick="openDetail('BarTender_SFC')">
                        SFC<span class="count">$btCountSFC</span>
                    </button>
                    <button class="segment-btn office" onclick="openDetail('BarTender_Office')">
                        Office<span class="count">$btCountOffice</span>
                    </button>
                </div>
            </div>
            
            <div class="issue-card">
                <div class="issue-title">🖨️ Printer Available PCs</div>
                <div class="issue-count">$prCountTotal</div>
                <div class="dual-buttons">
                    <button class="segment-btn sfc" onclick="openDetail('PrinterAvailable_SFC')">
                        SFC<span class="count">$prCountSFC</span>
                    </button>
                    <button class="segment-btn office" onclick="openDetail('PrinterAvailable_Office')">
                        Office<span class="count">$prCountOffice</span>
                    </button>
                </div>
            </div>
            
            <div class="issue-card">
                <div class="issue-title">⚠️ Windows Not Activated</div>
                <div class="issue-count">$waCountTotal</div>
                <div class="dual-buttons">
                    <button class="segment-btn sfc" onclick="openDetail('WindowsActivation_SFC')">
                        SFC<span class="count">$waCountSFC</span>
                    </button>
                    <button class="segment-btn office" onclick="openDetail('WindowsActivation_Office')">
                        Office<span class="count">$waCountOffice</span>
                    </button>
                </div>
            </div>
            
            <div class="issue-card">
                <div class="issue-title">🛡️ Symantec Inactive</div>
                <div class="issue-count">$syCountTotal</div>
                <div class="dual-buttons">
                    <button class="segment-btn sfc" onclick="openDetail('Symantec_SFC')">
                        SFC<span class="count">$syCountSFC</span>
                    </button>
                    <button class="segment-btn office" onclick="openDetail('Symantec_Office')">
                        Office<span class="count">$syCountOffice</span>
                    </button>
                </div>
            </div>
            
            <div class="issue-card">
                <div class="issue-title">✒️ Watermark Not Available</div>
                <div class="issue-count">$wmCountTotal</div>
                <div class="dual-buttons">
                    <button class="segment-btn sfc" onclick="openDetail('Watermark_SFC')">
                        SFC<span class="count">$wmCountSFC</span>
                    </button>
                    <button class="segment-btn office" onclick="openDetail('Watermark_Office')">
                        Office<span class="count">$wmCountOffice</span>
                    </button>
                </div>
            </div>
            
            <div class="issue-card">
                <div class="issue-title">⚙️ GPO Not Updated</div>
                <div class="issue-count">$gpCountTotal</div>
                <div class="dual-buttons">
                    <button class="segment-btn sfc" onclick="openDetail('GPO_SFC')">
                        SFC<span class="count">$gpCountSFC</span>
                    </button>
                    <button class="segment-btn office" onclick="openDetail('GPO_Office')">
                        Office<span class="count">$gpCountOffice</span>
                    </button>
                </div>
            </div>
            
        </div>
    </div>
</div>

<div class="status-bar">
    Dashboard: http://pc_Audit_Dashboard:$Port | 
    Data: $($totalAll) records | SFC: $totalSFC | Office: $totalOffice
</div>

<script>
function openDetail(type) { 
    window.open('/details?type=' + type, '_blank', 
                'width=1200,height=700,scrollbars=yes'); 
}

/* ====================================================================================================
   🔔 NEW: DASHBOARD NOTIFICATION + SOUND FUNCTION
   Purpose:
   • Compares current dashboard counts with the previous browser refresh.
   • Notifies when any issue card count changes.
   • Notifies when new PC audit count is detected.
   • Plays browser-generated beep sound with every notification.
   Important:
   • No external audio file is required.
   • User must click "Enable Notifications" once in the browser.
   ==================================================================================================== */
const AUDIT_NOTIFICATION_STORAGE_KEY = 'pcAuditDashboardLastSnapshot';
const AUDIT_NOTIFICATION_ENABLED_KEY = 'pcAuditDashboardNotificationEnabled';

const currentAuditSnapshot = {
    totalAll: $totalAll,
    dailyTotal: $dailyTotal,
    issues: {
        BarTender: $btCountTotal,
        PrinterAvailable: $prCountTotal,
        WindowsActivation: $waCountTotal,
        Symantec: $syCountTotal,
        Watermark: $wmCountTotal,
        GPO: $gpCountTotal
    }
};

function enableAuditNotifications() {
    localStorage.setItem(AUDIT_NOTIFICATION_ENABLED_KEY, 'YES');

    if ('Notification' in window && Notification.permission === 'default') {
        Notification.requestPermission().then(function () {
            playAuditBeep();
            showAuditToast('Notifications enabled. Sound test completed.');
        });
    } else {
        playAuditBeep();
        showAuditToast('Notifications enabled. Sound test completed.');
    }
}

function playAuditBeep() {
    try {
        const AudioContextClass = window.AudioContext || window.webkitAudioContext;
        if (!AudioContextClass) return;
        const ctx = new AudioContextClass();
        const now = ctx.currentTime;
        
        const o1 = ctx.createOscillator(), g1 = ctx.createGain();
        const o2 = ctx.createOscillator(), g2 = ctx.createGain();
        
        o1.type = 'sine';
        o1.frequency.setValueAtTime(659.25, now);  // E5 - "ding"
        
        o2.type = 'sine';
        o2.frequency.setValueAtTime(523.25, now + 0.15);  // C5 - "dong" (delayed)
        
        g1.gain.setValueAtTime(0, now);
        g1.gain.linearRampToValueAtTime(0.4, now + 0.02);
        g1.gain.exponentialRampToValueAtTime(0.01, now + 0.4);
        
        g2.gain.setValueAtTime(0, now + 0.15);
        g2.gain.linearRampToValueAtTime(0.4, now + 0.17);
        g2.gain.exponentialRampToValueAtTime(0.01, now + 0.8);
        
        o1.connect(g1); o2.connect(g2);
        g1.connect(ctx.destination); g2.connect(ctx.destination);
        
        o1.start(now); o2.start(now + 0.15);
        o1.stop(now + 0.4); o2.stop(now + 0.8);
        
        setTimeout(() => ctx.close(), 1000);
    } catch(e) { console.log('Sound failed:', e); }
}

function showAuditToast(message) {
    const toast = document.getElementById('auditNotificationToast');
    if (!toast) { return; }

    toast.innerHTML = '<strong>PC Audit Dashboard</strong><br>' + message.replace(/\n/g, '<br>');
    toast.style.display = 'block';

    setTimeout(function () {
        toast.style.display = 'none';
    }, 9000);
}

function showBrowserNotification(message) {
    if ('Notification' in window && Notification.permission === 'granted') {
        new Notification('PC Audit Dashboard', {
            body: message,
            silent: false
        });
    }
}

function checkAuditDashboardChanges() {
    const previousRaw = localStorage.getItem(AUDIT_NOTIFICATION_STORAGE_KEY);
    localStorage.setItem(AUDIT_NOTIFICATION_STORAGE_KEY, JSON.stringify(currentAuditSnapshot));

    if (!previousRaw) { return; }
    if (localStorage.getItem(AUDIT_NOTIFICATION_ENABLED_KEY) !== 'YES') { return; }

    let previousSnapshot = null;
    try {
        previousSnapshot = JSON.parse(previousRaw);
    } catch (e) {
        return;
    }

    const messages = [];

    if (currentAuditSnapshot.totalAll > previousSnapshot.totalAll) {
        messages.push('New PC audited. Total increased from ' + previousSnapshot.totalAll + ' to ' + currentAuditSnapshot.totalAll + '.');
    }

    if (currentAuditSnapshot.dailyTotal > previousSnapshot.dailyTotal) {
        messages.push('Today audit count increased from ' + previousSnapshot.dailyTotal + ' to ' + currentAuditSnapshot.dailyTotal + '.');
    }

    const issueNames = {
        BarTender: 'BarTender Available PCs',
        PrinterAvailable: 'Printer Available PCs',
        WindowsActivation: 'Windows Not Activated',
        Symantec: 'Symantec Inactive',
        Watermark: 'Watermark Not Available',
        GPO: 'GPO Not Updated'
    };

    Object.keys(currentAuditSnapshot.issues).forEach(function (key) {
        const oldCount = previousSnapshot.issues ? previousSnapshot.issues[key] : 0;
        const newCount = currentAuditSnapshot.issues[key];

        if (newCount !== oldCount) {
            messages.push(issueNames[key] + ' changed from ' + oldCount + ' to ' + newCount + '.');
        }
    });

    if (messages.length > 0) {
        const finalMessage = messages.join('\n');
        playAuditBeep();
        showAuditToast(finalMessage);
        showBrowserNotification(finalMessage);
    }
}

checkAuditDashboardChanges();
</script>

</body>
</html>
"@

   return $html
}

# ====================================================================================================
# 📌 SECTION 08: DETAILS HTML PAGE
# ====================================================================================================
# Function Name : Get-DetailsPage
# Purpose       : Generate popup detail pages based on clicked dashboard card/button.
# Why used      : Users can view exact PC records behind each count.
# Beginner notes:
# • $Type tells function what data to show. Example: Daily_SFC or Symantec_Office.
# • Period views show full master sheet columns.
# • Issue views show important columns and exact status value.
# ----------------------------------------------------------------------------------------------------
function Get-DetailsPage {
   param([string]$Type)
   
   Write-Log -Message "Generating details page for: $Type" -Level 'INFO'
   
   $stats = Get-Stats
   
   $baseType = $Type
   $segmentFilter = $null
   
   # 🔎 Split requested detail type into base type and segment filter.
   # Why: Example Daily_SFC becomes baseType=Daily and segmentFilter=SFC.
   if ($Type -match '^(.*)_(SFC|Office|All)$') {
       $baseType = $Matches[1]
       $segmentFilter = $Matches[2]
       if ($segmentFilter -eq 'All') { $segmentFilter = $null }
   }
   
   # ====================================================================================================
   # 📅 PERIOD BASED DETAILS
   # ====================================================================================================
   # Purpose:
   # • Handles Daily, Weekly, Monthly, and Total detail pages.
   # • Shows full master sheet columns for these views.
   # ----------------------------------------------------------------------------------------------------
   # Handle period-based types (Daily, Weekly, Monthly, Total) with segment filter
   $isPeriodType = @('Daily', 'Weekly', 'Monthly', 'Total') -contains $baseType
   $isIssueType = @('BarTender', 'PrinterAvailable', 'WindowsActivation', 
                    'Symantec', 'Watermark', 'GPO') -contains $baseType
   
   if ($isPeriodType) {
       # PERIOD VIEW WITH SEGMENT FILTER - SHOW FULL MASTER SHEET COLUMNS
       $records = @()
       $title = ""
       $dateRangeInfo = ""
       $segmentInfo = if ($segmentFilter) { "$segmentFilter Network Segment" } else { "All Networks" }
       
       $today = $stats.Today
       $weekStart = $stats.WeekStart
       $monthStart = $stats.MonthStart
       
       switch ($baseType) {
           'Daily' {
               $dateStr = $today.ToString('yyyy-MM-dd')
               $records = @($stats.Data | Where-Object { 
                   $_.Date -and $_.Date.Date -eq $today -and
                   (-not $segmentFilter -or $_.NetworkSegment -eq $segmentFilter)
               })
               $title = "Daily Records - $dateStr - $segmentInfo"
               $dateRangeInfo = "Date: $dateStr | Filter: $segmentInfo"
           }
           'Weekly' {
               $dateStrStart = $weekStart.ToString('yyyy-MM-dd')
               $dateStrEnd = $today.ToString('yyyy-MM-dd')
               $records = @($stats.Data | Where-Object { 
                   $_.Date -and $_.Date.Date -ge $weekStart -and
                   (-not $segmentFilter -or $_.NetworkSegment -eq $segmentFilter)
               })
               $title = "Weekly Records - $dateStrStart to $dateStrEnd - $segmentInfo"
               $dateRangeInfo = "Week: $dateStrStart to $dateStrEnd | Filter: $segmentInfo"
           }
           'Monthly' {
               $dateStrStart = $monthStart.ToString('yyyy-MM-dd')
               $dateStrEnd = $today.ToString('yyyy-MM-dd')
               $records = @($stats.Data | Where-Object { 
                   $_.Date -and $_.Date.Date -ge $monthStart -and
                   (-not $segmentFilter -or $_.NetworkSegment -eq $segmentFilter)
               })
               $title = "Monthly Records - $dateStrStart to $dateStrEnd - $segmentInfo"
               $dateRangeInfo = "Month: $dateStrStart to $dateStrEnd | Filter: $segmentInfo"
           }
           'Total' {
               $records = @($stats.Data | Where-Object { 
                   -not $segmentFilter -or $_.NetworkSegment -eq $segmentFilter
               })
               $title = "All Records - $segmentInfo"
               $dateRangeInfo = "Complete dataset | Filter: $segmentInfo"
           }
       }
       
       if (-not $records) { $records = @() }
       
       # FULL MASTER SHEET COLUMNS
       $columns = @("Duplicate Status") + $stats.OriginalColumns
       $hdr = ($columns | ForEach-Object { "<th><span class='header-filter-wrap'><span>$(Get-HtmlEncoded $_)</span><span class='filter-icon' title='Filter this column'>🔎</span></span></th>" }) -join ""
       
       $rows = ""
       $seqNum = 0
       
       foreach ($r in $records) {
           $seqNum++
           
           $rowClass = ""
           $dupStatus = ""
           
           if ($r._IsDuplicate) {
               $rowClass = " class='duplicate'"
               $dupStatus = "DUPLICATE"
           }
           
           $rowHtml = "<tr$rowClass>"
           $rowHtml += if ($dupStatus) { "<td class='dup-cell'>$dupStatus</td>" } else { "<td></td>" }
           
           foreach ($col in $stats.OriginalColumns) {
               $val = ""
               
               if ($col -eq "SNo" -or $col -eq "S.No" -or $col -eq "SNO" -or $col -eq "SerialNo" -or $col.ToLower().Replace(" ", "").Replace(".", "") -eq "sno") {
                   $val = $seqNum.ToString()
               }
               elseif ($r._OriginalRow -and $r._OriginalRow.$col) {
                   $val = Get-HtmlEncoded $r._OriginalRow.$col.ToString()
               }
               
               $rowHtml += "<td>$val</td>"
           }
           $rowHtml += "</tr>"
           $rows += "$rowHtml`n"
       }
       
       if ($rows -eq "") { 
           $rows = "<tr><td colspan='$($columns.Count)'>No records found for this criteria</td></tr>" 
       }

       $colCount = $columns.Count
       $viewDupCount = ($records | Where-Object { $_._IsDuplicate }).Count
       $dupInfo = ""
       if ($viewDupCount -gt 0) {
           $dupInfo = "<div class='dup-info'>⚠️ Duplicates found: $viewDupCount rows highlighted in orange | Based on Computer Name + Serial Number</div>"
       }

       $segmentBadge = ""
       if ($segmentFilter) {
           $badgeColor = if ($segmentFilter -eq 'SFC') { '#3498db' } else { '#2ecc71' }
           $segmentBadge = "<span style='background:$badgeColor;padding:4px 12px;border-radius:4px;font-size:12px;margin-left:10px;'>$segmentFilter</span>"
       }

       $html = @"
<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<title>$title</title>
<style>
* { margin: 0; padding: 0; box-sizing: border-box; }
body { font-family: 'Segoe UI', Arial, sans-serif; background: #1a1d21; color: #fff; padding: 20px; }
h1 { margin-bottom: 5px; display:flex;align-items:center; }
.info { color: #888; margin-bottom: 10px; }
.daterange { color: #2ecc71; margin-bottom: 10px; font-size: 14px; }
.dup-info { color: #f39c12; margin-bottom: 15px; font-size: 14px; font-weight: 600; }
.segment-info { color: #3498db; margin-bottom: 15px; font-size: 14px; font-weight: 600; }
.export { padding: 10px 20px; background: #27ae60; color: #fff; border: none; border-radius: 6px; cursor: pointer; margin-bottom: 20px; margin-right: 10px; }
.close { padding: 10px 20px; background: #e74c3c; color: #fff; border: none; border-radius: 6px; cursor: pointer; margin-bottom: 20px; }
.table-container { overflow-x: auto; max-width: 100%; }
table { width: 100%; border-collapse: collapse; background: #2d3139; border-radius: 8px; overflow: hidden; font-size: 12px; }
th, td { padding: 10px; text-align: left; border-bottom: 1px solid #444; white-space: nowrap; }
th { background: #363b44; position: sticky; top: 0; }
.filter-panel { display:flex; gap:10px; align-items:center; margin: 0 0 15px 0; flex-wrap:wrap; }
.filter-note { color:#aaa; font-size:12px; }
.filter-btn { padding: 8px 14px; background:#34495e; color:#fff; border:none; border-radius:6px; cursor:pointer; }
.filter-input { width: 100%; min-width: 120px; padding: 7px 28px 7px 8px; border: 1px solid #555; border-radius: 5px; background:#1f232a; color:#fff; font-size: 12px; }
.filter-row th { background:#252a32; top: 38px; padding: 6px; position: sticky; }
.header-filter-wrap { display:flex; align-items:center; gap:6px; }
.filter-icon { color:#f1c40f; font-size:13px; opacity:.95; cursor:pointer; user-select:none; }
.filter-cell { position:relative; }
.filter-cell .filter-icon { position:absolute; right:10px; top:50%; transform:translateY(-50%); }
tr.row-hidden { display:none; }
tr:hover { background: #363b44; }
tr.duplicate { background: #d35400 !important; }
tr.duplicate:hover { background: #e67e22 !important; }
td.dup-cell { color: #f39c12; font-weight: bold; }
</style>
</head>
<body>
<h1>$title $segmentBadge</h1>
$dupInfo
<div class="daterange">$dateRangeInfo</div>
<div class="info">$(($records).Count) records found | $colCount columns</div>
<button class="export" onclick="downloadCSV()">Export CSV</button>
<button class="close" onclick="window.close()">Close</button>
<div class="filter-panel">
  <button class="filter-btn" onclick="clearColumnFilters()">Clear Filters</button>
  <span class="filter-note">Type inside each column filter box to search that column. Export CSV will export only visible/filtered rows.</span>
</div>
<div class="table-container">
<table id="t">
<thead><tr>$hdr</tr><tr class="filter-row" id="filterRow"></tr></thead>
<tbody>$rows</tbody>
</table>
</div>
<script>
function setupColumnFilters() {
    const table = document.getElementById('t');
    if (!table) return;
    const headerCells = table.querySelectorAll('thead tr:first-child th');
    const filterRow = document.getElementById('filterRow');
    if (!filterRow || filterRow.children.length > 0) return;

    headerCells.forEach((th, index) => {
        const fth = document.createElement('th');
        fth.className = 'filter-cell';
        const input = document.createElement('input');
        input.className = 'filter-input';
        input.type = 'text';
        input.placeholder = 'Filter...';
        input.setAttribute('data-col', index);
        input.addEventListener('input', applyColumnFilters);
        const icon = document.createElement('span');
        icon.className = 'filter-icon';
        icon.innerText = '🔎';
        icon.title = 'Filter this column';
        icon.onclick = function(){ input.focus(); };
        fth.appendChild(input);
        fth.appendChild(icon);
        filterRow.appendChild(fth);
    });
}

function applyColumnFilters() {
    const table = document.getElementById('t');
    const filters = Array.from(document.querySelectorAll('.filter-input')).map(i => i.value.toLowerCase().trim());
    const rows = table.querySelectorAll('tbody tr');

    rows.forEach(row => {
        const cells = row.querySelectorAll('td');
        let show = true;
        filters.forEach((filter, index) => {
            if (filter) {
                const cellText = (cells[index]?.innerText || '').toLowerCase();
                if (!cellText.includes(filter)) show = false;
            }
        });
        row.classList.toggle('row-hidden', !show);
    });
}

function clearColumnFilters() {
    document.querySelectorAll('.filter-input').forEach(i => i.value = '');
    applyColumnFilters();
}

function downloadCSV() {
    let csv = [];
    let headerCells = document.querySelectorAll('#t thead tr:first-child th');
    csv.push(Array.from(headerCells).map(c => '"' + c.innerText.replace(/🔎/g, '').trim().replace(/"/g, '""') + '"').join(','));
    let rows = document.querySelectorAll('#t tbody tr:not(.row-hidden)');
    rows.forEach(r => {
        let row = [];
        let cols = r.querySelectorAll('td');
        let isDuplicate = r.classList.contains('duplicate');
        let colIndex = 0;
        cols.forEach(c => {
            let txt = c.innerText.replace(/"/g, '""');
            if (isDuplicate && colIndex === 0) { txt = '[DUPLICATE] ' + txt; }
            row.push('"' + txt + '"');
            colIndex++;
        });
        csv.push(row.join(','));
    });
    let blob = new Blob(['\ufeff' + csv.join('\n')], {type: 'text/csv;charset=utf-8;'});
    let a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = '$title-$(Get-Date -Format 'yyyyMMdd-HHmmss').csv';
    a.click();
}

document.addEventListener('DOMContentLoaded', setupColumnFilters);
</script>
</body>
</html>
"@
   }
   
   # ====================================================================================================
   # ⚠️ ISSUE BASED DETAILS
   # ====================================================================================================
   # Purpose:
   # • Handles BarTender, Printer, Windows Activation, Symantec, Watermark, and GPO pages.
   # • Shows important troubleshooting columns only.
   # ----------------------------------------------------------------------------------------------------
   elseif ($isIssueType) {
       # ISSUE TYPE - SHOW STATUS VALUE WITH EXACT COLUMN NAME
       $records = @()
       $title = ""
       $dateRangeInfo = ""
       $segmentInfo = ""
       $statusColumnName = ""
       
       switch ($baseType) {
           'BarTender' {
               $statusColumnName = "BarTender Available"
               if ($segmentFilter -eq 'SFC') {
                   $records = $stats.Issues['BarTender'].SFC
                   $title = 'BarTender Available - SFC Network'
                   $segmentInfo = "SFC Network (172.29.x.x)"
               } elseif ($segmentFilter -eq 'Office') {
                   $records = $stats.Issues['BarTender'].Office
                   $title = 'BarTender Available - Office Network'
                   $segmentInfo = "Office Network (10.209.x.x / 10.208.x.x)"
               } else {
                   $records = $stats.Issues['BarTender'].All
                   $title = 'BarTender Available - All Networks'
                   $segmentInfo = "All Networks"
               }
               $dateRangeInfo = "Filter: $statusColumnName = YES"
           }
           'PrinterAvailable' {
               $statusColumnName = "Printer Available"
               if ($segmentFilter -eq 'SFC') {
                   $records = $stats.Issues['PrinterAvailable'].SFC
                   $title = 'Printer Available PCs - SFC Network'
                   $segmentInfo = "SFC Network (172.29.x.x)"
               } elseif ($segmentFilter -eq 'Office') {
                   $records = $stats.Issues['PrinterAvailable'].Office
                   $title = 'Printer Available PCs - Office Network'
                   $segmentInfo = "Office Network (10.209.x.x / 10.208.x.x)"
               } else {
                   $records = $stats.Issues['PrinterAvailable'].All
                   $title = 'Printer Available PCs - All Networks'
                   $segmentInfo = "All Networks"
               }
               $dateRangeInfo = "Filter: $statusColumnName ≠ NO"
           }
           'WindowsActivation' {
               $statusColumnName = "Windows Activation"
               if ($segmentFilter -eq 'SFC') {
                   $records = $stats.Issues['WindowsActivation'].SFC
                   $title = 'Windows Not Activated - SFC Network'
                   $segmentInfo = "SFC Network (172.29.x.x)"
               } elseif ($segmentFilter -eq 'Office') {
                   $records = $stats.Issues['WindowsActivation'].Office
                   $title = 'Windows Not Activated - Office Network'
                   $segmentInfo = "Office Network (10.209.x.x / 10.208.x.x)"
               } else {
                   $records = $stats.Issues['WindowsActivation'].All
                   $title = 'Windows Not Activated - All Networks'
                   $segmentInfo = "All Networks"
               }
               $dateRangeInfo = "Filter: $statusColumnName ≠ YES"
           }
           'Symantec' {
               $statusColumnName = "Symantec Status"
               if ($segmentFilter -eq 'SFC') {
                   $records = $stats.Issues['Symantec'].SFC
                   $title = 'Symantec Inactive - SFC Network'
                   $segmentInfo = "SFC Network (172.29.x.x)"
               } elseif ($segmentFilter -eq 'Office') {
                   $records = $stats.Issues['Symantec'].Office
                   $title = 'Symantec Inactive - Office Network'
                   $segmentInfo = "Office Network (10.209.x.x / 10.208.x.x)"
               } else {
                   $records = $stats.Issues['Symantec'].All
                   $title = 'Symantec Inactive - All Networks'
                   $segmentInfo = "All Networks"
               }
               $dateRangeInfo = "Filter: $statusColumnName ≠ ACTIVE"
           }
           'Watermark' {
               $statusColumnName = "Watermark Status"
               if ($segmentFilter -eq 'SFC') {
                   $records = $stats.Issues['Watermark'].SFC
                   $title = 'Watermark Not Available - SFC Network'
                   $segmentInfo = "SFC Network (172.29.x.x)"
               } elseif ($segmentFilter -eq 'Office') {
                   $records = $stats.Issues['Watermark'].Office
                   $title = 'Watermark Not Available - Office Network'
                   $segmentInfo = "Office Network (10.209.x.x / 10.208.x.x)"
               } else {
                   $records = $stats.Issues['Watermark'].All
                   $title = 'Watermark Not Available - All Networks'
                   $segmentInfo = "All Networks"
               }
               $dateRangeInfo = "Filter: $statusColumnName ≠ YES"
           }
           'GPO' {
               $statusColumnName = "GPO Update Status"
               if ($segmentFilter -eq 'SFC') {
                   $records = $stats.Issues['GPO'].SFC
                   $title = 'GPO Not Updated - SFC Network'
                   $segmentInfo = "SFC Network (172.29.x.x)"
               } elseif ($segmentFilter -eq 'Office') {
                   $records = $stats.Issues['GPO'].Office
                   $title = 'GPO Not Updated - Office Network'
                   $segmentInfo = "Office Network (10.209.x.x / 10.208.x.x)"
               } else {
                   $records = $stats.Issues['GPO'].All
                   $title = 'GPO Not Updated - All Networks'
                   $segmentInfo = "All Networks"
               }
               $dateRangeInfo = "Filter: $statusColumnName ≠ UPDATED"
           }
       }
       
       if (-not $records) { $records = @() }
       
       # Get the actual status value from the record
       $hdr = "<th><span class='header-filter-wrap'><span>SNo</span><span class='filter-icon' title='Filter this column'>🔎</span></span></th><th><span class='header-filter-wrap'><span>Computer Name</span><span class='filter-icon' title='Filter this column'>🔎</span></span></th><th><span class='header-filter-wrap'><span>IP Address</span><span class='filter-icon' title='Filter this column'>🔎</span></span></th><th><span class='header-filter-wrap'><span>Network</span><span class='filter-icon' title='Filter this column'>🔎</span></span></th><th><span class='header-filter-wrap'><span>$statusColumnName</span><span class='filter-icon' title='Filter this column'>🔎</span></span></th><th><span class='header-filter-wrap'><span>Auditor ID</span><span class='filter-icon' title='Filter this column'>🔎</span></span></th><th><span class='header-filter-wrap'><span>Audit Date</span><span class='filter-icon' title='Filter this column'>🔎</span></span></th>"
       
       $rows = ""
       $seqNum = 0
       
       foreach ($r in $records) {
           $seqNum++
           
           # Get actual status value from the record
           $statusValue = ""
           switch ($baseType) {
               'BarTender' { $statusValue = $r.BarTender }
               'PrinterAvailable' { $statusValue = $r.Printer }
               'WindowsActivation' { $statusValue = $r.WindowsActivation }
               'Symantec' { $statusValue = $r.SymantecStatus }
               'Watermark' { $statusValue = $r.WatermarkStatus }
               'GPO' { $statusValue = $r.GPOStatus }
           }
           
           $computerName = Get-HtmlEncoded $r.ComputerName
           $ipAddress = Get-HtmlEncoded $r.IPAddress
           $network = $r.NetworkSegment
           $auditorId = Get-HtmlEncoded $r.AuditorID
           $auditDate = if ($r.Date) { $r.Date.ToString('yyyy-MM-dd') } else { Get-HtmlEncoded $r.DateString }
           
           $statusValueHtml = Get-HtmlEncoded $statusValue
           
           $rows += "<tr><td>$seqNum</td><td>$computerName</td><td>$ipAddress</td><td>$network</td><td class='status-value'>$statusValueHtml</td><td>$auditorId</td><td>$auditDate</td></tr>`n"
       }
       
       if ($rows -eq "") { 
           $rows = "<tr><td colspan='7'>No records found for this criteria</td></tr>" 
       }

       $segmentBadge = ""
       if ($segmentFilter) {
           $badgeColor = if ($segmentFilter -eq 'SFC') { '#3498db' } else { '#2ecc71' }
           $segmentBadge = "<span style='background:$badgeColor;padding:4px 12px;border-radius:4px;font-size:12px;margin-left:10px;'>$segmentFilter</span>"
       }

       $html = @"
<!DOCTYPE html>
<html>
<head>
<meta charset="UTF-8">
<title>$title</title>
<style>
* { margin: 0; padding: 0; box-sizing: border-box; }
body { font-family: 'Segoe UI', Arial, sans-serif; background: #1a1d21; color: #fff; padding: 20px; }
h1 { margin-bottom: 5px; display:flex;align-items:center; }
.info { color: #888; margin-bottom: 10px; }
.daterange { color: #2ecc71; margin-bottom: 10px; font-size: 14px; }
.segment-info { color: #3498db; margin-bottom: 15px; font-size: 14px; font-weight: 600; }
.export { padding: 10px 20px; background: #27ae60; color: #fff; border: none; border-radius: 6px; cursor: pointer; margin-bottom: 20px; margin-right: 10px; }
.close { padding: 10px 20px; background: #e74c3c; color: #fff; border: none; border-radius: 6px; cursor: pointer; margin-bottom: 20px; }
.table-container { overflow-x: auto; max-width: 100%; }
table { width: 100%; border-collapse: collapse; background: #2d3139; border-radius: 8px; overflow: hidden; font-size: 12px; }
th, td { padding: 10px; text-align: left; border-bottom: 1px solid #444; white-space: nowrap; }
th { background: #363b44; position: sticky; top: 0; }
.filter-panel { display:flex; gap:10px; align-items:center; margin: 0 0 15px 0; flex-wrap:wrap; }
.filter-note { color:#aaa; font-size:12px; }
.filter-btn { padding: 8px 14px; background:#34495e; color:#fff; border:none; border-radius:6px; cursor:pointer; }
.filter-input { width: 100%; min-width: 120px; padding: 7px 28px 7px 8px; border: 1px solid #555; border-radius: 5px; background:#1f232a; color:#fff; font-size: 12px; }
.filter-row th { background:#252a32; top: 38px; padding: 6px; position: sticky; }
.header-filter-wrap { display:flex; align-items:center; gap:6px; }
.filter-icon { color:#f1c40f; font-size:13px; opacity:.95; cursor:pointer; user-select:none; }
.filter-cell { position:relative; }
.filter-cell .filter-icon { position:absolute; right:10px; top:50%; transform:translateY(-50%); }
tr.row-hidden { display:none; }
tr:hover { background: #363b44; }
td.status-value { color: #e74c3c; font-weight: 600; }
</style>
</head>
<body>
<h1>$title $segmentBadge</h1>
<div class="daterange">$dateRangeInfo</div>
<div class="segment-info">$segmentInfo</div>
<div class="info">$(($records).Count) records found</div>
<button class="export" onclick="downloadCSV()">Export CSV</button>
<button class="close" onclick="window.close()">Close</button>
<div class="filter-panel">
  <button class="filter-btn" onclick="clearColumnFilters()">Clear Filters</button>
  <span class="filter-note">Type inside each column filter box to search that column. Export CSV will export only visible/filtered rows.</span>
</div>
<div class="table-container">
<table id="t">
<thead><tr>$hdr</tr><tr class="filter-row" id="filterRow"></tr></thead>
<tbody>$rows</tbody>
</table>
</div>
<script>
function setupColumnFilters() {
    const table = document.getElementById('t');
    if (!table) return;
    const headerCells = table.querySelectorAll('thead tr:first-child th');
    const filterRow = document.getElementById('filterRow');
    if (!filterRow || filterRow.children.length > 0) return;

    headerCells.forEach((th, index) => {
        const fth = document.createElement('th');
        fth.className = 'filter-cell';
        const input = document.createElement('input');
        input.className = 'filter-input';
        input.type = 'text';
        input.placeholder = 'Filter...';
        input.setAttribute('data-col', index);
        input.addEventListener('input', applyColumnFilters);
        const icon = document.createElement('span');
        icon.className = 'filter-icon';
        icon.innerText = '🔎';
        icon.title = 'Filter this column';
        icon.onclick = function(){ input.focus(); };
        fth.appendChild(input);
        fth.appendChild(icon);
        filterRow.appendChild(fth);
    });
}

function applyColumnFilters() {
    const table = document.getElementById('t');
    const filters = Array.from(document.querySelectorAll('.filter-input')).map(i => i.value.toLowerCase().trim());
    const rows = table.querySelectorAll('tbody tr');

    rows.forEach(row => {
        const cells = row.querySelectorAll('td');
        let show = true;
        filters.forEach((filter, index) => {
            if (filter) {
                const cellText = (cells[index]?.innerText || '').toLowerCase();
                if (!cellText.includes(filter)) show = false;
            }
        });
        row.classList.toggle('row-hidden', !show);
    });
}

function clearColumnFilters() {
    document.querySelectorAll('.filter-input').forEach(i => i.value = '');
    applyColumnFilters();
}

function downloadCSV() {
    let csv = [];
    let headerCells = document.querySelectorAll('#t thead tr:first-child th');
    csv.push(Array.from(headerCells).map(c => '"' + c.innerText.replace(/🔎/g, '').trim().replace(/"/g, '""') + '"').join(','));
    let rows = document.querySelectorAll('#t tbody tr:not(.row-hidden)');
    rows.forEach(r => {
        let row = [];
        let cols = r.querySelectorAll('td');
        cols.forEach(c => {
            let txt = c.innerText.replace(/"/g, '""');
            row.push('"' + txt + '"');
        });
        csv.push(row.join(','));
    });
    let blob = new Blob(['\ufeff' + csv.join('\n')], {type: 'text/csv;charset=utf-8;'});
    let a = document.createElement('a');
    a.href = URL.createObjectURL(blob);
    a.download = '$title-$(Get-Date -Format 'yyyyMMdd-HHmmss').csv';
    a.click();
}

document.addEventListener('DOMContentLoaded', setupColumnFilters);
</script>
</body>
</html>
"@
   }
   
   else {
       $html = @"
<!DOCTYPE html>
<html>
<head><title>Unknown Type</title></head>
<body style="font-family:Arial;padding:20px;">
<h1>Unknown Detail Type: $Type</h1>
<p>Valid types: Daily_SFC, Daily_Office, Weekly_SFC, Weekly_Office, Monthly_SFC, Monthly_Office, 
Total_SFC, Total_Office, BarTender_SFC, BarTender_Office, PrinterAvailable_SFC, PrinterAvailable_Office,
WindowsActivation_SFC, WindowsActivation_Office, Symantec_SFC, Symantec_Office,
Watermark_SFC, Watermark_Office, GPO_SFC, GPO_Office</p>
<p>Or use _All suffix for combined view: Daily_All, Weekly_All, Monthly_All, Total_All</p>
<button onclick="window.close()">Close</button>
</body>
</html>
"@
   }

   return $html
}

# ====================================================================================================
# 📌 SECTION 09: HTTP SERVER ENGINE
# ====================================================================================================
# Purpose:
# • Starts a PowerShell HttpListener web server.
# • Serves dashboard page and details pages to browser users.
# • Keeps listening until script is stopped.
# ====================================================================================================
# ====================================================================================================
# SECTION 8: HTTP SERVER ENGINE
# ====================================================================================================


# ----------------------------------------------------------------------------------------------------
# FUNCTION: Test-PcAuditBasicAuth
# PURPOSE : Ask browser users for username/password before opening dashboard pages.
# LOGIN   : Username = PCAudit | Password = PC_Audit
# NOTE    : Works for IP URL and DNS URL, for example http://pc_audit_dashboard:8080
# ----------------------------------------------------------------------------------------------------
function Test-PcAuditBasicAuth {
    param(
        [Parameter(Mandatory=$true)]
        [System.Net.HttpListenerContext]$Context
    )

    $AuthUser = 'PCAudit'
    $AuthPass = 'PC_Audit'
    $Realm    = 'PC Audit Dashboard'

    $AuthHeader = $Context.Request.Headers['Authorization']

    if ([string]::IsNullOrWhiteSpace($AuthHeader) -or (-not $AuthHeader.StartsWith('Basic ', [System.StringComparison]::OrdinalIgnoreCase))) {
        $Context.Response.StatusCode = 401
        $Context.Response.AddHeader('WWW-Authenticate', 'Basic realm="' + $Realm + '"')
        $Context.Response.Close()
        return $false
    }

    try {
        $Encoded = $AuthHeader.Substring(6).Trim()
        $Decoded = [System.Text.Encoding]::UTF8.GetString([Convert]::FromBase64String($Encoded))
        $Parts = $Decoded.Split(':', 2)

        if ($Parts.Count -lt 2 -or $Parts[0] -ne $AuthUser -or $Parts[1] -ne $AuthPass) {
            $Context.Response.StatusCode = 401
            $Context.Response.AddHeader('WWW-Authenticate', 'Basic realm="' + $Realm + '"')
            $Context.Response.Close()
            return $false
        }

        return $true
    }
    catch {
        $Context.Response.StatusCode = 401
        $Context.Response.AddHeader('WWW-Authenticate', 'Basic realm="' + $Realm + '"')
        $Context.Response.Close()
        return $false
    }
}

# ----------------------------------------------------------------------------------------------------
# FUNCTION: Start-DashboardServer
# PURPOSE:  Initialize and run the HTTP listener, handling incoming requests
# 
# URL RESERVATION:
#   If you get "access denied" errors, run this as Administrator:
#   netsh http add urlacl url=http://10.209.110.220:8080/ user=YOUR_USERNAME
#
# REQUEST HANDLING:
#   - GET /           → Returns main dashboard (Get-MainPage)
#   - GET /details    → Returns detail view (Get-DetailsPage with type parameter)
#
# AUTO-START:
#   Automatically opens browser to dashboard URL on startup
# ----------------------------------------------------------------------------------------------------
   # 🧠 Beginner explanation:
   # • HttpListener is a built-in .NET class used by PowerShell to receive web requests.
   # • Browser requests / for main dashboard and /details for popup pages.
   # • The server returns HTML text as the response.
function Start-DashboardServer {
   
   # --- STEP 8.1: INITIALIZE HTTP LISTENER ---
   # 🌐 Create HTTP listener object.
   # Why: This object waits for browser requests and sends dashboard HTML responses.
   $script:Listener = New-Object System.Net.HttpListener
   
   # IMPORTANT: Change IP if needed for your network
   # 🔗 Server binding URL.
   # Why: This tells HttpListener which IP address and port to listen on.
   $prefix = "http://10.209.110.220:$Port/"
   $script:Listener.Prefixes.Add($prefix)
   
   # --- STEP 8.2: START SERVER ---
   try {
       $script:Listener.Start()
   }
   catch {
       Write-Log -Message "Failed to start server: $_" -Level 'ERROR'
       # EXE FIX COMMENTED: Console Write-Host can become popup in converted EXE.
       # Write-Host "URL reservation needed. Run as Admin or:" -ForegroundColor Yellow
       # EXE FIX COMMENTED: Console Write-Host can become popup in converted EXE.
       # Write-Host "netsh http add urlacl url=http://10.209.110.220:$Port user=$env:USERNAME" -ForegroundColor Cyan
       throw
   }
   
   # --- STEP 8.3: DISPLAY STARTUP INFO ---
   # EXE FIX COMMENTED: Console Write-Host can become popup in converted EXE.
   # Write-Host "`n✅ SERVER STARTED" -ForegroundColor Green
   # EXE FIX COMMENTED: Console Write-Host can become popup in converted EXE.
   # Write-Host "URL: http://pc_Audit_Dashboard:$Port" -ForegroundColor Cyan
   # EXE FIX COMMENTED: Console Write-Host can become popup in converted EXE.
   # Write-Host "Log: $script:LogFile" -ForegroundColor Gray
   # EXE FIX COMMENTED: Console Write-Host can become popup in converted EXE.
   # Write-Host "Press Ctrl+C to stop`n" -ForegroundColor Yellow
   Write-Log -Message "Server started successfully on port $Port" -Level 'INFO'
   
   # Auto-open browser (silently continue if fails)
   Start-Process "http://pc_Audit_Dashboard:$Port" -ErrorAction SilentlyContinue

   # --- STEP 8.4: MAIN REQUEST LOOP ---
   $requestCount = 0
   
   # 🔁 Main server loop.
   # Why: Keeps dashboard active and handles each incoming browser request.
   while ($script:RunServer -and $script:Listener.IsListening) {
       try {
           # Wait for incoming HTTP request (blocking call)
           # ⏳ Wait for a browser/client request.
           # Why: Execution pauses here until someone opens or refreshes the dashboard.
           $context = $script:Listener.GetContext()
           
           # Parse request details
           $path = $context.Request.Url.LocalPath      # / or /details
           $query = $context.Request.Url.Query          # ?type=Daily
           $clientIP = $context.Request.RemoteEndPoint.Address.ToString()
           
           $requestCount++
           Write-Log -Message "Request #$requestCount from $clientIP : $path$query" -Level 'INFO'

           # --- LOGIN AUTHENTICATION ---
           # Browser will ask username/password before showing dashboard.
           if (-not (Test-PcAuditBasicAuth -Context $context)) {
               Write-Log -Message "Unauthorized request from $clientIP : $path$query" -Level 'WARN'
               continue
           }
           
           # --- ROUTING: Determine which page to serve ---
           $responseHtml = ""
           
           # 🧭 Route request to correct page.
           # Why: /audit.ico serves browser tab favicon + title logo, /details opens popup detail page.
           if ($path -eq '/PC Audit Dashbaord Service Controller.ico') {
               if (Test-Path $script:WebIconPath) {
                   $iconBytes = [System.IO.File]::ReadAllBytes($script:WebIconPath)
                   $context.Response.ContentType = 'image/x-icon'
                   $context.Response.ContentLength64 = $iconBytes.Length
                   $context.Response.OutputStream.Write($iconBytes, 0, $iconBytes.Length)
                   $context.Response.OutputStream.Close()
                   continue
               }
               else {
                   Write-Log -Message "Web icon file not found: $script:WebIconPath" -Level 'WARN'
                   $context.Response.StatusCode = 404
                   $context.Response.OutputStream.Close()
                   continue
               }
           }
           elseif ($path -match '^/details') {
               # Detail view: extract type parameter
               $type = [regex]::Match($query, 'type=([^&]+)').Groups[1].Value
               $responseHtml = Get-DetailsPage -Type $type
           } else {
               # Default: main dashboard
               $responseHtml = Get-MainPage
           }
           
           # --- SEND RESPONSE ---
           # 📤 Convert generated HTML text into bytes for browser response.
           # Why: HTTP responses must send byte data through output stream.
           $buffer = [System.Text.Encoding]::UTF8.GetBytes($responseHtml)
           $context.Response.ContentType = "text/html; charset=utf-8"
           $context.Response.ContentLength64 = $buffer.Length
           $context.Response.OutputStream.Write($buffer, 0, $buffer.Length)
           $context.Response.OutputStream.Close()
           
       }
       catch [System.ObjectDisposedException] {
           # Normal shutdown
           break
       }
       catch {
           # Log error but keep server running
           # EXE FIX COMMENTED: Console Write-Warning can become popup/noise in converted EXE.
           # Write-Warning "Request error: $_"
           Write-Log -Message "Request error: $_" -Level 'ERROR'
       }
   }
   
   Write-Log -Message "Server stopped. Total requests: $requestCount" -Level 'INFO'
}

# ----------------------------------------------------------------------------------------------------
# FUNCTION: Stop-Server
# PURPOSE:  Gracefully shutdown the HTTP listener
# TRIGGER:  Called on Ctrl+C or script exit
# ----------------------------------------------------------------------------------------------------
# ====================================================================================================
# 🛑 FUNCTION: Stop-Server
# ====================================================================================================
# PURPOSE: Gracefully shutdown the HTTP listener.
# TRIGGER: Called on Ctrl+C or script exit.
# Beginner note: Prevents the port from staying locked after closing script.
# ----------------------------------------------------------------------------------------------------
function Stop-Server {
   $script:RunServer = $false    # Signal main loop to exit
   
   if ($script:Listener) {
       try { $script:Listener.Stop() } catch {}
       try { $script:Listener.Close() } catch {}
   }
   
   Write-Log -Message "Server shutdown initiated" -Level 'INFO'
}


# ====================================================================================================
# SECTION 10: WINFORMS GUI SERVICE MANAGER - ADDED LAYER
# ====================================================================================================
# Purpose:
# - Run this file and show Login + GUI first instead of console startup.
# - START button starts the original dashboard HTTP server.
# - STOP button calls original Stop-Server and closes the listener.
# - Dashboard calculation / HTML / details / CSV logic above is not changed.
# ====================================================================================================

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

# Hide PowerShell console window when possible
try {
    Add-Type @"
using System;
using System.Runtime.InteropServices;
public class Win32Console {
    [DllImport("kernel32.dll")] public static extern IntPtr GetConsoleWindow();
    [DllImport("user32.dll")] public static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);
}
"@
    $consoleHandle = [Win32Console]::GetConsoleWindow()
    if ($consoleHandle -ne [IntPtr]::Zero) { [Win32Console]::ShowWindow($consoleHandle, 0) | Out-Null }
} catch { }

[System.Windows.Forms.Application]::EnableVisualStyles()
[System.Windows.Forms.Application]::SetCompatibleTextRenderingDefault($false)

# GUI globals
$script:GuiVersion = "2026.05.30 Service Controller Compact"
$script:DashboardHost = "pc_Audit_Dashboard"
$script:ServerRunspace = $null
$script:ServerPowerShell = $null
$script:ServerAsync = $null
$script:ServerStartedAt = $null
$script:MainForm = $null
$script:LogBox = $null
$script:StatusLabel = $null
$script:DetailLabel = $null
$script:UptimeLabel = $null
$script:BtnStart = $null
$script:BtnStop = $null
$script:BtnRestart = $null
$script:BtnOpen = $null
$script:BtnTest = $null
$script:BtnSettings = $null
$script:BtnExit = $null
$script:LockPanel = $null
$script:LockUserBox = $null
$script:LockPassBox = $null
$script:LockMessageLabel = $null
$script:ControllerLocked = $false
$script:AutoLockMinutes = 5
$script:AutoLockSeconds = 300
$script:MiStart = $null
$script:MiStop = $null
$script:MiRestart = $null
$script:MiTest = $null
$script:MiSettings = $null
$script:MiExit = $null
$script:NotifyIcon = $null
$script:LastLogLength = 0
$script:PendingContext = $null
$script:RequestCount = 0
$script:SettingsFile = Join-Path $script:ExeSafeScriptRoot "pc_audit_gui_settings.json"


# ====================================================================================================
# CONTROLLER AUTO-LOCK SETTINGS
# ====================================================================================================
# Purpose:
# • If the PC is idle for 5 minutes, only the WinForms controller is locked.
# • The dashboard web server continues running in the background.
# • Locked users cannot Stop / Restart / change Settings / Exit until unlocked.
# ----------------------------------------------------------------------------------------------------
try {
    Add-Type @"
using System;
using System.Runtime.InteropServices;
public static class PcAuditIdleTime {
    [StructLayout(LayoutKind.Sequential)]
    public struct LASTINPUTINFO {
        public uint cbSize;
        public uint dwTime;
    }
    [DllImport("user32.dll")]
    public static extern bool GetLastInputInfo(ref LASTINPUTINFO plii);
    public static uint GetIdleMilliseconds() {
        LASTINPUTINFO lii = new LASTINPUTINFO();
        lii.cbSize = (uint)System.Runtime.InteropServices.Marshal.SizeOf(typeof(LASTINPUTINFO));
        if (!GetLastInputInfo(ref lii)) { return 0; }
        return ((uint)Environment.TickCount - lii.dwTime);
    }
}
"@ -ErrorAction SilentlyContinue
} catch { }

function Get-ControllerIdleSeconds {
    try {
        return [int]([PcAuditIdleTime]::GetIdleMilliseconds() / 1000)
    } catch {
        return 0
    }
}

function Lock-Controller {
    param([string]$Reason = "Auto locked after 5 minutes of inactivity.")
    if ($script:ControllerLocked) { return }
    $script:ControllerLocked = $true
    Add-GuiLog "Controller locked. Server continues running." "WARN"
    if ($script:LockMessageLabel) { $script:LockMessageLabel.Text = $Reason }
    if ($script:LockUserBox) { $script:LockUserBox.Text = "" }
    if ($script:LockPassBox) { $script:LockPassBox.Text = "" }
    if ($script:LockPanel) {
        $script:LockPanel.Visible = $true
        $script:LockPanel.BringToFront()
        try { $script:LockUserBox.Focus() } catch { }
    }
    Update-ButtonState
}

function Unlock-Controller {
    $user = ""
    $pass = ""
    if ($script:LockUserBox) { $user = $script:LockUserBox.Text.Trim() }
    if ($script:LockPassBox) { $pass = $script:LockPassBox.Text }

    # Unlock uses the same controller/admin login as the startup controller login.
    if ($user -eq "HH0010520" -and $pass -eq "Foxconn-FXCN-IT") {
        $script:ControllerLocked = $false
        if ($script:LockPanel) { $script:LockPanel.Visible = $false }
        if ($script:LockPassBox) { $script:LockPassBox.Text = "" }
        Add-GuiLog "Controller unlocked." "INFO"
        Update-ButtonState
        return $true
    }
    else {
        if ($script:LockMessageLabel) { $script:LockMessageLabel.Text = "Invalid username or password. Please try again." }
        if ($script:LockPassBox) { $script:LockPassBox.Text = ""; $script:LockPassBox.Focus() }
        return $false
    }
}

function Require-ControllerUnlock {
    if (-not $script:ControllerLocked) { return $true }
    if ($script:LockPanel) {
        $script:LockPanel.Visible = $true
        $script:LockPanel.BringToFront()
        try { $script:LockUserBox.Focus() } catch { }
    }
    return $false
}

function Initialize-GuiSettings {
    $defaults = [ordered]@{
        Port = $Port
        BasePath = $BasePath
        LogPath = $script:LogPath
        DashboardHost = $script:DashboardHost
        DashboardIP = "10.209.110.220"
        CsvFileName = "master_audit.csv"
        AutoOpenBrowser = $true
        MinimizeToTray = $true
        StartHidden = $false
        AutoRefreshLog = $true
    }

    if (Test-Path $script:SettingsFile) {
        try {
            $loaded = Get-Content $script:SettingsFile -Raw | ConvertFrom-Json
            foreach ($k in @($defaults.Keys)) {
                if ($null -ne $loaded.$k) { $defaults[$k] = $loaded.$k }
            }
        } catch { }
    }

    $script:GuiSettings = [pscustomobject]$defaults
    Apply-GuiSettings
}

function Apply-GuiSettings {
    $script:Port = [int]$script:GuiSettings.Port
    Set-Variable -Name Port -Scope Script -Value ([int]$script:GuiSettings.Port)
    Set-Variable -Name BasePath -Scope Script -Value ([string]$script:GuiSettings.BasePath)
    $script:DashboardHost = [string]$script:GuiSettings.DashboardHost
    $script:DashboardIP = [string]$script:GuiSettings.DashboardIP
    $script:CsvFileName = [string]$script:GuiSettings.CsvFileName
    if ([string]::IsNullOrWhiteSpace($script:CsvFileName)) { $script:CsvFileName = "master_audit.csv" }
    $safeBasePath = [string]$script:GuiSettings.BasePath
    if ([string]::IsNullOrWhiteSpace($safeBasePath)) {
        # EXE FIX: Prevent Join-Path empty BasePath error.
        $safeBasePath = "\\10.208.193.241\pc_audit$\PCAUDIT-IT"
        $script:GuiSettings.BasePath = $safeBasePath
    }
    $safeLogPath = [string]$script:GuiSettings.LogPath
    if ([string]::IsNullOrWhiteSpace($safeLogPath)) {
        # EXE FIX: Prevent Join-Path empty LogPath error.
        $safeLogPath = "D:\PC_AUDIT_HOST\PC_AUDIT_LOG\LOGS"
        $script:GuiSettings.LogPath = $safeLogPath
    }
    $script:LogPath = $safeLogPath
    $script:LogFile = Get-DailyLogFile
    $script:MasterFile = Join-Path $safeBasePath $script:CsvFileName
}

function Get-GuiDashboardAddress {
    if (-not [string]::IsNullOrWhiteSpace($script:DashboardIP)) { return $script:DashboardIP }
    if (-not [string]::IsNullOrWhiteSpace($script:DashboardHost)) { return $script:DashboardHost }
    return "pc_Audit_Dashboard"
}

function Get-GuiDashboardUrl {
    return "http://$(Get-GuiDashboardAddress):$Port"
}

function Save-GuiSettings {
    try {
        $script:GuiSettings | ConvertTo-Json -Depth 4 | Set-Content -Path $script:SettingsFile -Encoding UTF8
    } catch {
        [System.Windows.Forms.MessageBox]::Show("Unable to save settings: $_", "Settings", "OK", "Warning") | Out-Null
    }
}

function Add-GuiLog {
    param([string]$Message, [string]$Level = "INFO")
    $line = "[{0}] [{1}] {2}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Level, $Message
    try { Write-Log -Message $Message -Level $Level } catch { }
    if ($script:LogBox -and -not $script:LogBox.IsDisposed) {
        $append = {
            param($txt)
            $script:LogBox.AppendText($txt + [Environment]::NewLine)
            $script:LogBox.SelectionStart = $script:LogBox.TextLength
            $script:LogBox.ScrollToCaret()
        }
        if ($script:LogBox.InvokeRequired) { $script:LogBox.BeginInvoke($append, $line) | Out-Null } else { & $append $line }
    }
}

function Is-DashboardRunning {
    return ($script:Listener -ne $null -and $script:Listener.IsListening)
}

function Start-GuiDashboardListener {
    Apply-GuiSettings
    if (Is-DashboardRunning) { return }

    $script:RunServer = $true
    $script:RequestCount = 0
    $script:PendingContext = $null
    $script:Listener = New-Object System.Net.HttpListener

    # Keep the original binding if your dashboard host name is only an alias, but allow Settings to change it.
    # Use + only when you have the matching URLACL/admin permission.
    $bindHost = Get-GuiDashboardAddress
    if ([string]::IsNullOrWhiteSpace($bindHost)) { $bindHost = 'pc_Audit_Dashboard' }
    $prefix = "http://$bindHost`:$Port/"
    $script:Listener.Prefixes.Add($prefix)

    try {
        $script:Listener.Start()
        $script:ServerStartedAt = Get-Date
        Write-Log -Message "Server started successfully on $prefix" -Level 'INFO'
        Add-GuiLog "SERVER STARTED: $prefix" "INFO"
        Begin-GuiDashboardAccept
    }
    catch {
        try { $script:Listener.Close() } catch { }
        $script:Listener = $null
        $script:RunServer = $false
        throw "Failed to start listener at $prefix. Run PowerShell as Administrator or add URLACL for this URL. Details: $_"
    }
}

function Begin-GuiDashboardAccept {
    if (-not (Is-DashboardRunning)) { return }
    if ($script:PendingContext -ne $null -and -not $script:PendingContext.IsCompleted) { return }
    try {
        $script:PendingContext = $script:Listener.BeginGetContext($null, $null)
    } catch {
        if ($script:RunServer) { Add-GuiLog "Listener accept failed: $_" "ERROR" }
    }
}

function Process-GuiDashboardRequests {
    if (-not (Is-DashboardRunning)) { return }
    if ($script:PendingContext -eq $null) { Begin-GuiDashboardAccept; return }
    if (-not $script:PendingContext.IsCompleted) { return }

    $async = $script:PendingContext
    $script:PendingContext = $null

    try {
        $context = $script:Listener.EndGetContext($async)
        $path = $context.Request.Url.LocalPath
        $query = $context.Request.Url.Query
        $clientIP = $context.Request.RemoteEndPoint.Address.ToString()
        $script:RequestCount++
        Write-Log -Message "Request #$($script:RequestCount) from $clientIP : $path$query" -Level 'INFO'

        # --- LOGIN AUTHENTICATION ---
        # Browser will ask username/password before showing dashboard.
        if (-not (Test-PcAuditBasicAuth -Context $context)) {
            Write-Log -Message "Unauthorized request from $clientIP : $path$query" -Level 'WARN'
            Begin-GuiDashboardAccept
            return
        }

        if ($path -eq '/PC Audit Dashbaord Service Controller.ico') {
            if (Test-Path $script:WebIconPath) {
                $iconBytes = [System.IO.File]::ReadAllBytes($script:WebIconPath)
                $context.Response.ContentType = 'image/x-icon'
                $context.Response.ContentLength64 = $iconBytes.Length
                $context.Response.OutputStream.Write($iconBytes, 0, $iconBytes.Length)
                $context.Response.OutputStream.Close()
                Begin-GuiDashboardAccept
                return
            }
            else {
                Write-Log -Message "Web icon file not found: $script:WebIconPath" -Level 'WARN'
                $context.Response.StatusCode = 404
                $context.Response.OutputStream.Close()
                Begin-GuiDashboardAccept
                return
            }
        }
        elseif ($path -match '^/details') {
            $type = [regex]::Match($query, 'type=([^&]+)').Groups[1].Value
            $responseHtml = Get-DetailsPage -Type $type
        } else {
            $responseHtml = Get-MainPage
        }

        $buffer = [System.Text.Encoding]::UTF8.GetBytes($responseHtml)
        $context.Response.ContentType = "text/html; charset=utf-8"
        $context.Response.ContentLength64 = $buffer.Length
        $context.Response.OutputStream.Write($buffer, 0, $buffer.Length)
        $context.Response.OutputStream.Close()
    }
    catch [System.ObjectDisposedException] { }
    catch {
        Add-GuiLog "Request error: $_" "ERROR"
    }
    finally {
        Begin-GuiDashboardAccept
    }
}

function Update-ButtonState {
    $running = Is-DashboardRunning
    $locked = [bool]$script:ControllerLocked
    if ($script:BtnStart) { $script:BtnStart.Enabled = ((-not $running) -and (-not $locked)) }
    if ($script:BtnStop) { $script:BtnStop.Enabled = ($running -and (-not $locked)) }
    if ($script:BtnRestart) { $script:BtnRestart.Enabled = (-not $locked) }
    if ($script:BtnTest) { $script:BtnTest.Enabled = (-not $locked) }
    if ($script:BtnSettings) { $script:BtnSettings.Enabled = (-not $locked) }
    if ($script:BtnExit) { $script:BtnExit.Enabled = (-not $locked) }
    if ($script:MiStart) { $script:MiStart.Enabled = ((-not $running) -and (-not $locked)) }
    if ($script:MiStop) { $script:MiStop.Enabled = ($running -and (-not $locked)) }
    if ($script:MiRestart) { $script:MiRestart.Enabled = (-not $locked) }
    if ($script:MiTest) { $script:MiTest.Enabled = (-not $locked) }
    if ($script:MiSettings) { $script:MiSettings.Enabled = (-not $locked) }
    if ($script:MiExit) { $script:MiExit.Enabled = (-not $locked) }

    if ($script:StatusLabel) {
        if ($locked) {
            $script:StatusLabel.Text = "● LOCKED"
            $script:StatusLabel.ForeColor = [System.Drawing.Color]::FromArgb(241,196,15)
        }
        elseif ($running) {
            $script:StatusLabel.Text = "● RUNNING"
            $script:StatusLabel.ForeColor = [System.Drawing.Color]::FromArgb(46,204,113)
        } else {
            $script:StatusLabel.Text = "● STOPPED"
            $script:StatusLabel.ForeColor = [System.Drawing.Color]::FromArgb(231,76,60)
        }
    }
    if ($script:DetailLabel) {
        $live = if ($running) { "LIVE" } else { "OFFLINE" }
        $script:DetailLabel.Text = "$live  |  URL: $(Get-GuiDashboardUrl)  |  Host: $($script:DashboardHost)  |  IP: $($script:DashboardIP)  |  CSV: $($script:MasterFile)"
    }
    if ($script:NotifyIcon) {
        $script:NotifyIcon.Text = if ($running) { "PC Audit Dashboard - Running" } else { "PC Audit Dashboard - Stopped" }
    }
}

function Start-GuiDashboard {
    if (Is-DashboardRunning) {
        Add-GuiLog "Dashboard already running." "WARN"
        Update-ButtonState
        return
    }

    Add-GuiLog "Starting PC Audit Dashboard server on port $Port ..." "INFO"
    try {
        Start-GuiDashboardListener
        Add-GuiLog "START completed. Dashboard server is running." "INFO"
        if ($script:GuiSettings.AutoOpenBrowser) { Open-GuiDashboard }
    } catch {
        Add-GuiLog "Start failed: $_" "ERROR"
        [System.Windows.Forms.MessageBox]::Show("Dashboard start failed.`n`n$_", "Start Failed", "OK", "Error") | Out-Null
        Stop-GuiDashboard
    }
    Update-ButtonState
}

function Stop-GuiDashboard {
    if (-not (Is-DashboardRunning)) {
        Update-ButtonState
        return
    }

    Add-GuiLog "Stopping PC Audit Dashboard server ..." "INFO"
    try { Stop-Server } catch { Add-GuiLog "Stop-Server warning: $_" "WARN" }
    $script:PendingContext = $null
    $script:ServerStartedAt = $null
    Add-GuiLog "Dashboard server stopped. Total requests: $($script:RequestCount)" "INFO"
    Update-ButtonState
}

function Restart-GuiDashboard {
    Add-GuiLog "Restart requested." "INFO"
    if (Is-DashboardRunning) { Stop-GuiDashboard }
    Start-Sleep -Seconds 1
    Start-GuiDashboard
}

function Open-GuiDashboard {
    $url = (Get-GuiDashboardUrl)
    try {
        Start-Process $url -ErrorAction Stop
        Add-GuiLog "Opened dashboard: $url" "INFO"
    } catch {
        Add-GuiLog "Unable to open browser: $_" "ERROR"
    }
}

function Test-GuiDashboardUrl {
    Apply-GuiSettings
    $url = (Get-GuiDashboardUrl)
    Add-GuiLog "Quick Test started: source path, CSV file, log folder, port, and server URL." "INFO"

    $report = New-Object System.Collections.Generic.List[string]
    $hasError = $false

    try {
        if ([string]::IsNullOrWhiteSpace($BasePath)) {
            $report.Add("ERROR: CSV Base Path is empty.")
            Add-GuiLog "TEST ERROR: CSV Base Path is empty." "ERROR"
            $hasError = $true
        }
        elseif (Test-Path $BasePath) {
            $report.Add("OK: Source path reachable: $BasePath")
            Add-GuiLog "TEST OK: Source path reachable: $BasePath" "INFO"
        }
        else {
            $report.Add("ERROR: Source path not reachable: $BasePath")
            Add-GuiLog "TEST ERROR: Source path not reachable: $BasePath" "ERROR"
            $hasError = $true
        }

        if (Test-Path $script:MasterFile) {
            $csvInfo = Get-Item $script:MasterFile
            $report.Add("OK: master_audit.csv found. Size: $([math]::Round($csvInfo.Length/1KB,2)) KB")
            Add-GuiLog "TEST OK: master_audit.csv found: $($script:MasterFile)" "INFO"

            try {
                $sample = Import-Csv $script:MasterFile -ErrorAction Stop | Select-Object -First 1
                if ($null -ne $sample) {
                    $cols = @($sample.PSObject.Properties.Name)
                    $report.Add("OK: CSV readable. Columns found: $($cols.Count)")
                    Add-GuiLog "TEST OK: CSV readable. Columns found: $($cols.Count)" "INFO"
                } else {
                    $report.Add("WARNING: CSV file is readable but appears empty.")
                    Add-GuiLog "TEST WARN: CSV file appears empty." "WARN"
                }
            } catch {
                $report.Add("ERROR: CSV abnormal/read failed: $_")
                Add-GuiLog "TEST ERROR: CSV abnormal/read failed: $_" "ERROR"
                $hasError = $true
            }
        }
        else {
            $report.Add("ERROR: master_audit.csv not found: $($script:MasterFile)")
            Add-GuiLog "TEST ERROR: master_audit.csv not found: $($script:MasterFile)" "ERROR"
            $hasError = $true
        }

        try {
            if (-not (Test-Path $script:LogPath)) { New-Item -ItemType Directory -Path $script:LogPath -Force | Out-Null }
            $testLog = Join-Path $script:LogPath "gui_write_test.tmp"
            "test" | Set-Content -Path $testLog -Encoding UTF8 -ErrorAction Stop
            Remove-Item $testLog -Force -ErrorAction SilentlyContinue
            $report.Add("OK: Log folder writable: $($script:LogPath)")
            Add-GuiLog "TEST OK: Log folder writable: $($script:LogPath)" "INFO"
        } catch {
            $report.Add("ERROR: Log folder not writable: $($script:LogPath) - $_")
            Add-GuiLog "TEST ERROR: Log folder not writable: $($script:LogPath) - $_" "ERROR"
            $hasError = $true
        }

        $portInUse = $false
        try {
            $portInUse = [bool](Get-NetTCPConnection -LocalPort $Port -ErrorAction SilentlyContinue | Where-Object { $_.State -eq 'Listen' })
        } catch {
            try { $portInUse = [bool](netstat -ano | Select-String ":$Port\s") } catch { }
        }

        if (Is-DashboardRunning) {
            $report.Add("OK: Server is RUNNING on $url")
            Add-GuiLog "TEST OK: Server is RUNNING on $url" "INFO"
            try {
                $response = Invoke-WebRequest -Uri $url -UseBasicParsing -TimeoutSec 5
                $report.Add("OK: Dashboard URL reachable. HTTP $($response.StatusCode)")
                Add-GuiLog "TEST OK: Dashboard URL reachable. HTTP $($response.StatusCode)" "INFO"
            } catch {
                $report.Add("WARNING: Server says running, but URL test failed: $_")
                Add-GuiLog "TEST WARN: Server running but URL failed: $_" "WARN"
            }
        }
        elseif ($portInUse) {
            $report.Add("WARNING: Port $Port is already used by another process. START may fail.")
            Add-GuiLog "TEST WARN: Port $Port already used by another process." "WARN"
        }
        else {
            $report.Add("OK: Port $Port appears free. Server currently STOPPED.")
            Add-GuiLog "TEST OK: Port $Port appears free. Server stopped." "INFO"
        }
    }
    catch {
        $report.Add("ERROR: Unexpected quick test failure: $_")
        Add-GuiLog "TEST ERROR: Unexpected quick test failure: $_" "ERROR"
        $hasError = $true
    }

    $icon = if ($hasError) { "Warning" } else { "Information" }
    $title = if ($hasError) { "Quick Test Completed - Issues Found" } else { "Quick Test Completed - OK" }
    [System.Windows.Forms.MessageBox]::Show(($report -join "`n"), $title, "OK", $icon) | Out-Null
    Add-GuiLog "Quick Test completed." "INFO"
}
function Show-SettingsWindow {
    if (Is-DashboardRunning) {
        [System.Windows.Forms.MessageBox]::Show("Please STOP the dashboard before changing settings.", "Settings", "OK", "Information") | Out-Null
        return
    }

    $f = New-Object System.Windows.Forms.Form
    $f.Text = "PC Audit Dashboard Settings"
    $f.Size = New-Object System.Drawing.Size(660, 470)
    $f.StartPosition = "CenterParent"
    $f.FormBorderStyle = "FixedDialog"
    $f.MaximizeBox = $false
    $f.MinimizeBox = $false
    $f.BackColor = [System.Drawing.Color]::FromArgb(27,31,36)
    $f.ForeColor = [System.Drawing.Color]::White

    function New-Label($text,$x,$y) {
        $l = New-Object System.Windows.Forms.Label
        $l.Text=$text; $l.Location=New-Object System.Drawing.Point($x,$y); $l.Size=New-Object System.Drawing.Size(140,22)
        $l.ForeColor=[System.Drawing.Color]::Gainsboro
        $f.Controls.Add($l); return $l
    }
    function New-Text($text,$x,$y,$w) {
        $t=New-Object System.Windows.Forms.TextBox
        $t.Text=$text; $t.Location=New-Object System.Drawing.Point($x,$y); $t.Size=New-Object System.Drawing.Size($w,24)
        $t.BackColor=[System.Drawing.Color]::FromArgb(38,43,51); $t.ForeColor=[System.Drawing.Color]::White; $t.BorderStyle='FixedSingle'
        $f.Controls.Add($t); return $t
    }

    New-Label "Dashboard IP / URL Host" 24 28 | Out-Null
    $txtIP = New-Text ([string]$script:GuiSettings.DashboardIP) 190 26 390
    New-Label "Dashboard Host Name" 24 66 | Out-Null
    $txtHost = New-Text ([string]$script:GuiSettings.DashboardHost) 190 64 390
    New-Label "Dashboard Port" 24 104 | Out-Null
    $txtPort = New-Text ([string]$script:GuiSettings.Port) 190 102 110
    New-Label "CSV Base Path" 24 142 | Out-Null
    $txtBase = New-Text ([string]$script:GuiSettings.BasePath) 190 140 390
    New-Label "CSV Master File" 24 180 | Out-Null
    $txtCsv = New-Text ([string]$script:GuiSettings.CsvFileName) 190 178 220
    New-Label "Log Folder" 24 218 | Out-Null
    $txtLog = New-Text ([string]$script:GuiSettings.LogPath) 190 216 390

    $chkBrowser = New-Object System.Windows.Forms.CheckBox
    $chkBrowser.Text="Auto open browser after START"; $chkBrowser.Checked=[bool]$script:GuiSettings.AutoOpenBrowser
    $chkBrowser.Location=New-Object System.Drawing.Point(190,258); $chkBrowser.Size=New-Object System.Drawing.Size(260,24)
    $chkBrowser.ForeColor=[System.Drawing.Color]::White; $chkBrowser.BackColor=$f.BackColor
    $f.Controls.Add($chkBrowser)

    $chkTray = New-Object System.Windows.Forms.CheckBox
    $chkTray.Text="Tray function active / minimize to tray"; $chkTray.Checked=[bool]$script:GuiSettings.MinimizeToTray
    $chkTray.Location=New-Object System.Drawing.Point(190,286); $chkTray.Size=New-Object System.Drawing.Size(300,24)
    $chkTray.ForeColor=[System.Drawing.Color]::White; $chkTray.BackColor=$f.BackColor
    $f.Controls.Add($chkTray)

    $btnSave = New-Object System.Windows.Forms.Button
    $btnSave.Text="SAVE"; $btnSave.Location=New-Object System.Drawing.Point(374,376); $btnSave.Size=New-Object System.Drawing.Size(95,34)
    $btnSave.BackColor=[System.Drawing.Color]::FromArgb(46,204,113); $btnSave.ForeColor=[System.Drawing.Color]::White; $btnSave.FlatStyle='Flat'
    $f.Controls.Add($btnSave)
    $btnCancel = New-Object System.Windows.Forms.Button
    $btnCancel.Text="CANCEL"; $btnCancel.Location=New-Object System.Drawing.Point(486,376); $btnCancel.Size=New-Object System.Drawing.Size(95,34)
    $btnCancel.BackColor=[System.Drawing.Color]::FromArgb(95,101,112); $btnCancel.ForeColor=[System.Drawing.Color]::White; $btnCancel.FlatStyle='Flat'
    $f.Controls.Add($btnCancel)

    $btnCancel.Add_Click({ $f.Close() })
    $btnSave.Add_Click({
        $p = 0
        if (-not [int]::TryParse($txtPort.Text.Trim(), [ref]$p) -or $p -lt 1 -or $p -gt 65535) {
            [System.Windows.Forms.MessageBox]::Show("Enter a valid port number 1-65535.", "Settings", "OK", "Warning") | Out-Null
            return
        }
        $script:GuiSettings.Port = $p
        $script:GuiSettings.DashboardIP = $txtIP.Text.Trim()
        $script:GuiSettings.DashboardHost = $txtHost.Text.Trim()
        $script:GuiSettings.BasePath = $txtBase.Text.Trim()
        $script:GuiSettings.CsvFileName = $txtCsv.Text.Trim()
        if ([string]::IsNullOrWhiteSpace($script:GuiSettings.CsvFileName)) { $script:GuiSettings.CsvFileName = "master_audit.csv" }
        $script:GuiSettings.LogPath = $txtLog.Text.Trim()
        $script:GuiSettings.AutoOpenBrowser = $chkBrowser.Checked
        $script:GuiSettings.MinimizeToTray = $chkTray.Checked
        Save-GuiSettings
        Apply-GuiSettings
        Add-GuiLog "Settings saved. URL=$(Get-GuiDashboardUrl) CSV=$($script:MasterFile)" "INFO"
        Update-ButtonState
        $f.Close()
    })

    $f.ShowDialog($script:MainForm) | Out-Null
}

function Show-HelpWindow {
    # ================================================================================================
    # HELP WINDOW - PROPER ALIGNED SUGGESTIONS AND DETAILS
    # ================================================================================================
    # Purpose:
    # • Opens a clean WinForms help window instead of a plain popup message.
    # • Shows useful suggestions, service details, manual checks, and button guide with emojis.
    # • This does not change original dashboard web/dashboard calculation logic.
    # ================================================================================================

    $help = New-Object System.Windows.Forms.Form
    $help.Text = "❓ Help / Suggestions / Service Details"
    $help.Size = New-Object System.Drawing.Size(720, 560)
    $help.StartPosition = "CenterParent"
    $help.BackColor = [System.Drawing.Color]::FromArgb(18,22,28)
    $help.ForeColor = [System.Drawing.Color]::White
    $help.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $help.FormBorderStyle = "FixedDialog"
    $help.MaximizeBox = $false
    $help.MinimizeBox = $false
    $help.SizeGripStyle = "Hide"

    try {
        $help.Icon = Convert-Base64PngToIcon $script:GuiLogoBase64
    } catch { }

    # ========== TOP HEADER PANEL ==========
    $top = New-Object System.Windows.Forms.Panel
    $top.Location = New-Object System.Drawing.Point(0, 0)
    $top.Size = New-Object System.Drawing.Size(704, 58)
    $top.BackColor = [System.Drawing.Color]::FromArgb(27,32,40)
    $help.Controls.Add($top)

    $title = New-Object System.Windows.Forms.Label
    $title.Text = "❓ PC AUDIT HELP CENTER"
    $title.Font = New-Object System.Drawing.Font("Segoe UI", 13, [System.Drawing.FontStyle]::Bold)
    $title.Location = New-Object System.Drawing.Point(16, 10)
    $title.Size = New-Object System.Drawing.Size(672, 26)
    $title.ForeColor = [System.Drawing.Color]::FromArgb(100,170,225)
    $top.Controls.Add($title)

    $sub = New-Object System.Windows.Forms.Label
    $sub.Text = "Valid service suggestions, details, and manual checks for dashboard support"
    $sub.Font = New-Object System.Drawing.Font("Segoe UI", 8.5)
    $sub.Location = New-Object System.Drawing.Point(18, 36)
    $sub.Size = New-Object System.Drawing.Size(670, 18)
    $sub.ForeColor = [System.Drawing.Color]::FromArgb(165,172,182)
    $top.Controls.Add($sub)

    # ========== TAB CONTROL ==========
    $tabs = New-Object System.Windows.Forms.TabControl
    $tabs.Location = New-Object System.Drawing.Point(12, 70)
    $tabs.Size = New-Object System.Drawing.Size(680, 400)
    $tabs.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $help.Controls.Add($tabs)

    # ========== HELPER: CREATE TAB PAGE ==========
    function New-HelpPage($pageTitle, $content) {
        $page = New-Object System.Windows.Forms.TabPage
        $page.Text = $pageTitle
        $page.BackColor = [System.Drawing.Color]::FromArgb(23,28,36)

        $box = New-Object System.Windows.Forms.RichTextBox
        $box.Multiline = $true
        $box.ReadOnly = $true
        $box.ScrollBars = 'Vertical'
        $box.WordWrap = $true
        $box.BorderStyle = 'FixedSingle'
        $box.Dock = 'Fill'
        $box.BackColor = [System.Drawing.Color]::FromArgb(10,13,18)
        $box.ForeColor = [System.Drawing.Color]::FromArgb(230,235,240)
        
        # ✅ Use Segoe UI Symbol for maximum Unicode compatibility
        $box.Font = New-Object System.Drawing.Font("Segoe UI Symbol", 10)
        $box.Text = $content.Trim()
        
        $box.SelectionStart = 0
        $box.SelectionLength = $box.Text.Length
        $box.SelectionIndent = 8
        $box.SelectionRightIndent = 8
        $box.SelectionStart = 0
        $box.SelectionLength = 0

        $page.Controls.Add($box)
        [void]$tabs.TabPages.Add($page)
    }

    # ========== CONTENT STRINGS WITH UNICODE SYMBOLS ==========
    $serverStatusText = if (Is-DashboardRunning) { 'RUNNING ✓' } else { 'STOPPED ✗' }

    $helpSuggestions = @"
✅ SUGGESTIONS / VALID POINTS
────────────────────────────────────────────────────────────────────────────
1. ▶  Click START only after confirming the CSV source path is reachable.
2. 🔧 Use TEST first if dashboard does not open or data count shows zero.
3. 🌐 OPEN button launches the browser dashboard URL directly.
4. ↻  RESTART is safe when data is not refreshing or listener is stuck.
5. ⏹  STOP releases the HTTP listener and dashboard port.
6. ⚙  SETTINGS is used to change port, host name, CSV path, and log folder.
7. 📝 ACTIVITY LOG shows latest service actions and dashboard log lines.
8. 📌 Keep master_audit.csv file name correct in Settings.
9. 🔒 Run EXE as Administrator if listener/port permission issue appears.
10.⚡ If port is busy, stop old EXE/process or change port in Settings.

⚠️ COMMON ISSUE FIXES
────────────────────────────────────────────────────────────────────────────
• Path empty error      → Check Base Path and Log Path in Settings.
• Dashboard not opening → Check host/IP/port and firewall.
• CSV not found         → Confirm master_audit.csv exists in source path.
• Popup messages in EXE → Use PS2EXE with -noConsole -STA -requireAdmin.
• Icon not showing      → Confirm .ico/.png path exists and is accessible.
"@

    $buttonGuide = @"
▶️ BUTTON GUIDE
────────────────────────────────────────────────────────────────────────────
▶ START        → Starts the original PC Audit Dashboard HTTP server.
⏹ STOP         → Stops the server and releases listener/port.
↻ RESTART      → Stops and starts the service again safely.
🔧 TEST        → Checks source path, CSV file, log folder, port, and URL.
🌐 OPEN        → Opens dashboard URL in browser.
⚙ SETTINGS     → Edit port, host, IP, base path, CSV name, and log location.
❓ HELP         → Opens this help window with suggestions and service details.
📌 TRAY        → Minimize keeps controller running in system tray.
🚪 EXIT         → Stops server first, then closes the controller.
"@

    $serviceDetails = @"
📊 CURRENT SERVICE DETAILS
────────────────────────────────────────────────────────────────────────────
Service Status        → $serverStatusText
Application Version → $($script:GuiVersion)
Port                    → $Port
Dashboard URL       → $(Get-GuiDashboardUrl)
Dashboard Host      → $($script:DashboardHost)
Dashboard IP          → $($script:DashboardIP)
CSV File Name         → $($script:CsvFileName)
CSV Source Path      → $($script:MasterFile)
Base Path           → $($script:GuiSettings.BasePath)
Log Folder          → $($script:LogPath)
Current Log File    → $($script:LogFile)
PowerShell Version  → $($PSVersionTable.PSVersion)
Started At          → $(if ($script:ServerStartedAt) { $script:ServerStartedAt } else { '-' })
Tray Enabled        → $($script:GuiSettings.MinimizeToTray)
Auto Open Browser   → $($script:GuiSettings.AutoOpenBrowser)
"@

    $manualChecks = @"
🛠️ USEFUL MANUAL CHECKS
────────────────────────────────────────────────────────────────────────────
Test-Path "$($script:GuiSettings.BasePath)"
Test-Path "$($script:MasterFile)"
Test-Path "$($script:LogPath)"
Get-Content "$($script:LogFile)" -Tail 50
netstat -ano | findstr :$Port

✅ RECOMMENDED PS2EXE COMMAND
────────────────────────────────────────────────────────────────────────────
ps2exe .\PC_Audit_Dashboard_Service_EXE_FIXED.ps1 .\PC_Audit_Dashboard_Service.exe -noConsole -STA -requireAdmin

📝 NOTE
────────────────────────────────────────────────────────────────────────────
This Help window is only GUI support. It does not modify original dashboard HTML,
CSV import, statistics, issue cards, or web dashboard calculation logic.
"@

    # Create tabs
    New-HelpPage "✅ Suggestions" $helpSuggestions
    New-HelpPage "▶️ Buttons" $buttonGuide
    New-HelpPage "📊 Details" $serviceDetails
    New-HelpPage "🛠️ Checks" $manualChecks

    # ========== CLOSE BUTTON ==========
    $btnClose = New-Object System.Windows.Forms.Button
    $btnClose.Text = "CLOSE"
    $btnClose.Size = New-Object System.Drawing.Size(95, 32)
    $btnClose.Location = New-Object System.Drawing.Point(597, 480)
    $btnClose.BackColor = [System.Drawing.Color]::FromArgb(92,101,112)
    $btnClose.ForeColor = [System.Drawing.Color]::White
    $btnClose.FlatStyle = 'Flat'
    $btnClose.FlatAppearance.BorderSize = 0
    $btnClose.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Bold)
    $btnClose.Add_Click({ $help.Close() })
    $help.Controls.Add($btnClose)

    # ========== BOTTOM HINT LABEL ==========
    $hint = New-Object System.Windows.Forms.Label
    $hint.Text = "Click tabs to view different help sections"
    $hint.Location = New-Object System.Drawing.Point(12, 486)
    $hint.Size = New-Object System.Drawing.Size(360, 22)
    $hint.ForeColor = [System.Drawing.Color]::FromArgb(165,172,182)
    $help.Controls.Add($hint)

    $help.ShowDialog($script:MainForm) | Out-Null
}



# ====================================================================================================
# GUI LOGO / ICON Base64 SETTINGS
# ====================================================================================================
# Add your .png base64 code added below.
# Login title icon uses $script:GuiLogoBase64.
# Main GUI header logo uses $script:GuiLogoBase64.
# Window/tray icon uses $script:GuiLogoBase64.
# Keep the base code unchange or modify '$script:GuiLogoBase64'
# ====================================================================================================
# Base64 code for GUI
$script:GuiLogoBase64 = @"
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
"@

function Set-Base64PictureImage {
    param(
        [Parameter(Mandatory=$true)] [System.Windows.Forms.PictureBox]$PictureBox,
        [string]$Base64
    )
    try {
        $bytes = [Convert]::FromBase64String($Base64)
        $ms = New-Object System.IO.MemoryStream(,$bytes)
        $PictureBox.Image = [System.Drawing.Image]::FromStream($ms)
        $PictureBox.SizeMode = [System.Windows.Forms.PictureBoxSizeMode]::Zoom
    } catch { }
}

function Convert-Base64PngToIcon {
    param([string]$Base64)

    $bytes = [Convert]::FromBase64String($Base64)
    $ms = New-Object System.IO.MemoryStream(,$bytes)
    $bmp = New-Object System.Drawing.Bitmap($ms)
    $hIcon = $bmp.GetHicon()

    return [System.Drawing.Icon]::FromHandle($hIcon)
}

function Show-LoginWindow {
    $login = New-Object System.Windows.Forms.Form
    $login.Text = "🔑 PC Audit Login"
    $login.Size = New-Object System.Drawing.Size(455, 285)
    $login.StartPosition = "CenterScreen"
    $login.FormBorderStyle = "FixedDialog"
    $login.MaximizeBox = $false
    $login.BackColor = [System.Drawing.Color]::FromArgb(22,25,30)
    $login.ForeColor = [System.Drawing.Color]::White

    $title = New-Object System.Windows.Forms.Label
    $loginLogo = New-Object System.Windows.Forms.PictureBox
    $loginLogo.Location = New-Object System.Drawing.Point(34, 23)
    $loginLogo.Size = New-Object System.Drawing.Size(34, 34)
    $loginLogo.BackColor = [System.Drawing.Color]::Transparent
    Set-Base64PictureImage -PictureBox $loginLogo -Base64 $script:GuiLogoBase64
    $login.Controls.Add($loginLogo)

    $title.Text = "PC AUDIT DASHBOARD SERVICE CONTROLLER"
    $title.Font = New-Object System.Drawing.Font("Segoe UI", 14, [System.Drawing.FontStyle]::Bold)
    $title.Location = New-Object System.Drawing.Point(73, 24)
    $title.Size = New-Object System.Drawing.Size(350, 35)
    $title.TextAlign = "MiddleLeft"
    $title.ForeColor = [System.Drawing.Color]::FromArgb(86,156,214)
    $login.Controls.Add($title)

    $lblUser = New-Object System.Windows.Forms.Label
    $lblUser.Text="👥Username"; $lblUser.Location=New-Object System.Drawing.Point(55,88); $lblUser.Size=New-Object System.Drawing.Size(90,24)
    $login.Controls.Add($lblUser)
    $txtUser = New-Object System.Windows.Forms.TextBox
    $txtUser.Location=New-Object System.Drawing.Point(150,86); $txtUser.Size=New-Object System.Drawing.Size(200,25); $txtUser.Text="HH0010520"
    $login.Controls.Add($txtUser)

    $lblPass = New-Object System.Windows.Forms.Label
    $lblPass.Text="🔐Password"; $lblPass.Location=New-Object System.Drawing.Point(55,127); $lblPass.Size=New-Object System.Drawing.Size(90,24)
    $login.Controls.Add($lblPass)
    $txtPass = New-Object System.Windows.Forms.TextBox
    $txtPass.Location=New-Object System.Drawing.Point(150,125); $txtPass.Size=New-Object System.Drawing.Size(200,25); $txtPass.UseSystemPasswordChar=$true
    $login.Controls.Add($txtPass)

    $btnLogin = New-Object System.Windows.Forms.Button
    $btnLogin.Text="LOGIN"; $btnLogin.Location=New-Object System.Drawing.Point(150,176); $btnLogin.Size=New-Object System.Drawing.Size(95,34)
    $btnLogin.BackColor=[System.Drawing.Color]::FromArgb(46,204,113); $btnLogin.ForeColor=[System.Drawing.Color]::White; $btnLogin.FlatStyle='Flat'
    $login.Controls.Add($btnLogin)
    $btnExit = New-Object System.Windows.Forms.Button
    $btnExit.Text="EXIT"; $btnExit.Location=New-Object System.Drawing.Point(255,176); $btnExit.Size=New-Object System.Drawing.Size(95,34)
    $btnExit.BackColor=[System.Drawing.Color]::FromArgb(231,76,60); $btnExit.ForeColor=[System.Drawing.Color]::White; $btnExit.FlatStyle='Flat'
    $login.Controls.Add($btnExit)

    $script:LoginOk = $false
    $btnExit.Add_Click({ $login.Close() })
    $btnLogin.Add_Click({
        if ($txtUser.Text.Trim() -eq "HH0010520" -and $txtPass.Text -eq "Foxconn-FXCN-IT") {
            $script:LoginOk = $true
            $login.Close()
        } else {
            [System.Windows.Forms.MessageBox]::Show("Invalid username or password.", "Login Failed", "OK", "Warning") | Out-Null
            $txtPass.Clear(); $txtPass.Focus()
        }
    })
    $txtPass.Add_KeyDown({ if ($_.KeyCode -eq 'Enter') { $btnLogin.PerformClick() } })
    $login.ShowDialog() | Out-Null
    return $script:LoginOk
}

function Show-MainGui {
    $form = New-Object System.Windows.Forms.Form
    $script:MainForm = $form
    # EXE FIX: Load icon only if the Base64 code exists. Missing Base64 code should not crash converted EXE.
    $GuiIcon = Convert-Base64PngToIcon $script:GuiLogoBase64
    $form.Icon = $GuiIcon
    # To remove/change GUI title icon, edit $iconPath above.
    $form.Text = "PC Audit Dashboard Service Controller"
    $form.Size = New-Object System.Drawing.Size(620, 455)
    $form.MinimumSize = New-Object System.Drawing.Size(620, 455)
    # EXPAND VIEW: Window is now resizable. Do not set MaximumSize.
    $form.StartPosition = "CenterScreen"
    $form.FormBorderStyle = "Sizable"
    $form.SizeGripStyle = "Show"
    $form.MaximizeBox = $true
    $form.BackColor = [System.Drawing.Color]::FromArgb(18,22,28)
    $form.ForeColor = [System.Drawing.Color]::White
    $form.Font = New-Object System.Drawing.Font("Segoe UI", 8.5)

    # ================================================================================================
    # COMPACT PROFESSIONAL SERVICE CONTROLLER LAYOUT
    # ================================================================================================
    # Purpose:
    # • Keep GUI small and service-controller style.
    # • Browser dashboard remains unchanged; this window only controls and monitors the service.
    # ================================================================================================

    $header = New-Object System.Windows.Forms.Panel
    $header.Location = New-Object System.Drawing.Point(0,0)
    $header.Size = New-Object System.Drawing.Size(620,54)
    $header.Anchor = "Top,Left,Right"
    $header.BackColor = [System.Drawing.Color]::FromArgb(27,32,40)
    $form.Controls.Add($header)

    # Header logo placeholder - add your image path in $script:HeaderLogoPath above.
    $headerLogo = New-Object System.Windows.Forms.PictureBox
    $headerLogo.Location = New-Object System.Drawing.Point(10,10)
    $headerLogo.Size = New-Object System.Drawing.Size(34,34)
    $headerLogo.BackColor = [System.Drawing.Color]::Transparent
    Set-Base64PictureImage -PictureBox $headerLogo -Base64 $script:GuiLogoBase64
    $header.Controls.Add($headerLogo)

    $btnTopHelp = New-Object System.Windows.Forms.Button
    $btnTopHelp.Text = "?"
    $btnTopHelp.Location = New-Object System.Drawing.Point(548,12)
    $btnTopHelp.Anchor = "Top,Right"
    $btnTopHelp.Size = New-Object System.Drawing.Size(30,28)
    $btnTopHelp.BackColor = [System.Drawing.Color]::FromArgb(58,64,74)
    $btnTopHelp.ForeColor = [System.Drawing.Color]::White
    $btnTopHelp.FlatStyle = 'Flat'
    $btnTopHelp.FlatAppearance.BorderSize = 0
    $btnTopHelp.Font = New-Object System.Drawing.Font("Segoe UI", 10, [System.Drawing.FontStyle]::Bold)
    $btnTopHelp.Add_Click({ Show-HelpWindow })
    $header.Controls.Add($btnTopHelp)

    $title = New-Object System.Windows.Forms.Label
    $title.Text = "PC AUDIT DASHBOARD SERVICE CONTROLLER"
    $title.Font = New-Object System.Drawing.Font("Segoe UI", 12.5, [System.Drawing.FontStyle]::Bold)
    $title.Location = New-Object System.Drawing.Point(54,9)
    $title.Size = New-Object System.Drawing.Size(485,22)
    $title.Anchor = "Top,Left,Right"
    $title.ForeColor = [System.Drawing.Color]::FromArgb(210,220,235)
    $header.Controls.Add($title)

    $subTitle = New-Object System.Windows.Forms.Label
    $subTitle.Text = "Compact service manager for dashboard start / stop / monitor"
    $subTitle.Font = New-Object System.Drawing.Font("Segoe UI", 8)
    $subTitle.Location = New-Object System.Drawing.Point(55,31)
    $subTitle.Size = New-Object System.Drawing.Size(470,18)
    $subTitle.Anchor = "Top,Left,Right"
    $subTitle.ForeColor = [System.Drawing.Color]::FromArgb(145,155,168)
    $header.Controls.Add($subTitle)

    # Status strip inside top content area
    $statusPanel = New-Object System.Windows.Forms.Panel
    $statusPanel.Location = New-Object System.Drawing.Point(10,62)
    $statusPanel.Size = New-Object System.Drawing.Size(584,58)
    $statusPanel.Anchor = "Top,Left,Right"
    $statusPanel.BackColor = [System.Drawing.Color]::FromArgb(23,28,36)
    $statusPanel.BorderStyle = 'FixedSingle'
    $form.Controls.Add($statusPanel)

    $script:StatusLabel = New-Object System.Windows.Forms.Label
    $script:StatusLabel.Font = New-Object System.Drawing.Font("Segoe UI", 11, [System.Drawing.FontStyle]::Bold)
    $script:StatusLabel.Location = New-Object System.Drawing.Point(12,7)
    $script:StatusLabel.Size = New-Object System.Drawing.Size(145,22)
    $statusPanel.Controls.Add($script:StatusLabel)

    $script:DetailLabel = New-Object System.Windows.Forms.Label
    $script:DetailLabel.Font = New-Object System.Drawing.Font("Segoe UI", 8.5)
    $script:DetailLabel.Location = New-Object System.Drawing.Point(12,31)
    $script:DetailLabel.Size = New-Object System.Drawing.Size(560,20)
    $script:DetailLabel.ForeColor = [System.Drawing.Color]::Gainsboro
    $statusPanel.Controls.Add($script:DetailLabel)

    # Compact toolbar
    $toolbar = New-Object System.Windows.Forms.Panel
    $toolbar.Location = New-Object System.Drawing.Point(10,127)
    $toolbar.Size = New-Object System.Drawing.Size(584,43)
    $toolbar.Anchor = "Top,Left,Right"
    $toolbar.BackColor = [System.Drawing.Color]::FromArgb(18,22,28)
    $form.Controls.Add($toolbar)

    function New-CompactButton($text, $x, $w, $color) {
        $b = New-Object System.Windows.Forms.Button
        $b.Text = $text
        $b.Location = New-Object System.Drawing.Point($x,5)
        $b.Size = New-Object System.Drawing.Size($w,30)
        $b.BackColor = $color
        $b.ForeColor = [System.Drawing.Color]::White
        $b.FlatStyle = 'Flat'
        $b.FlatAppearance.BorderSize = 0
        $b.Font = New-Object System.Drawing.Font("Segoe UI", 8.2, [System.Drawing.FontStyle]::Bold)
        $toolbar.Controls.Add($b)
        return $b
    }

    $script:BtnStart   = New-CompactButton "START"    0   72 ([System.Drawing.Color]::FromArgb(36,155,86))
    $script:BtnStop    = New-CompactButton "STOP"     78  72 ([System.Drawing.Color]::FromArgb(174,50,44))
    $script:BtnRestart = New-CompactButton "RESTART"  156 82 ([System.Drawing.Color]::FromArgb(198,111,34))
    $script:BtnTest           = New-CompactButton "TEST"     244 72 ([System.Drawing.Color]::FromArgb(105,75,165))
    $script:BtnOpen           = New-CompactButton "OPEN"     322 72 ([System.Drawing.Color]::FromArgb(42,125,185))
    $script:BtnSettings       = New-CompactButton "SETTINGS" 400 86 ([System.Drawing.Color]::FromArgb(63,72,84))
    $script:BtnExit           = New-CompactButton "EXIT"     492 72 ([System.Drawing.Color]::FromArgb(92,101,112))

    # Service information, kept short and aligned
    $infoPanel = New-Object System.Windows.Forms.Panel
    $infoPanel.Location = New-Object System.Drawing.Point(10,176)
    $infoPanel.Size = New-Object System.Drawing.Size(584,78)
    $infoPanel.Anchor = "Top,Left,Right"
    $infoPanel.BackColor = [System.Drawing.Color]::FromArgb(23,28,36)
    $infoPanel.BorderStyle = 'FixedSingle'
    $form.Controls.Add($infoPanel)

    $infoTitle = New-Object System.Windows.Forms.Label
    $infoTitle.Text = "ℹ️ SERVICE INFORMATION"
    $infoTitle.Font = New-Object System.Drawing.Font("Segoe UI", 8.8, [System.Drawing.FontStyle]::Bold)
    $infoTitle.Location = New-Object System.Drawing.Point(10,6)
    $infoTitle.Size = New-Object System.Drawing.Size(230,18)
    $infoTitle.ForeColor = [System.Drawing.Color]::FromArgb(100,170,225)
    $infoPanel.Controls.Add($infoTitle)

    $script:UptimeLabel = New-Object System.Windows.Forms.Label
    $script:UptimeLabel.Text = "Host: -  |  Port: -  |  CSV: -"
    $script:UptimeLabel.Font = New-Object System.Drawing.Font("Segoe UI", 8.3)
    $script:UptimeLabel.Location = New-Object System.Drawing.Point(10,28)
    $script:UptimeLabel.Size = New-Object System.Drawing.Size(560,18)
    $script:UptimeLabel.Anchor = "Top,Left,Right"
    $script:UptimeLabel.ForeColor = [System.Drawing.Color]::Gainsboro
    $infoPanel.Controls.Add($script:UptimeLabel)

    $script:MiniInfoLabel = New-Object System.Windows.Forms.Label
    $script:MiniInfoLabel.Text = "Base Path: -"
    $script:MiniInfoLabel.Font = New-Object System.Drawing.Font("Segoe UI", 8.1)
    $script:MiniInfoLabel.Location = New-Object System.Drawing.Point(10,50)
    $script:MiniInfoLabel.Size = New-Object System.Drawing.Size(560,18)
    $script:MiniInfoLabel.Anchor = "Top,Left,Right"
    $script:MiniInfoLabel.ForeColor = [System.Drawing.Color]::FromArgb(165,172,182)
    $infoPanel.Controls.Add($script:MiniInfoLabel)

    $logTitle = New-Object System.Windows.Forms.Label
    $logTitle.Text = "📝 ACTIVITY LOG"
    $logTitle.Font = New-Object System.Drawing.Font("Segoe UI", 8.8, [System.Drawing.FontStyle]::Bold)
    $logTitle.Location = New-Object System.Drawing.Point(12,262)
    $logTitle.Size = New-Object System.Drawing.Size(300,18)
    $logTitle.ForeColor = [System.Drawing.Color]::FromArgb(100,170,225)
    $form.Controls.Add($logTitle)

    $script:LogBox = New-Object System.Windows.Forms.TextBox
    $script:LogBox.Location = New-Object System.Drawing.Point(10,283)
    $script:LogBox.Size = New-Object System.Drawing.Size(584,102)
    $script:LogBox.Anchor = "Top,Bottom,Left,Right"
    $script:LogBox.Multiline = $true
    $script:LogBox.ScrollBars = 'Vertical'
    $script:LogBox.ReadOnly = $true
    $script:LogBox.WordWrap = $false
    $script:LogBox.BackColor = [System.Drawing.Color]::FromArgb(10,13,18)
    $script:LogBox.ForeColor = [System.Drawing.Color]::FromArgb(220,220,220)
    $script:LogBox.BorderStyle = 'FixedSingle'
    $script:LogBox.Font = New-Object System.Drawing.Font("Consolas", 8.2)
    $form.Controls.Add($script:LogBox)

    $script:FooterLabel = New-Object System.Windows.Forms.Label
    $script:FooterLabel.Location = New-Object System.Drawing.Point(0,394)
    $script:FooterLabel.Size = New-Object System.Drawing.Size(620,22)
    $script:FooterLabel.Anchor = "Bottom,Left,Right"
    $script:FooterLabel.BackColor = [System.Drawing.Color]::FromArgb(27,32,40)
    $script:FooterLabel.ForeColor = [System.Drawing.Color]::FromArgb(170,178,188)
    $script:FooterLabel.Font = New-Object System.Drawing.Font("Segoe UI", 8)
    $script:FooterLabel.TextAlign = 'MiddleLeft'
    $script:FooterLabel.Padding = New-Object System.Windows.Forms.Padding(10,0,0,0)
    $script:FooterLabel.Text = "Stopped | Tray Active | Build $($script:GuiVersion)"
    $form.Controls.Add($script:FooterLabel)

# ================================================================================================
# EXPAND VIEW - BOTTOM RIGHT DRAG GRIP
# ================================================================================================
$script:ResizeDragActive = $false
$script:ResizeStartMouse = [System.Drawing.Point]::Empty
$script:ResizeStartSize = [System.Drawing.Size]::Empty

$resizeGripRight = New-Object System.Windows.Forms.Label
$resizeGripRight.Text = "◢"
$resizeGripRight.Location = New-Object System.Drawing.Point(584, 400)    # ← CHANGED: X = 620-36 = 584, Y adjusted
$resizeGripRight.Size = New-Object System.Drawing.Size(30, 18)         # ← CHANGED: 58% smaller (was 84×22)
$resizeGripRight.Anchor = "Bottom,Right"                                 # ← CHANGED: Right anchor
$resizeGripRight.BackColor = [System.Drawing.Color]::FromArgb(36,42,52)
$resizeGripRight.ForeColor = [System.Drawing.Color]::FromArgb(160, 170, 185)  # ← CHANGED: Softer color
$resizeGripRight.Font = New-Object System.Drawing.Font("Segoe UI Symbol", 9, [System.Drawing.FontStyle]::Regular)  # ← CHANGED: Symbol font, smaller
$resizeGripRight.TextAlign = "MiddleCenter"  # ← CHANGED: Align to corner
$resizeGripRight.Cursor = [System.Windows.Forms.Cursors]::SizeNWSE        # ← CHANGED: NWSE cursor
$form.Controls.Add($resizeGripRight)
$resizeGripRight.BringToFront()

$resizeGripRight.Add_MouseDown({
    if ($_.Button -eq [System.Windows.Forms.MouseButtons]::Left) {
        $script:ResizeDragActive = $true
        $script:ResizeStartMouse = [System.Windows.Forms.Cursor]::Position
        $script:ResizeStartSize = $form.Size
    }
})

$resizeGripRight.Add_MouseMove({
    if ($script:ResizeDragActive) {
        $currentMouse = [System.Windows.Forms.Cursor]::Position
        $dx = $currentMouse.X - $script:ResizeStartMouse.X    # ← CHANGED: reversed for right-side drag
        $dy = $currentMouse.Y - $script:ResizeStartMouse.Y
        $newWidth = [Math]::Max($form.MinimumSize.Width, $script:ResizeStartSize.Width + $dx)
        $newHeight = [Math]::Max($form.MinimumSize.Height, $script:ResizeStartSize.Height + $dy)
        $form.Size = New-Object System.Drawing.Size($newWidth, $newHeight)
    }
})

$resizeGripRight.Add_MouseUp({ $script:ResizeDragActive = $false })
$form.Add_MouseUp({ $script:ResizeDragActive = $false })



    # ================================================================================================
    # AUTO-LOCK OVERLAY - shows after 5 minutes of PC inactivity
    # ================================================================================================
    $script:LockPanel = New-Object System.Windows.Forms.Panel
    $script:LockPanel.Location = New-Object System.Drawing.Point(0,0)
    $script:LockPanel.Size = New-Object System.Drawing.Size(620,455)
    $script:LockPanel.Anchor = "Top,Bottom,Left,Right"
    $script:LockPanel.BackColor = [System.Drawing.Color]::FromArgb(14,18,24)
    $script:LockPanel.Visible = $false

    $lockTitle = New-Object System.Windows.Forms.Label
    $lockTitle.Text = "🔒 CONTROLLER LOCKED"
    $lockTitle.Font = New-Object System.Drawing.Font("Segoe UI", 18, [System.Drawing.FontStyle]::Bold)
    $lockTitle.ForeColor = [System.Drawing.Color]::FromArgb(241,196,15)
    $lockTitle.TextAlign = "MiddleCenter"
    $lockTitle.Location = New-Object System.Drawing.Point(0,70)
    $lockTitle.Size = New-Object System.Drawing.Size(620,42)
    $lockTitle.Anchor = "Top,Left,Right"
    $script:LockPanel.Controls.Add($lockTitle)

    $lockSub = New-Object System.Windows.Forms.Label
    $lockSub.Text = "Server is still running. Unlock is required to Stop / Restart / Settings / Exit."
    $lockSub.Font = New-Object System.Drawing.Font("Segoe UI", 9.5)
    $lockSub.ForeColor = [System.Drawing.Color]::FromArgb(190,198,210)
    $lockSub.TextAlign = "MiddleCenter"
    $lockSub.Location = New-Object System.Drawing.Point(0,116)
    $lockSub.Size = New-Object System.Drawing.Size(620,26)
    $lockSub.Anchor = "Top,Left,Right"
    $script:LockPanel.Controls.Add($lockSub)

    $lblLockUser = New-Object System.Windows.Forms.Label
    $lblLockUser.Text = "Username"
    $lblLockUser.Location = New-Object System.Drawing.Point(175,176)
    $lblLockUser.Size = New-Object System.Drawing.Size(90,24)
    $lblLockUser.ForeColor = [System.Drawing.Color]::White
    $script:LockPanel.Controls.Add($lblLockUser)

    $script:LockUserBox = New-Object System.Windows.Forms.TextBox
    $script:LockUserBox.Location = New-Object System.Drawing.Point(270,174)
    $script:LockUserBox.Size = New-Object System.Drawing.Size(185,25)
    $script:LockPanel.Controls.Add($script:LockUserBox)

    $lblLockPass = New-Object System.Windows.Forms.Label
    $lblLockPass.Text = "Password"
    $lblLockPass.Location = New-Object System.Drawing.Point(175,216)
    $lblLockPass.Size = New-Object System.Drawing.Size(90,24)
    $lblLockPass.ForeColor = [System.Drawing.Color]::White
    $script:LockPanel.Controls.Add($lblLockPass)

    $script:LockPassBox = New-Object System.Windows.Forms.TextBox
    $script:LockPassBox.Location = New-Object System.Drawing.Point(270,214)
    $script:LockPassBox.Size = New-Object System.Drawing.Size(185,25)
    $script:LockPassBox.UseSystemPasswordChar = $true
    $script:LockPanel.Controls.Add($script:LockPassBox)

    $btnUnlock = New-Object System.Windows.Forms.Button
    $btnUnlock.Text = "UNLOCK"
    $btnUnlock.Location = New-Object System.Drawing.Point(270,258)
    $btnUnlock.Size = New-Object System.Drawing.Size(185,34)
    $btnUnlock.BackColor = [System.Drawing.Color]::FromArgb(46,204,113)
    $btnUnlock.ForeColor = [System.Drawing.Color]::White
    $btnUnlock.FlatStyle = 'Flat'
    $script:LockPanel.Controls.Add($btnUnlock)

    $script:LockMessageLabel = New-Object System.Windows.Forms.Label
    $script:LockMessageLabel.Text = "Auto locked after 5 minutes of inactivity."
    $script:LockMessageLabel.Font = New-Object System.Drawing.Font("Segoe UI", 8.5)
    $script:LockMessageLabel.ForeColor = [System.Drawing.Color]::FromArgb(220,220,220)
    $script:LockMessageLabel.TextAlign = "MiddleCenter"
    $script:LockMessageLabel.Location = New-Object System.Drawing.Point(0,304)
    $script:LockMessageLabel.Size = New-Object System.Drawing.Size(620,24)
    $script:LockMessageLabel.Anchor = "Top,Left,Right"
    $script:LockPanel.Controls.Add($script:LockMessageLabel)

    $btnUnlock.Add_Click({ Unlock-Controller | Out-Null })
    $script:LockPassBox.Add_KeyDown({ if ($_.KeyCode -eq 'Enter') { Unlock-Controller | Out-Null } })
    $form.Controls.Add($script:LockPanel)
    $script:LockPanel.BringToFront()
    
    # Tray setup - active system tray controller
    $script:NotifyIcon = New-Object System.Windows.Forms.NotifyIcon
    # EXE FIX: Load tray icon safely. if  base64 code not work use default application icon.
    $script:NotifyIcon.Icon = $GuiIcon
    $script:NotifyIcon.Visible = $true
    $trayMenu = New-Object System.Windows.Forms.ContextMenuStrip
    $miOpenManager = $trayMenu.Items.Add("Show Controller")
    $miOpenDashboard = $trayMenu.Items.Add("Open Dashboard")
    $script:MiStart = $trayMenu.Items.Add("Start Server")
    $script:MiStop = $trayMenu.Items.Add("Stop Server")
    $script:MiRestart = $trayMenu.Items.Add("Restart Server")
    $script:MiTest = $trayMenu.Items.Add("Quick Test")
    $script:MiSettings = $trayMenu.Items.Add("Settings")
    [void]$trayMenu.Items.Add("-")
    $script:MiExit = $trayMenu.Items.Add("Exit")
    $script:NotifyIcon.ContextMenuStrip = $trayMenu
    $script:NotifyIcon.Text = "PC Audit Dashboard Controller"
    $script:NotifyIcon.Add_DoubleClick({ $form.Show(); $form.WindowState = 'Normal'; $form.Activate() })
    $miOpenManager.Add_Click({ $form.Show(); $form.WindowState = 'Normal'; $form.Activate() })
    $miOpenDashboard.Add_Click({ Open-GuiDashboard })
    $script:MiStart.Add_Click({ if (Require-ControllerUnlock) { Start-GuiDashboard } })
    $script:MiStop.Add_Click({ if (Require-ControllerUnlock) { Stop-GuiDashboard } })
    $script:MiRestart.Add_Click({ if (Require-ControllerUnlock) { Restart-GuiDashboard } })
    $script:MiTest.Add_Click({ if (Require-ControllerUnlock) { Test-GuiDashboardUrl } })
    $script:MiSettings.Add_Click({ if (Require-ControllerUnlock) { $form.Show(); $form.WindowState = 'Normal'; $form.Activate(); Show-SettingsWindow } })
    $script:MiExit.Add_Click({ if (Require-ControllerUnlock) { Stop-GuiDashboard; $script:NotifyIcon.Visible=$false; $form.Close() } })

    $script:BtnStart.Add_Click({ if (Require-ControllerUnlock) { Start-GuiDashboard } })
    $script:BtnStop.Add_Click({ if (Require-ControllerUnlock) { Stop-GuiDashboard } })
    $script:BtnRestart.Add_Click({ if (Require-ControllerUnlock) { Restart-GuiDashboard } })
    $script:BtnOpen.Add_Click({ Open-GuiDashboard })
    $script:BtnTest.Add_Click({ if (Require-ControllerUnlock) { Test-GuiDashboardUrl } })
    $script:BtnSettings.Add_Click({ if (Require-ControllerUnlock) { Show-SettingsWindow } })
    $script:BtnExit.Add_Click({ if (Require-ControllerUnlock) { Stop-GuiDashboard; $script:NotifyIcon.Visible=$false; $form.Close() } })

    $form.Add_Resize({
        if ($form.WindowState -eq 'Minimized' -and $script:GuiSettings.MinimizeToTray) {
            $form.Hide()
            $script:NotifyIcon.ShowBalloonTip(1200, "PC Audit Dashboard", "Controller minimized to tray.", [System.Windows.Forms.ToolTipIcon]::Info)
        }
    })

    $form.Add_FormClosing({
        if (Is-DashboardRunning) {
            $choice = [System.Windows.Forms.MessageBox]::Show("Dashboard is running. Stop server and exit?", "Exit", "YesNo", "Question")
            if ($choice -ne [System.Windows.Forms.DialogResult]::Yes) { $_.Cancel = $true; return }
        }
        Stop-GuiDashboard
        if ($script:NotifyIcon) { $script:NotifyIcon.Visible = $false; $script:NotifyIcon.Dispose() }
    })

    $timer = New-Object System.Windows.Forms.Timer
    $timer.Interval = 1000
    $timer.Add_Tick({
        Process-GuiDashboardRequests
        if (-not $script:ControllerLocked -and (Get-ControllerIdleSeconds) -ge $script:AutoLockSeconds) {
            Lock-Controller "Auto locked after $($script:AutoLockMinutes) minutes of inactivity."
        }
        Update-ButtonState

        $running = Is-DashboardRunning
        $statusWord = if ($running) { "Running" } else { "Stopped" }
        $trayWord = if ($script:GuiSettings.MinimizeToTray) { "Tray Active" } else { "Tray Off" }

        if ($running -and $script:ServerStartedAt) {
            $up = (Get-Date) - $script:ServerStartedAt
            $script:UptimeLabel.Text = "Host: {0}  |  Port: {1}  |  Uptime: {2:00}:{3:00}:{4:00}  |  CSV: {5}" -f $script:DashboardHost, $Port, [int]$up.TotalHours, $up.Minutes, $up.Seconds, $script:CsvFileName
            $script:FooterLabel.Text = "Running | $trayWord | Build $($script:GuiVersion) | Last check: $(Get-Date -Format 'HH:mm:ss')"
        } else {
            $script:UptimeLabel.Text = "Host: $($script:DashboardHost)  |  Port: $Port  |  Uptime: -  |  CSV: $($script:CsvFileName)"
            $script:FooterLabel.Text = "Stopped | $trayWord | Build $($script:GuiVersion) | Last check: $(Get-Date -Format 'HH:mm:ss')"
        }

        if ($script:MiniInfoLabel) {
            $script:MiniInfoLabel.Text = "Base Path: $($script:GuiSettings.BasePath)"
        }

        # Pull new lines from the real dashboard log file into Activity Log.
        if ($script:GuiSettings.AutoRefreshLog -and (Test-Path $script:LogFile)) {
            try {
                $fi = Get-Item $script:LogFile
                if ($fi.Length -lt $script:LastLogLength) { $script:LastLogLength = 0 }
                if ($fi.Length -gt $script:LastLogLength) {
                    $fs = [System.IO.File]::Open($script:LogFile, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
                    $fs.Seek($script:LastLogLength, [System.IO.SeekOrigin]::Begin) | Out-Null
                    $sr = New-Object System.IO.StreamReader($fs)
                    $newText = $sr.ReadToEnd()
                    $script:LastLogLength = $fs.Position
                    $sr.Close(); $fs.Close()
                    if ($newText) { $script:LogBox.AppendText($newText.TrimEnd() + [Environment]::NewLine); $script:LogBox.ScrollToCaret() }
                }
            } catch { }
        }
    })
    $timer.Start()

    Add-GuiLog "Compact service controller loaded. Login authorized. Click START to run dashboard." "INFO"
    Update-ButtonState
    [System.Windows.Forms.Application]::Run($form)
}

# Register shutdown handler for clean exit
$null = Register-EngineEvent -SourceIdentifier ([System.Management.Automation.PsEngineEvent]::Exiting) -Action { Stop-Server }

# ====================================================================================================
# EXE CONVERSION NOTES
# ====================================================================================================
# Recommended PS2EXE options:
#   ps2exe .\PC_Audit_Dashboard_Service_EXE_FIXED.ps1 .\PC_Audit_Dashboard_Service.exe -noConsole -STA -requireAdmin
#
# Why these fixes are added:
#   1. $PSScriptRoot can be empty in EXE mode, so settings path now uses $script:ExeSafeScriptRoot.
#   2. Write-Host / Write-Warning lines are commented because they can appear as popup message boxes.
#   3. Missing icon file no longer crashes the GUI.
#   4. Empty BasePath / LogPath values are replaced with safe defaults before Join-Path runs.
# ====================================================================================================

try {
    Initialize-GuiSettings
    if (Show-LoginWindow) {
        Show-MainGui
    }
}
finally {
    Stop-Server
    try { Write-Log -Message "=== PC Audit Dashboard GUI Closed ===" -Level 'INFO' } catch { }
}
