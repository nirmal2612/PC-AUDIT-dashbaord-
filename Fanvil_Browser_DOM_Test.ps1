# ============================================================
# FANVIL SINGLE PHONE BROWSER-DOM TEST
# Purpose:
#   When Invoke-WebRequest login works but cannot read phone data,
#   this tool uses an embedded browser page and reads the visible DOM text.
#
# How to use:
#   1. Enter phone IP
#   2. Click Open Phone
#   3. If login page opens, login normally OR click Auto Login
#   4. Open the Information page in the embedded browser
#   5. Click Fetch Visible Data
#
# Output:
#   D:\fanvil\Fanvil_Browser_DOM_Test_yyyyMMdd_HHmmss.csv
# ============================================================

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

$script:OutputFolder = "D:\fanvil"
$script:LastResult = $null

if (!(Test-Path $script:OutputFolder)) {
    New-Item -ItemType Directory -Path $script:OutputFolder -Force | Out-Null
}

function Add-Log {
    param([string]$Text)
    $time = Get-Date -Format "HH:mm:ss"
    $txtLog.AppendText("[$time] $Text`r`n")
    $txtLog.SelectionStart = $txtLog.Text.Length
    $txtLog.ScrollToCaret()
    [System.Windows.Forms.Application]::DoEvents()
}

function Clean-Text {
    param([string]$Text)
    if (!$Text) { return "N/A" }
    $t = $Text -replace "`r"," " -replace "`n"," " -replace "`t"," "
    $t = [regex]::Replace($t, "\s+", " ").Trim()
    if (!$t) { return "N/A" }
    return $t
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

function Extract-AfterLabel {
    param([string]$Text, [string[]]$Labels)
    foreach ($label in $Labels) {
        $pattern = [regex]::Escape($label) + "\s*:?\s*([A-Za-z0-9\.\-_:\/@]+(?:\s+[A-Za-z0-9\.\-_:\/@]+)?)"
        $m = [regex]::Match($Text, $pattern, "IgnoreCase")
        if ($m.Success) {
            $v = Clean-Text $m.Groups[1].Value
            if ($v -and $v -ne "N/A") { return $v }
        }
    }
    return "N/A"
}

function Get-BrowserText {
    try {
        if ($browser.Document -and $browser.Document.Body) {
            return [string]$browser.Document.Body.InnerText
        }
    } catch {}
    return ""
}

function Get-BrowserHtml {
    try {
        if ($browser.Document -and $browser.Document.Body) {
            return [string]$browser.Document.Body.OuterHtml
        }
    } catch {}
    return ""
}

function Auto-Login {
    param([string]$User, [string]$Pass)

    try {
        $doc = $browser.Document
        if (!$doc) {
            Add-Log "No browser document found."
            return
        }

        $inputs = $doc.GetElementsByTagName("input")
        $userSet = $false
        $passSet = $false

        foreach ($inp in $inputs) {
            $name = ""
            $id = ""
            $type = ""
            try { $name = [string]$inp.GetAttribute("name") } catch {}
            try { $id = [string]$inp.GetAttribute("id") } catch {}
            try { $type = [string]$inp.GetAttribute("type") } catch {}

            $key = ($name + " " + $id + " " + $type).ToLower()

            if (!$userSet -and ($key -match "user|name|account|login|admin")) {
                try { $inp.SetAttribute("value", $User); $userSet = $true } catch {}
            }

            if (!$passSet -and ($key -match "pass|pwd")) {
                try { $inp.SetAttribute("value", $Pass); $passSet = $true } catch {}
            }
        }

        Add-Log "Username filled: $userSet | Password filled: $passSet"

        $clicked = $false

        foreach ($inp in $inputs) {
            $type = ""
            $value = ""
            try { $type = [string]$inp.GetAttribute("type") } catch {}
            try { $value = [string]$inp.GetAttribute("value") } catch {}

            if (($type.ToLower() -match "submit|button") -or ($value.ToLower() -match "login|log in|submit")) {
                try {
                    $inp.InvokeMember("click")
                    $clicked = $true
                    break
                } catch {}
            }
        }

        if (!$clicked) {
            $buttons = $doc.GetElementsByTagName("button")
            foreach ($btn in $buttons) {
                $txt = ""
                try { $txt = [string]$btn.InnerText } catch {}
                if ($txt.ToLower() -match "login|log in|submit") {
                    try {
                        $btn.InvokeMember("click")
                        $clicked = $true
                        break
                    } catch {}
                }
            }
        }

        Add-Log "Login button clicked: $clicked"

        if (!$clicked) {
            Add-Log "Auto login could not find login button. Please login manually in the browser area."
        }
    }
    catch {
        Add-Log "Auto login error: $($_.Exception.Message)"
    }
}

function Fetch-Data {
    $ip = $txtIP.Text.Trim()

    $visibleText = Get-BrowserText
    $html = Get-BrowserHtml
    $combined = Clean-Text ($visibleText + " " + $html)

    if (!$combined -or $combined -eq "N/A") {
        Add-Log "No visible text found. Open Information page first."
        return
    }

    $stamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $debugTxt = Join-Path $script:OutputFolder "Fanvil_DOM_Text_$stamp.txt"
    $debugHtml = Join-Path $script:OutputFolder "Fanvil_DOM_Html_$stamp.html"

    $visibleText | Out-File $debugTxt -Encoding UTF8
    $html | Out-File $debugHtml -Encoding UTF8

    $model = "N/A"
    if ($combined -match "X301G") { $model = "X301G" }
    elseif ($combined -match "X3SG") { $model = "X3SG" }
    else { $model = Extract-AfterLabel $combined @("Model", "Phone Model", "Device Model") }

    $hardware = Extract-AfterLabel $combined @("Hardware")
    $software = Extract-AfterLabel $combined @("Software", "Firmware", "Firmware Version", "Version")
    $networkMode = Extract-AfterLabel $combined @("Network mode", "Network Mode", "IP Mode")

    $ethernetMac = Extract-Regex $combined @(
        "Ethernet\s*MAC\s*:?\s*([0-9A-Fa-f]{2}[:-][0-9A-Fa-f]{2}[:-][0-9A-Fa-f]{2}[:-][0-9A-Fa-f]{2}[:-][0-9A-Fa-f]{2}[:-][0-9A-Fa-f]{2})",
        "MAC\s*Address\s*:?\s*([0-9A-Fa-f]{2}[:-][0-9A-Fa-f]{2}[:-][0-9A-Fa-f]{2}[:-][0-9A-Fa-f]{2}[:-][0-9A-Fa-f]{2}[:-][0-9A-Fa-f]{2})",
        "MAC\s*:?\s*([0-9A-Fa-f]{2}[:-][0-9A-Fa-f]{2}[:-][0-9A-Fa-f]{2}[:-][0-9A-Fa-f]{2}[:-][0-9A-Fa-f]{2}[:-][0-9A-Fa-f]{2})"
    )

    $ethernetIP = Extract-Regex $combined @(
        "Ethernet\s*IP\s*:?\s*((?:\d{1,3}\.){3}\d{1,3})",
        "IP\s*Address\s*:?\s*((?:\d{1,3}\.){3}\d{1,3})",
        "IPv4\s*:?\s*((?:\d{1,3}\.){3}\d{1,3})"
    )

    if ($ethernetIP -eq "N/A") { $ethernetIP = $ip }

    $subnet = Extract-Regex $combined @(
        "Subnet\s*mask\s*:?\s*((?:\d{1,3}\.){3}\d{1,3})",
        "Subnet\s*Mask\s*:?\s*((?:\d{1,3}\.){3}\d{1,3})"
    )

    $gateway = Extract-Regex $combined @(
        "Default\s*gateway\s*:?\s*((?:\d{1,3}\.){3}\d{1,3})",
        "Default\s*Gateway\s*:?\s*((?:\d{1,3}\.){3}\d{1,3})",
        "Gateway\s*:?\s*((?:\d{1,3}\.){3}\d{1,3})"
    )

    $extension = Extract-Regex $combined @(
        "Extension\s*:?\s*([0-9]{2,10})",
        "SIP\s*User\s*ID\s*:?\s*([0-9]{2,10})",
        "User\s*ID\s*:?\s*([0-9]{2,10})",
        "Register\s*Name\s*:?\s*([0-9]{2,10})",
        "Phone\s*Number\s*:?\s*([0-9]{2,10})",
        "Account\s*1\s*:?\s*([0-9]{2,10})",
        "Auth\s*User\s*:?\s*([0-9]{2,10})"
    )

    $script:LastResult = [PSCustomObject]@{
        TestDate       = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        PhoneIP        = $ip
        Model          = $model
        Extension      = $extension
        EthernetMAC    = $ethernetMac
        EthernetIP     = $ethernetIP
        NetworkMode    = $networkMode
        SubnetMask     = $subnet
        DefaultGateway = $gateway
        Hardware       = $hardware
        Software       = $software
        DebugTextFile  = $debugTxt
        DebugHtmlFile  = $debugHtml
    }

    $txtResult.Text = @"
Phone IP        : $ip
Model           : $model
Extension       : $extension
Ethernet MAC    : $ethernetMac
Ethernet IP     : $ethernetIP
Network Mode    : $networkMode
Subnet Mask     : $subnet
Default Gateway : $gateway
Hardware        : $hardware
Software        : $software

Debug Text      : $debugTxt
Debug HTML      : $debugHtml
"@

    Add-Log "Fetch completed. MAC: $ethernetMac | Extension: $extension"

    $csv = Join-Path $script:OutputFolder "Fanvil_Browser_DOM_Test_$stamp.csv"
    $script:LastResult | Export-Csv -Path $csv -NoTypeInformation -Encoding UTF8
    Add-Log "CSV saved: $csv"
}

# ---------------- GUI ----------------

$form = New-Object System.Windows.Forms.Form
$form.Text = "Fanvil Browser DOM Test"
$form.Size = New-Object System.Drawing.Size(1100, 760)
$form.StartPosition = "CenterScreen"
$form.Font = New-Object System.Drawing.Font("Segoe UI", 9)

$lblTitle = New-Object System.Windows.Forms.Label
$lblTitle.Text = "Fanvil Browser DOM Test"
$lblTitle.Font = New-Object System.Drawing.Font("Segoe UI", 16, [System.Drawing.FontStyle]::Bold)
$lblTitle.AutoSize = $true
$lblTitle.Location = New-Object System.Drawing.Point(15, 12)
$form.Controls.Add($lblTitle)

$lblIP = New-Object System.Windows.Forms.Label
$lblIP.Text = "Phone IP:"
$lblIP.AutoSize = $true
$lblIP.Location = New-Object System.Drawing.Point(20, 55)
$form.Controls.Add($lblIP)

$txtIP = New-Object System.Windows.Forms.TextBox
$txtIP.Location = New-Object System.Drawing.Point(90, 51)
$txtIP.Size = New-Object System.Drawing.Size(150, 25)
$form.Controls.Add($txtIP)

$lblUser = New-Object System.Windows.Forms.Label
$lblUser.Text = "User:"
$lblUser.AutoSize = $true
$lblUser.Location = New-Object System.Drawing.Point(255, 55)
$form.Controls.Add($lblUser)

$txtUser = New-Object System.Windows.Forms.TextBox
$txtUser.Location = New-Object System.Drawing.Point(300, 51)
$txtUser.Size = New-Object System.Drawing.Size(90, 25)
$txtUser.Text = "admin"
$form.Controls.Add($txtUser)

$lblPass = New-Object System.Windows.Forms.Label
$lblPass.Text = "Password:"
$lblPass.AutoSize = $true
$lblPass.Location = New-Object System.Drawing.Point(405, 55)
$form.Controls.Add($lblPass)

$txtPass = New-Object System.Windows.Forms.TextBox
$txtPass.Location = New-Object System.Drawing.Point(480, 51)
$txtPass.Size = New-Object System.Drawing.Size(90, 25)
$txtPass.Text = "admin"
$form.Controls.Add($txtPass)

$btnOpen = New-Object System.Windows.Forms.Button
$btnOpen.Text = "Open Phone"
$btnOpen.Location = New-Object System.Drawing.Point(590, 48)
$btnOpen.Size = New-Object System.Drawing.Size(100, 30)
$form.Controls.Add($btnOpen)

$btnLogin = New-Object System.Windows.Forms.Button
$btnLogin.Text = "Auto Login"
$btnLogin.Location = New-Object System.Drawing.Point(700, 48)
$btnLogin.Size = New-Object System.Drawing.Size(100, 30)
$form.Controls.Add($btnLogin)

$btnInfo = New-Object System.Windows.Forms.Button
$btnInfo.Text = "Info Page"
$btnInfo.Location = New-Object System.Drawing.Point(810, 48)
$btnInfo.Size = New-Object System.Drawing.Size(90, 30)
$form.Controls.Add($btnInfo)

$btnFetch = New-Object System.Windows.Forms.Button
$btnFetch.Text = "Fetch Visible Data"
$btnFetch.Location = New-Object System.Drawing.Point(910, 48)
$btnFetch.Size = New-Object System.Drawing.Size(140, 30)
$form.Controls.Add($btnFetch)

$browser = New-Object System.Windows.Forms.WebBrowser
$browser.Location = New-Object System.Drawing.Point(20, 90)
$browser.Size = New-Object System.Drawing.Size(650, 430)
$browser.ScriptErrorsSuppressed = $true
$form.Controls.Add($browser)

$txtResult = New-Object System.Windows.Forms.TextBox
$txtResult.Location = New-Object System.Drawing.Point(690, 90)
$txtResult.Size = New-Object System.Drawing.Size(370, 430)
$txtResult.Multiline = $true
$txtResult.ScrollBars = "Vertical"
$txtResult.ReadOnly = $true
$txtResult.Font = New-Object System.Drawing.Font("Consolas", 10)
$form.Controls.Add($txtResult)

$txtLog = New-Object System.Windows.Forms.TextBox
$txtLog.Location = New-Object System.Drawing.Point(20, 540)
$txtLog.Size = New-Object System.Drawing.Size(1040, 160)
$txtLog.Multiline = $true
$txtLog.ScrollBars = "Vertical"
$txtLog.ReadOnly = $true
$form.Controls.Add($txtLog)

$btnOpen.Add_Click({
    $ip = $txtIP.Text.Trim()
    if (!$ip -or $ip -notmatch "^(\d{1,3}\.){3}\d{1,3}$") {
        [System.Windows.Forms.MessageBox]::Show("Enter valid phone IP.", "Invalid IP", "OK", "Warning") | Out-Null
        return
    }
    Add-Log "Opening http://$ip"
    $browser.Navigate("http://$ip")
})

$btnLogin.Add_Click({
    Add-Log "Trying auto login."
    Auto-Login -User $txtUser.Text.Trim() -Pass $txtPass.Text.Trim()
})

$btnInfo.Add_Click({
    $ip = $txtIP.Text.Trim()
    if (!$ip) { return }
    Add-Log "Opening common information page."
    $browser.Navigate("http://$ip/information.htm")
})

$btnFetch.Add_Click({
    Fetch-Data
})

Add-Log "Ready."
Add-Log "Open phone, login, open Information page, then click Fetch Visible Data."

[void]$form.ShowDialog()
