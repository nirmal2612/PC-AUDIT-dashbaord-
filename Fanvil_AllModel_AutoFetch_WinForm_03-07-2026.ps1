# Fanvil All Model Auto Fetch WinForms Tool
# Date: 03-07-2026
# Purpose:
#   1) Enter IP segment like 10.209.110 or start/end IP range
#   2) Scan active web devices
#   3) Detect Fanvil landline/IP phones
#   4) Fetch IP + MAC using ARP
#   5) Try normal web login using entered credential
#   6) Try multiple Fanvil web/config pages and parse model, firmware, extension/SIP fields where available
#   7) Export CSV and Excel-compatible file to D:\fanvil
# Requirement: Windows PowerShell 5.1
# Note: This tool does not bypass login. It only reads pages using the username/password provided in the app.

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Web

[System.Net.ServicePointManager]::Expect100Continue = $false
[System.Net.ServicePointManager]::SecurityProtocol = [System.Net.SecurityProtocolType]::Tls12
$ErrorActionPreference = 'Continue'

$script:OutDir = 'D:\fanvil'
$script:StopScan = $false
$script:Results = New-Object System.Collections.ArrayList

if (!(Test-Path $script:OutDir)) {
    New-Item -Path $script:OutDir -ItemType Directory -Force | Out-Null
}

function Write-UiLog {
    param([string]$Message)
    try {
        $time = Get-Date -Format 'HH:mm:ss'
        $script:txtLog.AppendText("[$time] $Message`r`n")
        $script:txtLog.SelectionStart = $script:txtLog.Text.Length
        $script:txtLog.ScrollToCaret()
        [System.Windows.Forms.Application]::DoEvents()
    } catch {}
}

function Set-UiStatus {
    param([string]$Message)
    try {
        $script:lblStatus.Text = $Message
        [System.Windows.Forms.Application]::DoEvents()
    } catch {}
}

function ConvertTo-PlainText {
    param([string]$Html)
    if ([string]::IsNullOrWhiteSpace($Html)) { return '' }
    try {
        $x = $Html -replace '<script[\s\S]*?</script>', ' '
        $x = $x -replace '<style[\s\S]*?</style>', ' '
        $x = $x -replace '<[^>]+>', ' '
        $x = [System.Web.HttpUtility]::HtmlDecode($x)
        $x = $x -replace '\s+', ' '
        return $x.Trim()
    } catch { return $Html }
}

function Test-PortOpen {
    param([string]$Ip, [int]$Port, [int]$TimeoutMs = 700)
    try {
        $client = New-Object System.Net.Sockets.TcpClient
        $iar = $client.BeginConnect($Ip, $Port, $null, $null)
        $ok = $iar.AsyncWaitHandle.WaitOne($TimeoutMs, $false)
        if ($ok) { $client.EndConnect($iar); $client.Close(); return $true }
        $client.Close(); return $false
    } catch { return $false }
}

function Get-ArpMac {
    param([string]$Ip)
    try {
        ping $Ip -n 1 -w 500 | Out-Null
        Start-Sleep -Milliseconds 80
        $arp = arp -a $Ip 2>$null | Out-String
        foreach ($line in ($arp -split "`r?`n")) {
            if ($line -match [regex]::Escape($Ip) -and $line -match '([0-9A-Fa-f]{2}[-:]){5}[0-9A-Fa-f]{2}') {
                return $Matches[0].ToUpper().Replace('-', ':')
            }
        }
        return ''
    } catch { return '' }
}

function Invoke-WebSafe {
    param(
        [string]$Url,
        [Microsoft.PowerShell.Commands.WebRequestSession]$Session = $null,
        [string]$Method = 'GET',
        [object]$Body = $null,
        [hashtable]$Headers = $null,
        [int]$TimeoutSec = 5
    )
    try {
        $params = @{
            Uri = $Url
            Method = $Method
            TimeoutSec = $TimeoutSec
            UseBasicParsing = $true
            MaximumRedirection = 5
            ErrorAction = 'Stop'
        }
        if ($Session) { $params.WebSession = $Session }
        if ($Body) { $params.Body = $Body }
        if ($Headers) { $params.Headers = $Headers }
        return Invoke-WebRequest @params
    } catch {
        return $null
    }
}

function Get-BaseUrls {
    param([string]$Ip)
    $urls = New-Object System.Collections.ArrayList
    if (Test-PortOpen -Ip $Ip -Port 80) { [void]$urls.Add("http://$Ip") }
    if (Test-PortOpen -Ip $Ip -Port 8080) { [void]$urls.Add("http://$Ip`:8080") }
    if (Test-PortOpen -Ip $Ip -Port 443) { [void]$urls.Add("https://$Ip") }
    return @($urls)
}

function Detect-FanvilWeb {
    param([string]$Ip)
    $bases = Get-BaseUrls -Ip $Ip
    foreach ($base in $bases) {
        $resp = Invoke-WebSafe -Url $base -TimeoutSec 4
        $html = if ($resp) { [string]$resp.Content } else { '' }
        $server = if ($resp -and $resp.Headers['Server']) { [string]$resp.Headers['Server'] } else { '' }
        $plain = ConvertTo-PlainText $html
        $isFanvil = $false
        if ($html -match '(?i)fanvil|X301|X3S|X3SG|X4|X5|X6|V62|V64|V65|V67|H2U|H3|H5|SIP Phone|VoIP Phone') { $isFanvil = $true }
        if ($server -match '(?i)fanvil|GoAhead|Boa|lighttpd|mini_httpd') { $isFanvil = $true }
        if ($plain -match '(?i)Fanvil|SIP Phone|VoIP Phone|Phone Information|Account|Line') { $isFanvil = $true }
        if ($isFanvil) {
            return [pscustomobject]@{ IsFanvil=$true; BaseUrl=$base; Server=$server; HomeHtml=$html }
        }
    }
    if ($bases.Count -gt 0) {
        return [pscustomobject]@{ IsFanvil=$false; BaseUrl=$bases[0]; Server=''; HomeHtml='' }
    }
    return [pscustomobject]@{ IsFanvil=$false; BaseUrl=''; Server=''; HomeHtml='' }
}

function Try-FanvilLogin {
    param(
        [string]$BaseUrl,
        [string]$Username,
        [string]$Password
    )
    $session = New-Object Microsoft.PowerShell.Commands.WebRequestSession
    [void](Invoke-WebSafe -Url $BaseUrl -Session $session -TimeoutSec 4)

    $loginAttempts = @(
        @{ Url='/cgi-bin/ConfigManApp.com'; Body=@{ username=$Username; password=$Password; submit='Login' } },
        @{ Url='/cgi-bin/ConfigManApp.com'; Body=@{ user=$Username; pwd=$Password; Login='Login' } },
        @{ Url='/login.cgi'; Body=@{ username=$Username; password=$Password } },
        @{ Url='/cgi-bin/login.cgi'; Body=@{ username=$Username; password=$Password } },
        @{ Url='/goform/login'; Body=@{ username=$Username; password=$Password } },
        @{ Url='/login'; Body=@{ username=$Username; password=$Password } }
    )

    foreach ($a in $loginAttempts) {
        $url = $BaseUrl + $a.Url
        $resp = Invoke-WebSafe -Url $url -Session $session -Method 'POST' -Body $a.Body -TimeoutSec 5
        if ($resp) {
            $txt = [string]$resp.Content
            if ($txt -match '(?i)logout|phone information|basic|account|line|sip|system|network|config|status' -and $txt -notmatch '(?i)login failed|invalid password|unauthorized') {
                return [pscustomobject]@{ Success=$true; Session=$session; Method=$a.Url; Message='Login Success' }
            }
        }
        $test = Invoke-WebSafe -Url ($BaseUrl + '/cgi-bin/ConfigManApp.com') -Session $session -TimeoutSec 4
        if ($test -and ([string]$test.Content) -match '(?i)logout|phone information|account|line|sip|system|status') {
            return [pscustomobject]@{ Success=$true; Session=$session; Method=$a.Url; Message='Login Success' }
        }
    }

    # Basic authentication fallback
    try {
        $pair = "$Username`:$Password"
        $bytes = [System.Text.Encoding]::ASCII.GetBytes($pair)
        $basic = [Convert]::ToBase64String($bytes)
        $headers = @{ Authorization = "Basic $basic" }
        $resp = Invoke-WebSafe -Url ($BaseUrl + '/cgi-bin/ConfigManApp.com') -Headers $headers -TimeoutSec 5
        if ($resp -and ([string]$resp.Content) -match '(?i)logout|phone information|account|line|sip|system|status') {
            return [pscustomobject]@{ Success=$true; Session=$session; Method='BasicAuth'; Message='Login Success' }
        }
    } catch {}

    return [pscustomobject]@{ Success=$false; Session=$session; Method=''; Message='Login Failed / Page protected' }
}

function Add-FoundValue {
    param([hashtable]$Map, [string]$Key, [string]$Value)
    if ([string]::IsNullOrWhiteSpace($Value)) { return }
    $clean = ($Value -replace '[\r\n\t]+',' ' -replace '\s+',' ').Trim(' ',':','=','"','''')
    if ([string]::IsNullOrWhiteSpace($clean)) { return }
    if (!$Map.ContainsKey($Key) -or [string]::IsNullOrWhiteSpace($Map[$Key])) { $Map[$Key] = $clean }
}

function Parse-FanvilDetails {
    param([string[]]$Pages)
    $m = @{}
    foreach ($html in $Pages) {
        if ([string]::IsNullOrWhiteSpace($html)) { continue }
        $text = ConvertTo-PlainText $html
        $all = $html + ' ' + $text

        $patterns = @(
            @{K='Model'; R='(?i)(?:Model|Product\s*Model|Device\s*Model|Phone\s*Model)[\s:=\"''>]+([A-Za-z0-9\-_\. ]{2,40})'},
            @{K='Firmware'; R='(?i)(?:Firmware|Software\s*Version|Version|SW\s*Version)[\s:=\"''>]+([A-Za-z0-9\-_\. ]{2,50})'},
            @{K='Hardware'; R='(?i)(?:Hardware\s*Version|HW\s*Version)[\s:=\"''>]+([A-Za-z0-9\-_\. ]{2,50})'},
            @{K='DeviceName'; R='(?i)(?:Device\s*Name|Phone\s*Name|Host\s*Name|Hostname)[\s:=\"''>]+([A-Za-z0-9\-_\. ]{2,50})'},
            @{K='SerialNumber'; R='(?i)(?:Serial\s*Number|SN|S/N)[\s:=\"''>]+([A-Za-z0-9\-_\.]{4,60})'},
            @{K='WebMac'; R='(?i)(?:MAC|MAC\s*Address)[\s:=\"''>]+(([0-9A-Fa-f]{2}[:-]){5}[0-9A-Fa-f]{2})'},
            @{K='Extension'; R='(?i)(?:Extension|Ext\.?|Line\s*Number|Phone\s*Number|Number)[\s:=\"''>]+([0-9A-Za-z_\-\.]{2,40})'},
            @{K='SipUser'; R='(?i)(?:SIP\s*User|SIP\s*Username|User\s*Name|Username|Auth\s*User|Register\s*Name|Account)[\s:=\"''>]+([0-9A-Za-z_\-\.@]{2,80})'},
            @{K='DisplayName'; R='(?i)(?:Display\s*Name|Label|Name)[\s:=\"''>]+([0-9A-Za-z_\-\. ]{2,80})'},
            @{K='SipServer'; R='(?i)(?:SIP\s*Server|Registrar|Proxy\s*Server|Server\s*Address)[\s:=\"''>]+([0-9A-Za-z_\-\.@:\/]{2,100})'},
            @{K='LineStatus'; R='(?i)(?:Line\s*Status|Register\s*Status|Registration\s*Status|Status)[\s:=\"''>]+([A-Za-z0-9_\-\. ]{2,50})'}
        )
        foreach ($p in $patterns) {
            $rx = [regex]::Matches($all, $p.R)
            foreach ($match in $rx) {
                if ($match.Groups.Count -gt 1) { Add-FoundValue -Map $m -Key $p.K -Value $match.Groups[1].Value; break }
            }
        }

        # Parse HTML input values where Fanvil pages store data in form fields.
        $inputMatches = [regex]::Matches($html, '(?is)<input[^>]+>')
        foreach ($im in $inputMatches) {
            $tag = $im.Value
            $name = ''
            $value = ''
            if ($tag -match '(?i)\b(?:name|id)=["'']?([^"''\s>]+)') { $name = $Matches[1] }
            if ($tag -match '(?i)\bvalue=["'']([^"'']*)') { $value = $Matches[1] }
            if ([string]::IsNullOrWhiteSpace($name) -or [string]::IsNullOrWhiteSpace($value)) { continue }
            if ($name -match '(?i)model') { Add-FoundValue $m 'Model' $value }
            elseif ($name -match '(?i)firmware|version|sw') { Add-FoundValue $m 'Firmware' $value }
            elseif ($name -match '(?i)mac') { Add-FoundValue $m 'WebMac' $value }
            elseif ($name -match '(?i)display|label') { Add-FoundValue $m 'DisplayName' $value }
            elseif ($name -match '(?i)user|account|auth|extension|number') { Add-FoundValue $m 'SipUser' $value; if ($value -match '^\d{2,10}$') { Add-FoundValue $m 'Extension' $value } }
            elseif ($name -match '(?i)server|proxy|registrar') { Add-FoundValue $m 'SipServer' $value }
        }
    }
    return $m
}

function Fetch-FanvilPages {
    param(
        [string]$BaseUrl,
        [Microsoft.PowerShell.Commands.WebRequestSession]$Session
    )
    $paths = @(
        '/',
        '/cgi-bin/ConfigManApp.com',
        '/cgi-bin/ConfigManApp.com?Id=0',
        '/cgi-bin/ConfigManApp.com?Id=1',
        '/cgi-bin/ConfigManApp.com?Id=2',
        '/cgi-bin/ConfigManApp.com?Id=3',
        '/cgi-bin/ConfigManApp.com?Id=4',
        '/cgi-bin/ConfigManApp.com?Id=5',
        '/cgi-bin/ConfigManApp.com?Id=6',
        '/cgi-bin/ConfigManApp.com?Id=7',
        '/cgi-bin/ConfigManApp.com?Id=8',
        '/cgi-bin/ConfigManApp.com?Id=9',
        '/cgi-bin/ConfigManApp.com?Id=10',
        '/cgi-bin/ConfigManApp.com?Id=11',
        '/cgi-bin/ConfigManApp.com?Id=12',
        '/cgi-bin/ConfigManApp.com?Id=13',
        '/cgi-bin/ConfigManApp.com?Id=14',
        '/cgi-bin/ConfigManApp.com?Id=15',
        '/cgi-bin/ConfigManApp.com?Id=16',
        '/cgi-bin/ConfigManApp.com?Id=17',
        '/cgi-bin/ConfigManApp.com?Id=18',
        '/cgi-bin/ConfigManApp.com?Id=19',
        '/cgi-bin/ConfigManApp.com?Id=20',
        '/phone_info.htm',
        '/phoneInfo.htm',
        '/status.htm',
        '/Status.htm',
        '/device.htm',
        '/network.htm',
        '/line.htm',
        '/account.htm',
        '/Account.htm',
        '/sip.htm',
        '/config.htm',
        '/cgi-bin/status.cgi',
        '/cgi-bin/phoneinfo.cgi',
        '/cgi-bin/Account.cgi'
    )
    $pages = New-Object System.Collections.ArrayList
    foreach ($p in $paths) {
        if ($script:StopScan) { break }
        $resp = Invoke-WebSafe -Url ($BaseUrl + $p) -Session $Session -TimeoutSec 4
        if ($resp -and $resp.Content) {
            $content = [string]$resp.Content
            if ($content.Length -gt 30) { [void]$pages.Add($content) }
        }
    }
    return @($pages)
}

function New-IpListFromInput {
    param([string]$SegmentOrStart, [string]$EndIp)
    $ips = New-Object System.Collections.ArrayList
    $s = $SegmentOrStart.Trim()
    $e = $EndIp.Trim()

    if ($s -match '^\d{1,3}\.\d{1,3}\.\d{1,3}$') {
        for ($i=1; $i -le 254; $i++) { [void]$ips.Add("$s.$i") }
    } elseif ($s -match '^\d{1,3}\.\d{1,3}\.\d{1,3}\.$') {
        $prefix = $s.TrimEnd('.')
        for ($i=1; $i -le 254; $i++) { [void]$ips.Add("$prefix.$i") }
    } elseif ($s -match '^\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}$' -and $e -match '^\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}$') {
        $sp = $s.Split('.')
        $ep = $e.Split('.')
        if (($sp[0] -eq $ep[0]) -and ($sp[1] -eq $ep[1]) -and ($sp[2] -eq $ep[2])) {
            $prefix = "$($sp[0]).$($sp[1]).$($sp[2])"
            $startLast = [int]$sp[3]
            $endLast = [int]$ep[3]
            if ($startLast -gt $endLast) { $tmp=$startLast; $startLast=$endLast; $endLast=$tmp }
            for ($i=$startLast; $i -le $endLast; $i++) { [void]$ips.Add("$prefix.$i") }
        }
    } elseif ($s -match '^\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}$') {
        [void]$ips.Add($s)
    }
    return @($ips)
}

function Export-Results {
    param([object[]]$Data)
    if (!$Data -or $Data.Count -eq 0) { return '' }
    $stamp = Get-Date -Format 'yyyyMMdd_HHmmss'
    $csv = Join-Path $script:OutDir "Fanvil_AutoFetch_$stamp.csv"
    $xlsx = Join-Path $script:OutDir "Fanvil_AutoFetch_$stamp.xlsx"
    $Data | Export-Csv -Path $csv -NoTypeInformation -Encoding UTF8

    # Create XLSX if Excel is installed. CSV is always created.
    try {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $wb = $excel.Workbooks.Open($csv)
        $ws = $wb.Worksheets.Item(1)
        $ws.Columns.AutoFit() | Out-Null
        $wb.SaveAs($xlsx, 51)
        $wb.Close($true)
        $excel.Quit()
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($ws) | Out-Null
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($wb) | Out-Null
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
        return $xlsx
    } catch {
        return $csv
    }
}

function Add-GridRow {
    param([pscustomobject]$Obj)
    try {
        [void]$script:grid.Rows.Add(
            $Obj.SNo, $Obj.IPAddress, $Obj.MACAddress, $Obj.Brand, $Obj.Model,
            $Obj.Firmware, $Obj.Extension, $Obj.SIPUsername, $Obj.DisplayName,
            $Obj.LineStatus, $Obj.LoginStatus, $Obj.Remarks
        )
        [System.Windows.Forms.Application]::DoEvents()
    } catch {}
}

function Start-FanvilScan {
    $script:StopScan = $false
    $script:Results.Clear() | Out-Null
    $script:grid.Rows.Clear()

    $segment = $script:txtSegment.Text.Trim()
    $endIp = $script:txtEndIp.Text.Trim()
    $user = $script:txtUser.Text.Trim()
    $pass = $script:txtPass.Text

    if ([string]::IsNullOrWhiteSpace($segment)) {
        [System.Windows.Forms.MessageBox]::Show('Enter IP segment like 10.209.110 or start IP like 10.209.110.1', 'Input Required') | Out-Null
        return
    }

    $ips = New-IpListFromInput -SegmentOrStart $segment -EndIp $endIp
    if (!$ips -or $ips.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show('Invalid input. Use segment 10.209.110 or range 10.209.110.1 to 10.209.110.254', 'Invalid IP') | Out-Null
        return
    }

    $script:btnScan.Enabled = $false
    $script:btnStop.Enabled = $true
    Write-UiLog "Scan started. Total IPs: $($ips.Count)"

    $sno = 0
    foreach ($ip in $ips) {
        if ($script:StopScan) { break }
        Set-UiStatus "Scanning $ip"
        Write-UiLog "Checking $ip"

        $baseInfo = Detect-FanvilWeb -Ip $ip
        if ([string]::IsNullOrWhiteSpace($baseInfo.BaseUrl)) { continue }

        $mac = Get-ArpMac -Ip $ip
        $brand = if ($baseInfo.IsFanvil) { 'Fanvil' } else { 'Web Device' }
        $model = ''
        $firmware = ''
        $extension = ''
        $sipUser = ''
        $displayName = ''
        $lineStatus = ''
        $loginStatus = 'Not Tried'
        $remarks = 'Web found'
        $deviceName = ''
        $serial = ''
        $hardware = ''
        $sipServer = ''

        if ($baseInfo.IsFanvil) {
            $remarks = 'Fanvil detected'
            Write-UiLog "Fanvil detected: $ip"
            if (![string]::IsNullOrWhiteSpace($user) -and ![string]::IsNullOrWhiteSpace($pass)) {
                $login = Try-FanvilLogin -BaseUrl $baseInfo.BaseUrl -Username $user -Password $pass
                $loginStatus = $login.Message
                if ($login.Success) {
                    Write-UiLog "Login success: $ip using $($login.Method)"
                    $pages = Fetch-FanvilPages -BaseUrl $baseInfo.BaseUrl -Session $login.Session
                    $pages = @($baseInfo.HomeHtml) + @($pages)
                    $details = Parse-FanvilDetails -Pages $pages
                    if ($details.ContainsKey('Model')) { $model = $details['Model'] }
                    if ($details.ContainsKey('Firmware')) { $firmware = $details['Firmware'] }
                    if ($details.ContainsKey('Extension')) { $extension = $details['Extension'] }
                    if ($details.ContainsKey('SipUser')) { $sipUser = $details['SipUser'] }
                    if ($details.ContainsKey('DisplayName')) { $displayName = $details['DisplayName'] }
                    if ($details.ContainsKey('LineStatus')) { $lineStatus = $details['LineStatus'] }
                    if ($details.ContainsKey('DeviceName')) { $deviceName = $details['DeviceName'] }
                    if ($details.ContainsKey('SerialNumber')) { $serial = $details['SerialNumber'] }
                    if ($details.ContainsKey('Hardware')) { $hardware = $details['Hardware'] }
                    if ($details.ContainsKey('SipServer')) { $sipServer = $details['SipServer'] }
                    if ([string]::IsNullOrWhiteSpace($mac) -and $details.ContainsKey('WebMac')) { $mac = $details['WebMac'].ToUpper().Replace('-', ':') }
                    if ([string]::IsNullOrWhiteSpace($model) -and $baseInfo.HomeHtml -match '(?i)(X301G|X3SG|X3S|X4|X5U|X6U|V62|V64|V65|V67|H2U|H3|H5)') { $model = $Matches[1] }
                    if ([string]::IsNullOrWhiteSpace($extension) -and $sipUser -match '^\d{2,10}$') { $extension = $sipUser }
                    $remarks = if ([string]::IsNullOrWhiteSpace($extension) -and [string]::IsNullOrWhiteSpace($sipUser)) { 'Login success - page format maybe not matched' } else { 'Details fetched' }
                } else {
                    Write-UiLog "Login failed/protected: $ip"
                    $remarks = 'Fanvil detected - login failed or protected'
                }
            } else {
                $loginStatus = 'Credential not entered'
                $remarks = 'Fanvil detected - enter username/password to fetch line details'
            }
        } else {
            $loginStatus = 'Not Fanvil'
            $remarks = 'Web device found, Fanvil keyword not confirmed'
        }

        $sno++
        $obj = [pscustomobject]@{
            SNo = $sno
            IPAddress = $ip
            MACAddress = $mac
            Brand = $brand
            Model = $model
            Firmware = $firmware
            Extension = $extension
            SIPUsername = $sipUser
            DisplayName = $displayName
            PhoneNumber = $extension
            LineStatus = $lineStatus
            DeviceName = $deviceName
            SerialNumber = $serial
            Hardware = $hardware
            SIPServer = $sipServer
            WebUrl = $baseInfo.BaseUrl
            WebServer = $baseInfo.Server
            LoginStatus = $loginStatus
            Remarks = $remarks
            ScanTime = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
        }
        [void]$script:Results.Add($obj)
        Add-GridRow -Obj $obj
    }

    $out = Export-Results -Data @($script:Results)
    if ($out) {
        Write-UiLog "Export completed: $out"
        Set-UiStatus "Completed. File: $out"
        [System.Windows.Forms.MessageBox]::Show("Scan completed.`r`nFile saved:`r`n$out", 'Completed') | Out-Null
    } else {
        Write-UiLog 'No device result to export.'
        Set-UiStatus 'Completed. No result.'
        [System.Windows.Forms.MessageBox]::Show('Scan completed. No web/Fanvil device found.', 'Completed') | Out-Null
    }
    $script:btnScan.Enabled = $true
    $script:btnStop.Enabled = $false
}

# ---------------------- UI ----------------------
$form = New-Object System.Windows.Forms.Form
$form.Text = 'Fanvil Auto Fetch Tool - All Model Scanner'
$form.Size = New-Object System.Drawing.Size(1160, 720)
$form.StartPosition = 'CenterScreen'
$form.Font = New-Object System.Drawing.Font('Segoe UI', 9)
$form.BackColor = [System.Drawing.Color]::FromArgb(245,247,250)

$title = New-Object System.Windows.Forms.Label
$title.Text = 'Fanvil Auto Fetch Tool'
$title.Font = New-Object System.Drawing.Font('Segoe UI', 18, [System.Drawing.FontStyle]::Bold)
$title.AutoSize = $true
$title.Location = New-Object System.Drawing.Point(20, 15)
$form.Controls.Add($title)

$sub = New-Object System.Windows.Forms.Label
$sub.Text = 'Scan IP segment, detect Fanvil phones, login with provided credential, fetch IP / MAC / model / extension / SIP details and export to D:\fanvil'
$sub.AutoSize = $true
$sub.Location = New-Object System.Drawing.Point(23, 55)
$sub.ForeColor = [System.Drawing.Color]::FromArgb(75,85,99)
$form.Controls.Add($sub)

$panel = New-Object System.Windows.Forms.Panel
$panel.Location = New-Object System.Drawing.Point(20, 85)
$panel.Size = New-Object System.Drawing.Size(1100, 95)
$panel.BackColor = [System.Drawing.Color]::White
$panel.BorderStyle = 'FixedSingle'
$form.Controls.Add($panel)

function Add-Label($text, $x, $y) {
    $l = New-Object System.Windows.Forms.Label
    $l.Text = $text
    $l.AutoSize = $true
    $l.Location = New-Object System.Drawing.Point($x, $y)
    $panel.Controls.Add($l)
    return $l
}

Add-Label 'IP Segment / Start IP' 15 15 | Out-Null
$txtSegment = New-Object System.Windows.Forms.TextBox
$txtSegment.Location = New-Object System.Drawing.Point(15, 38)
$txtSegment.Size = New-Object System.Drawing.Size(160, 25)
$txtSegment.Text = '10.209.110'
$panel.Controls.Add($txtSegment)
$script:txtSegment = $txtSegment

Add-Label 'End IP optional' 190 15 | Out-Null
$txtEndIp = New-Object System.Windows.Forms.TextBox
$txtEndIp.Location = New-Object System.Drawing.Point(190, 38)
$txtEndIp.Size = New-Object System.Drawing.Size(160, 25)
$txtEndIp.Text = ''
$panel.Controls.Add($txtEndIp)
$script:txtEndIp = $txtEndIp

Add-Label 'Username' 365 15 | Out-Null
$txtUser = New-Object System.Windows.Forms.TextBox
$txtUser.Location = New-Object System.Drawing.Point(365, 38)
$txtUser.Size = New-Object System.Drawing.Size(130, 25)
$txtUser.Text = 'admin'
$panel.Controls.Add($txtUser)
$script:txtUser = $txtUser

Add-Label 'Password' 510 15 | Out-Null
$txtPass = New-Object System.Windows.Forms.TextBox
$txtPass.Location = New-Object System.Drawing.Point(510, 38)
$txtPass.Size = New-Object System.Drawing.Size(130, 25)
$txtPass.UseSystemPasswordChar = $true
$txtPass.Text = 'admin'
$panel.Controls.Add($txtPass)
$script:txtPass = $txtPass

$btnScan = New-Object System.Windows.Forms.Button
$btnScan.Text = 'SCAN && EXPORT'
$btnScan.Location = New-Object System.Drawing.Point(665, 30)
$btnScan.Size = New-Object System.Drawing.Size(140, 38)
$btnScan.BackColor = [System.Drawing.Color]::FromArgb(37,99,235)
$btnScan.ForeColor = [System.Drawing.Color]::White
$btnScan.FlatStyle = 'Flat'
$panel.Controls.Add($btnScan)
$script:btnScan = $btnScan

$btnStop = New-Object System.Windows.Forms.Button
$btnStop.Text = 'STOP'
$btnStop.Location = New-Object System.Drawing.Point(815, 30)
$btnStop.Size = New-Object System.Drawing.Size(90, 38)
$btnStop.Enabled = $false
$btnStop.BackColor = [System.Drawing.Color]::FromArgb(220,38,38)
$btnStop.ForeColor = [System.Drawing.Color]::White
$btnStop.FlatStyle = 'Flat'
$panel.Controls.Add($btnStop)
$script:btnStop = $btnStop

$btnOpen = New-Object System.Windows.Forms.Button
$btnOpen.Text = 'OPEN OUTPUT'
$btnOpen.Location = New-Object System.Drawing.Point(915, 30)
$btnOpen.Size = New-Object System.Drawing.Size(120, 38)
$btnOpen.FlatStyle = 'Flat'
$panel.Controls.Add($btnOpen)

$grid = New-Object System.Windows.Forms.DataGridView
$grid.Location = New-Object System.Drawing.Point(20, 195)
$grid.Size = New-Object System.Drawing.Size(1100, 330)
$grid.AllowUserToAddRows = $false
$grid.AllowUserToDeleteRows = $false
$grid.ReadOnly = $true
$grid.AutoSizeColumnsMode = 'Fill'
$grid.BackgroundColor = [System.Drawing.Color]::White
$grid.RowHeadersVisible = $false
$grid.SelectionMode = 'FullRowSelect'
$form.Controls.Add($grid)
$script:grid = $grid

$cols = @('SNo','IP Address','MAC Address','Brand','Model','Firmware','Extension','SIP Username','Display Name','Line Status','Login Status','Remarks')
foreach ($c in $cols) {
    $col = New-Object System.Windows.Forms.DataGridViewTextBoxColumn
    $col.HeaderText = $c
    $col.Name = $c.Replace(' ','')
    [void]$grid.Columns.Add($col)
}
$grid.Columns[0].Width = 45

$txtLog = New-Object System.Windows.Forms.TextBox
$txtLog.Location = New-Object System.Drawing.Point(20, 540)
$txtLog.Size = New-Object System.Drawing.Size(1100, 100)
$txtLog.Multiline = $true
$txtLog.ScrollBars = 'Vertical'
$txtLog.ReadOnly = $true
$txtLog.BackColor = [System.Drawing.Color]::White
$form.Controls.Add($txtLog)
$script:txtLog = $txtLog

$lblStatus = New-Object System.Windows.Forms.Label
$lblStatus.Text = 'Ready'
$lblStatus.Location = New-Object System.Drawing.Point(20, 650)
$lblStatus.Size = New-Object System.Drawing.Size(1100, 25)
$lblStatus.ForeColor = [System.Drawing.Color]::FromArgb(55,65,81)
$form.Controls.Add($lblStatus)
$script:lblStatus = $lblStatus

$btnScan.Add_Click({ Start-FanvilScan })
$btnStop.Add_Click({ $script:StopScan = $true; Write-UiLog 'Stop requested...' })
$btnOpen.Add_Click({ if (!(Test-Path $script:OutDir)) { New-Item -Path $script:OutDir -ItemType Directory -Force | Out-Null }; Start-Process explorer.exe $script:OutDir })

Write-UiLog 'Ready. Enter segment like 10.209.110 or range start/end IP, then click SCAN & EXPORT.'
[void]$form.ShowDialog()
