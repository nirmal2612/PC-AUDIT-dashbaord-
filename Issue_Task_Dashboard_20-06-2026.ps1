#requires -version 5.1
<#
================================================================================
 ISSUE / TASK WEB DASHBOARD - PowerShell Hosted
 Date    : 20-06-2026
 Host    : http://10.209.110.220:8080/
 Purpose : Section / Category / Project / Station wise Issue & Task Dashboard
 Author  : Generated for internal dashboard use
================================================================================

MAIN FEATURES
- PowerShell HttpListener web app, no IIS required
- Domain login support for idpbgtn.com when enabled for user
- Local admin fallback login
- Role/access based menu visibility
- Users can search/find issues by section/category/project/station/status/error code/DRI/date
- Add Issue menu visible only when admin gives AddIssue access
- Admin can manage users, sections, categories, projects, station types
- Multiple visible sections per user
- Image upload support, max 10 images per issue
- Issue fields: section, category, project, station, title, error code, description,
  need to check, DRI, issue happen date, root cause, temporary/permanent solution,
  status, images

FIRST LOGIN
- User ID : admin
- Password: admin@123
- Change password by editing D:\IssueTaskDashboard\Data\Users.json after first run,
  or create another admin from Admin > Users.

RUN AS ADMIN RECOMMENDED
- Start PowerShell as Administrator
- Run: powershell -ExecutionPolicy Bypass -File .\Issue_Task_Dashboard_20-06-2026.ps1
================================================================================
#>

# ========================= CONFIG =========================
$Script:Config = [ordered]@{
    AppName       = 'ISSUE / TASK DASHBOARD'
    IP            = '10.209.110.220'
    Port          = 8080
    Domain        = 'idpbgtn.com'
    RootPath      = 'D:\IssueTaskDashboard'
    DataPath      = 'D:\IssueTaskDashboard\Data'
    UploadPath    = 'D:\IssueTaskDashboard\Uploads'
    BackupPath    = 'D:\IssueTaskDashboard\Backup'
    LogPath       = 'D:\IssueTaskDashboard\Logs'
    MaxImages     = 10
    MaxUploadMB   = 8
    SessionHours  = 10
}

$Script:UsersFile      = Join-Path $Script:Config.DataPath 'Users.json'
$Script:IssuesFile     = Join-Path $Script:Config.DataPath 'Issues.json'
$Script:MastersFile    = Join-Path $Script:Config.DataPath 'Masters.json'
$Script:Sessions       = @{}

# ========================= BASIC UTILITIES =========================
function Ensure-Folders {
    foreach ($p in @($Config.RootPath,$Config.DataPath,$Config.UploadPath,$Config.BackupPath,$Config.LogPath)) {
        if (-not (Test-Path $p)) { New-Item -Path $p -ItemType Directory -Force | Out-Null }
    }
}

function Write-AppLog([string]$Message, [string]$Level='INFO') {
    try {
        $line = "$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss') [$Level] $Message"
        Add-Content -Path (Join-Path $Config.LogPath "dashboard_$(Get-Date -Format 'yyyyMMdd').log") -Value $line -Encoding UTF8
    } catch {}
}

function ConvertTo-JsonSafe($Object) { $Object | ConvertTo-Json -Depth 20 }

function Load-JsonFile($Path, $Default) {
    try {
        if (Test-Path $Path) {
            $raw = Get-Content $Path -Raw -Encoding UTF8
            if ($raw.Trim()) { return ($raw | ConvertFrom-Json) }
        }
    } catch { Write-AppLog "Load JSON failed: $Path - $($_.Exception.Message)" 'ERROR' }
    return $Default
}

function Save-JsonFile($Path, $Object) {
    try {
        $tmp = "$Path.tmp"
        $Object | ConvertTo-Json -Depth 30 | Set-Content -Path $tmp -Encoding UTF8
        Move-Item -Path $tmp -Destination $Path -Force
        return $true
    } catch {
        Write-AppLog "Save JSON failed: $Path - $($_.Exception.Message)" 'ERROR'
        return $false
    }
}

function New-Id([string]$Prefix) { "$Prefix-$([guid]::NewGuid().ToString('N').Substring(0,12).ToUpper())" }

function HtmlEncode([string]$s) {
    if ($null -eq $s) { return '' }
    return [System.Net.WebUtility]::HtmlEncode($s)
}

function UrlDecode([string]$s) { [System.Net.WebUtility]::UrlDecode($s) }

function Get-BodyText($Request) {
    $reader = New-Object IO.StreamReader($Request.InputStream, $Request.ContentEncoding)
    $body = $reader.ReadToEnd()
    $reader.Close()
    return $body
}

function Parse-FormUrlEncoded([string]$Body) {
    $dict = @{}
    if ([string]::IsNullOrWhiteSpace($Body)) { return $dict }
    foreach ($pair in $Body -split '&') {
        if ($pair -match '=') {
            $k,$v = $pair -split '=',2
            $key = UrlDecode($k)
            $val = UrlDecode($v.Replace('+',' '))
            if ($dict.ContainsKey($key)) {
                if ($dict[$key] -is [array]) { $dict[$key] += $val } else { $dict[$key] = @($dict[$key],$val) }
            } else { $dict[$key] = $val }
        }
    }
    return $dict
}

function Send-Response($Context, [string]$Content, [string]$ContentType='text/html; charset=utf-8', [int]$Status=200) {
    try {
        $bytes = [Text.Encoding]::UTF8.GetBytes($Content)
        $Context.Response.StatusCode = $Status
        $Context.Response.ContentType = $ContentType
        $Context.Response.ContentLength64 = $bytes.Length
        $Context.Response.OutputStream.Write($bytes,0,$bytes.Length)
    } finally { $Context.Response.OutputStream.Close() }
}

function Send-Redirect($Context, [string]$Url) {
    $Context.Response.StatusCode = 302
    $Context.Response.RedirectLocation = $Url
    $Context.Response.OutputStream.Close()
}

function Get-Cookie($Request, [string]$Name) {
    foreach ($c in $Request.Cookies) { if ($c.Name -eq $Name) { return $c.Value } }
    return $null
}

function Set-Cookie($Response, [string]$Name, [string]$Value) {
    $cookie = New-Object System.Net.Cookie($Name,$Value,'/')
    $cookie.Expires = (Get-Date).AddHours($Config.SessionHours)
    $Response.Cookies.Add($cookie)
}

# ========================= INITIAL DATA =========================
function Initialize-Data {
    Ensure-Folders

    if (-not (Test-Path $Script:MastersFile)) {
        $masters = [ordered]@{
            Sections   = @('Network','OA','IMES')
            Categories = @('FATP','MLB','WH','TT','Other')
            Projects   = @('Roma','Jesko','Alpine','Elva')
            Stations   = @('A2','Q6','Ri')
            Statuses   = @('Open','Ongoing','Hold','Closed')
            SolutionTypes = @('Temporary','Permanent')
        }
        Save-JsonFile $Script:MastersFile $masters | Out-Null
    }

    if (-not (Test-Path $Script:UsersFile)) {
        $admin = [ordered]@{
            UserId = 'admin'
            DisplayName = 'System Admin'
            Role = 'Admin'
            Password = 'admin@123'
            DomainLogin = $false
            IsActive = $true
            VisibleSections = @('Network','OA','IMES')
            CanAddIssue = $true
            CanEditIssue = $true
            CanManageMasters = $true
            CanManageUsers = $true
            CreatedAt = (Get-Date).ToString('s')
        }
        Save-JsonFile $Script:UsersFile @($admin) | Out-Null
    }

    if (-not (Test-Path $Script:IssuesFile)) { Save-JsonFile $Script:IssuesFile @() | Out-Null }
}

function Get-Users { @(Load-JsonFile $Script:UsersFile @()) }
function Save-Users($u) { Save-JsonFile $Script:UsersFile @($u) | Out-Null }
function Get-Issues { @(Load-JsonFile $Script:IssuesFile @()) }
function Save-Issues($i) { Save-JsonFile $Script:IssuesFile @($i) | Out-Null }
function Get-Masters { Load-JsonFile $Script:MastersFile ([ordered]@{}) }
function Save-Masters($m) { Save-JsonFile $Script:MastersFile $m | Out-Null }

# ========================= AUTH =========================
function Test-DomainCredential([string]$UserId, [string]$Password) {
    try {
        Add-Type -AssemblyName System.DirectoryServices.AccountManagement -ErrorAction SilentlyContinue
        $ctx = New-Object System.DirectoryServices.AccountManagement.PrincipalContext('Domain', $Config.Domain)
        return $ctx.ValidateCredentials($UserId.ToLower(), $Password)
    } catch {
        Write-AppLog "Domain auth failed for $UserId : $($_.Exception.Message)" 'WARN'
        return $false
    }
}

function Authenticate-User([string]$UserId, [string]$Password) {
    $users = Get-Users
    $user = $users | Where-Object { $_.UserId -ieq $UserId -and $_.IsActive -eq $true } | Select-Object -First 1
    if (-not $user) { return $null }

    if ($user.DomainLogin -eq $true) {
        if (Test-DomainCredential $UserId $Password) { return $user }
    } else {
        if ($user.Password -eq $Password) { return $user }
    }
    return $null
}

function New-Session($Response, $User) {
    $token = [guid]::NewGuid().ToString('N')
    $Script:Sessions[$token] = [ordered]@{ UserId=$User.UserId; Expires=(Get-Date).AddHours($Config.SessionHours) }
    Set-Cookie $Response 'ITD_SESSION' $token
}

function Get-CurrentUser($Request) {
    $token = Get-Cookie $Request 'ITD_SESSION'
    if (-not $token -or -not $Script:Sessions.ContainsKey($token)) { return $null }
    $s = $Script:Sessions[$token]
    if ([datetime]$s.Expires -lt (Get-Date)) { $Script:Sessions.Remove($token); return $null }
    $users = Get-Users
    return ($users | Where-Object { $_.UserId -eq $s.UserId -and $_.IsActive -eq $true } | Select-Object -First 1)
}

function Require-Login($Context) {
    $u = Get-CurrentUser $Context.Request
    if (-not $u) { Send-Redirect $Context '/login'; return $null }
    return $u
}

function Is-Admin($User) { return ($User.Role -eq 'Admin') }
function User-Sections($User) { @($User.VisibleSections) }
function Has-SectionAccess($User, [string]$Section) { (Is-Admin $User) -or ((User-Sections $User) -contains $Section) }

# ========================= HTML LAYOUT =========================
function Get-Style {
@'
<style>
:root{--bg:#eef3f8;--card:#ffffff;--ink:#1d2939;--muted:#667085;--brand:#174ea6;--brand2:#0b7285;--ok:#108548;--bad:#c2410c;--line:#d8e0ea;--soft:#f7f9fc}
*{box-sizing:border-box} body{margin:0;font-family:Segoe UI,Arial,sans-serif;background:var(--bg);color:var(--ink)}
a{text-decoration:none;color:inherit}.top{background:linear-gradient(135deg,var(--brand),var(--brand2));color:#fff;padding:16px 24px;display:flex;align-items:center;justify-content:space-between;box-shadow:0 3px 12px #0002}.top h1{font-size:20px;margin:0;letter-spacing:.5px}.wrap{display:flex;min-height:calc(100vh - 64px)}.side{width:245px;background:#fff;border-right:1px solid var(--line);padding:14px;position:sticky;top:0;height:calc(100vh - 64px)}.side a{display:block;padding:11px 13px;margin:6px 0;border-radius:10px;color:#344054;font-weight:600}.side a:hover,.side .active{background:#e8f0fe;color:#174ea6}.main{flex:1;padding:22px}.card{background:var(--card);border:1px solid var(--line);border-radius:16px;padding:18px;margin-bottom:16px;box-shadow:0 2px 8px #0000000d}.grid{display:grid;grid-template-columns:repeat(4,minmax(160px,1fr));gap:14px}.metric{background:var(--soft);border:1px solid var(--line);border-radius:14px;padding:16px}.metric b{font-size:26px;color:var(--brand)}.metric span{display:block;color:var(--muted);font-weight:600;margin-top:6px}.formgrid{display:grid;grid-template-columns:repeat(2,minmax(240px,1fr));gap:12px}.full{grid-column:1/-1}label{font-weight:700;color:#344054;display:block;margin-bottom:6px}input,select,textarea{width:100%;padding:10px 12px;border:1px solid #cfd8e3;border-radius:10px;font:inherit;background:#fff}textarea{min-height:90px}.btn{display:inline-block;border:0;border-radius:10px;padding:10px 14px;font-weight:800;cursor:pointer;background:var(--brand);color:#fff;margin:4px}.btn.secondary{background:#475467}.btn.ok{background:var(--ok)}.btn.bad{background:var(--bad)}.btn.light{background:#e8eef7;color:#1d2939}table{width:100%;border-collapse:collapse;background:#fff}th,td{padding:10px;border-bottom:1px solid var(--line);text-align:left;vertical-align:top}th{background:#f2f6fb;font-size:13px;color:#344054}.tag{display:inline-block;border-radius:999px;padding:4px 9px;background:#edf2ff;color:#174ea6;font-size:12px;font-weight:800}.status-Open{background:#fff1f2;color:#be123c}.status-Ongoing{background:#e0f2fe;color:#0369a1}.status-Hold{background:#fef3c7;color:#92400e}.status-Closed{background:#dcfce7;color:#166534}.login{max-width:430px;margin:8vh auto}.login .card{padding:28px}.hint{color:var(--muted);font-size:13px}.alert{padding:10px 12px;border-radius:10px;margin-bottom:12px;background:#fff7ed;color:#9a3412;border:1px solid #fed7aa}.imgs img{height:75px;border-radius:10px;border:1px solid #ddd;margin:3px}.small{font-size:12px;color:var(--muted)}@media(max-width:850px){.wrap{display:block}.side{width:auto;height:auto;position:relative}.grid,.formgrid{grid-template-columns:1fr}.main{padding:12px}}
</style>
'@
}

function Layout($User, [string]$Title, [string]$Body, [string]$Active='') {
    $canAdd = ($User.CanAddIssue -eq $true -or (Is-Admin $User))
    $canMaster = ($User.CanManageMasters -eq $true -or (Is-Admin $User))
    $canUsers = ($User.CanManageUsers -eq $true -or (Is-Admin $User))
    $menu = @()
    $menu += "<a class='$(if($Active -eq 'dashboard'){'active'})' href='/'>Dashboard</a>"
    $menu += "<a class='$(if($Active -eq 'issues'){'active'})' href='/issues'>Search / Find Issues</a>"
    if ($canAdd) { $menu += "<a class='$(if($Active -eq 'add'){'active'})' href='/issue/add'>Add Issue</a>" }
    if ($canMaster) { $menu += "<a class='$(if($Active -eq 'masters'){'active'})' href='/admin/masters'>Admin Masters</a>" }
    if ($canUsers) { $menu += "<a class='$(if($Active -eq 'users'){'active'})' href='/admin/users'>Admin Users</a>" }
    $menu += "<a href='/logout'>Logout</a>"
@"
<!doctype html><html><head><meta charset='utf-8'><meta name='viewport' content='width=device-width,initial-scale=1'>
<title>$($Config.AppName)</title>$(Get-Style)</head><body>
<div class='top'><h1>$($Config.AppName)</h1><div>$(HtmlEncode $User.DisplayName) <span class='tag'>$(HtmlEncode $User.Role)</span></div></div>
<div class='wrap'><div class='side'>$($menu -join "`n")</div><main class='main'><h2>$Title</h2>$Body</main></div>
</body></html>
"@
}

function Login-Page([string]$Error='') {
    $err = if($Error){"<div class='alert'>$(HtmlEncode $Error)</div>"}else{''}
@"
<!doctype html><html><head><meta charset='utf-8'><meta name='viewport' content='width=device-width,initial-scale=1'>
<title>Login - $($Config.AppName)</title>$(Get-Style)</head><body>
<div class='login'><div class='card'>
<h2 style='margin-top:0'>$($Config.AppName)</h2>
<p class='hint'>Login with local password or domain password when domain login is enabled for your user.</p>
$err
<form method='post' action='/login'>
<label>User ID</label><input name='userid' required autofocus placeholder='Employee ID / admin'>
<label>Password</label><input name='password' type='password' required placeholder='Password'>
<button class='btn' style='width:100%;margin-top:14px'>Login</button>
</form>
<p class='small'>Default first login: admin / admin@123</p>
</div></div></body></html>
"@
}

# ========================= PAGES =========================
function Dashboard-Page($Context) {
    $u = Require-Login $Context; if(!$u){return}
    $issues = Get-Issues | Where-Object { Has-SectionAccess $u $_.Section }
    $open = @($issues | Where-Object Status -eq 'Open').Count
    $ongo = @($issues | Where-Object Status -eq 'Ongoing').Count
    $hold = @($issues | Where-Object Status -eq 'Hold').Count
    $closed = @($issues | Where-Object Status -eq 'Closed').Count
    $sections = User-Sections $u
    if (Is-Admin $u) { $sections = @((Get-Masters).Sections) }
    $sectionRows = foreach($s in $sections){
        $cnt = @($issues | Where-Object Section -eq $s).Count
        "<tr><td><a class='tag' href='/issues?section=$( [uri]::EscapeDataString($s) )'>$s</a></td><td>$cnt</td><td>$(@($issues|?{ $_.Section -eq $s -and $_.Status -eq 'Open'}).Count)</td><td>$(@($issues|?{ $_.Section -eq $s -and $_.Status -eq 'Closed'}).Count)</td></tr>"
    }
    $body = @"
<div class='grid'>
 <div class='metric'><b>$(@($issues).Count)</b><span>Total Issues</span></div>
 <div class='metric'><b>$open</b><span>Open</span></div>
 <div class='metric'><b>$ongo</b><span>Ongoing</span></div>
 <div class='metric'><b>$closed</b><span>Closed</span></div>
</div>
<div class='card'><h3>Section Wise Summary</h3><table><tr><th>Section</th><th>Total</th><th>Open</th><th>Closed</th></tr>$($sectionRows -join "`n")</table></div>
<div class='card'><a class='btn' href='/issues'>Search / Find Existing Issue</a>$(if($u.CanAddIssue -or (Is-Admin $u)){"<a class='btn ok' href='/issue/add'>Add New Issue</a>"})</div>
"@
    Send-Response $Context (Layout $u 'Dashboard Summary' $body 'dashboard')
}

function Issues-Page($Context) {
    $u = Require-Login $Context; if(!$u){return}
    $m = Get-Masters
    $q = Parse-FormUrlEncoded ($Context.Request.Url.Query.TrimStart('?'))
    $issues = Get-Issues | Where-Object { Has-SectionAccess $u $_.Section }
    foreach($key in @('section','category','project','station','status')){ if($q[$key]){ $issues = $issues | Where-Object { $_.$key.Substring(0,1).ToUpper() + $_.$key.Substring(1) } } }
    if($q.section){ $issues = $issues | Where-Object { $_.Section -eq $q.section } }
    if($q.category){ $issues = $issues | Where-Object { $_.Category -eq $q.category } }
    if($q.project){ $issues = $issues | Where-Object { $_.Project -eq $q.project } }
    if($q.station){ $issues = $issues | Where-Object { $_.Station -eq $q.station } }
    if($q.status){ $issues = $issues | Where-Object { $_.Status -eq $q.status } }
    if($q.search){ $s=$q.search; $issues = $issues | Where-Object { $_.IssueTitle -like "*$s*" -or $_.ErrorCode -like "*$s*" -or $_.Description -like "*$s*" -or $_.DRI -like "*$s*" } }
    $issues = @($issues | Sort-Object CreatedAt -Descending)
    $sectionOptions = (@('') + @(User-Sections $u | Sort-Object)) | ForEach-Object { "<option value='$(HtmlEncode $_)' $(if($q.section -eq $_){'selected'})>$(if($_){HtmlEncode $_}else{'All Sections'})</option>" }
    if(Is-Admin $u){ $sectionOptions = (@('') + @($m.Sections)) | % { "<option value='$(HtmlEncode $_)' $(if($q.section -eq $_){'selected'})>$(if($_){HtmlEncode $_}else{'All Sections'})</option>" } }
    $rows = foreach($x in $issues){
        "<tr><td><a class='tag' href='/issue/view?id=$($x.Id)'>$($x.Id)</a></td><td>$(HtmlEncode $x.Section)</td><td>$(HtmlEncode $x.Category)</td><td>$(HtmlEncode $x.Project)</td><td>$(HtmlEncode $x.Station)</td><td>$(HtmlEncode $x.IssueTitle)<br><span class='small'>Error: $(HtmlEncode $x.ErrorCode)</span></td><td><span class='tag status-$($x.Status)'>$(HtmlEncode $x.Status)</span></td><td>$(HtmlEncode $x.DRI)</td><td>$(HtmlEncode $x.IssueDate)</td></tr>"
    }
    if(-not $rows){ $rows = "<tr><td colspan='9'>No issues found.</td></tr>" }
    $body = @"
<div class='card'><form method='get' action='/issues' class='formgrid'>
 <div><label>Section</label><select name='section'>$($sectionOptions -join '')</select></div>
 <div><label>Category</label><select name='category'><option value=''>All</option>$(($m.Categories|%{"<option $(if($q.category -eq $_){'selected'})>$(HtmlEncode $_)</option>"}) -join '')</select></div>
 <div><label>Project</label><select name='project'><option value=''>All</option>$(($m.Projects|%{"<option $(if($q.project -eq $_){'selected'})>$(HtmlEncode $_)</option>"}) -join '')</select></div>
 <div><label>Station Type</label><select name='station'><option value=''>All</option>$(($m.Stations|%{"<option $(if($q.station -eq $_){'selected'})>$(HtmlEncode $_)</option>"}) -join '')</select></div>
 <div><label>Status</label><select name='status'><option value=''>All</option>$(($m.Statuses|%{"<option $(if($q.status -eq $_){'selected'})>$(HtmlEncode $_)</option>"}) -join '')</select></div>
 <div><label>Search</label><input name='search' value='$(HtmlEncode $q.search)' placeholder='Title / error code / DRI / description'></div>
 <div class='full'><button class='btn'>Find Issue</button><a class='btn light' href='/issues'>Clear</a></div>
</form></div>
<div class='card'><h3>Issue List</h3><table><tr><th>ID</th><th>Section</th><th>Category</th><th>Project</th><th>Station</th><th>Issue</th><th>Status</th><th>DRI</th><th>Date</th></tr>$($rows -join "`n")</table></div>
"@
    Send-Response $Context (Layout $u 'Search / Find Issues' $body 'issues')
}

function Issue-Form-Page($Context, [string]$Error='') {
    $u = Require-Login $Context; if(!$u){return}
    if(-not ($u.CanAddIssue -eq $true -or (Is-Admin $u))){ Send-Response $Context (Layout $u 'Access Denied' '<div class="card">You do not have Add Issue access.</div>') ; return }
    $m = Get-Masters
    $sections = if(Is-Admin $u){@($m.Sections)}else{@($u.VisibleSections)}
    $err = if($Error){"<div class='alert'>$(HtmlEncode $Error)</div>"}else{''}
    $body = @"
$err
<div class='card'>
<form method='post' action='/issue/add' enctype='multipart/form-data'>
<div class='formgrid'>
 <div><label>Section</label><select name='section' required>$(($sections|%{"<option>$(HtmlEncode $_)</option>"}) -join '')</select></div>
 <div><label>Category</label><select name='category' required>$(($m.Categories|%{"<option>$(HtmlEncode $_)</option>"}) -join '')</select></div>
 <div><label>Project</label><select name='project'><option value=''>Not Applicable</option>$(($m.Projects|%{"<option>$(HtmlEncode $_)</option>"}) -join '')</select></div>
 <div><label>Station Type</label><select name='station'><option value=''>Not Applicable</option>$(($m.Stations|%{"<option>$(HtmlEncode $_)</option>"}) -join '')</select></div>
 <div class='full'><label>Requested Issue Title</label><input name='title' required maxlength='180'></div>
 <div><label>Error Code</label><input name='errorcode' placeholder='Optional'></div>
 <div><label>Respective DRI</label><input name='dri' required placeholder='Responsible person/team'></div>
 <div><label>Issue Happen Date</label><input name='issuedate' type='date' required value='$(Get-Date -Format yyyy-MM-dd)'></div>
 <div class='full'><label>Description</label><textarea name='description' required></textarea></div>
 <div class='full'><label>Need To Check</label><textarea name='needcheck'></textarea></div>
 <div class='full'><label>Root Cause</label><textarea name='rootcause'></textarea></div>
 <div><label>Solution Type</label><select name='solutiontype'><option>Temporary</option><option>Permanent</option></select></div>
 <div><label>Status</label><select name='status'><option>Open</option><option>Ongoing</option><option>Hold</option><option>Closed</option></select></div>
 <div class='full'><label>Temporary / Permanent Solution</label><textarea name='solution'></textarea></div>
 <div class='full'><label>Error Images - max $($Config.MaxImages)</label><input type='file' name='images' accept='image/*' multiple><p class='hint'>Upload up to $($Config.MaxImages) images. Each image max $($Config.MaxUploadMB) MB.</p></div>
</div>
<button class='btn ok' type='submit'>Submit Issue</button><a class='btn secondary' href='/issues'>Cancel</a>
</form></div>
<script>
document.querySelector('form').addEventListener('submit',function(e){var f=document.querySelector('input[type=file]'); if(f.files.length>$($Config.MaxImages)){alert('Maximum $($Config.MaxImages) images allowed'); e.preventDefault();}});
</script>
"@
    Send-Response $Context (Layout $u 'Add Issue' $body 'add')
}

# Simple multipart parser sufficient for normal browser form upload
function Parse-Multipart($Request) {
    $ct = $Request.ContentType
    if($ct -notmatch 'boundary=(.+)$'){ return @{Fields=@{};Files=@()} }
    $boundary = '--' + $Matches[1].Trim('"')
    $ms = New-Object IO.MemoryStream
    $Request.InputStream.CopyTo($ms)
    $bytes = $ms.ToArray()
    $text = [Text.Encoding]::ISO8859_1.GetString($bytes)
    $parts = $text -split [regex]::Escape($boundary)
    $fields=@{}; $files=@()
    foreach($part in $parts){
        if($part -notmatch 'Content-Disposition') { continue }
        $p = $part.Trim("`r","`n","-")
        $split = $p.IndexOf("`r`n`r`n")
        if($split -lt 0){ continue }
        $headers = $p.Substring(0,$split)
        $content = $p.Substring($split+4)
        if($headers -match 'name="([^"]+)"'){ $name=$Matches[1] } else { continue }
        if($headers -match 'filename="([^"]*)"'){
            $fn = [IO.Path]::GetFileName($Matches[1])
            if([string]::IsNullOrWhiteSpace($fn)){ continue }
            $contentType = if($headers -match 'Content-Type:\s*([^\r\n]+)'){$Matches[1].Trim()}else{'application/octet-stream'}
            $fileBytes = [Text.Encoding]::ISO8859_1.GetBytes($content)
            $files += [ordered]@{ Field=$name; FileName=$fn; ContentType=$contentType; Bytes=$fileBytes }
        } else {
            $fields[$name] = [Text.Encoding]::UTF8.GetString([Text.Encoding]::ISO8859_1.GetBytes($content)).TrimEnd("`r","`n")
        }
    }
    return @{Fields=$fields;Files=$files}
}

function Save-NewIssue($Context) {
    $u = Require-Login $Context; if(!$u){return}
    if(-not ($u.CanAddIssue -eq $true -or (Is-Admin $u))){ Send-Redirect $Context '/'; return }
    $mp = Parse-Multipart $Context.Request
    $f = $mp.Fields
    if(-not (Has-SectionAccess $u $f.section)){ Issue-Form-Page $Context 'Selected section is not allowed for your login.'; return }
    $id = New-Id 'ISSUE'
    $issueUploadDir = Join-Path $Config.UploadPath $id
    New-Item -Path $issueUploadDir -ItemType Directory -Force | Out-Null
    $savedImages=@()
    $imageFiles = @($mp.Files | Where-Object { $_.Field -eq 'images' })
    if($imageFiles.Count -gt $Config.MaxImages){ Issue-Form-Page $Context "Maximum $($Config.MaxImages) images only."; return }
    foreach($file in $imageFiles){
        if($file.Bytes.Length -gt ($Config.MaxUploadMB * 1MB)){ continue }
        if($file.ContentType -notlike 'image/*'){ continue }
        $ext = [IO.Path]::GetExtension($file.FileName); if(-not $ext){$ext='.jpg'}
        $safe = "IMG_$([guid]::NewGuid().ToString('N').Substring(0,8))$ext"
        [IO.File]::WriteAllBytes((Join-Path $issueUploadDir $safe), $file.Bytes)
        $savedImages += "/uploads/$id/$safe"
    }
    $issue = [ordered]@{
        Id=$id; Section=$f.section; Category=$f.category; Project=$f.project; Station=$f.station
        IssueTitle=$f.title; ErrorCode=$f.errorcode; Description=$f.description; NeedToCheck=$f.needcheck
        DRI=$f.dri; IssueDate=$f.issuedate; RootCause=$f.rootcause; SolutionType=$f.solutiontype
        Solution=$f.solution; Status=$f.status; Images=@($savedImages)
        CreatedBy=$u.UserId; CreatedByName=$u.DisplayName; CreatedAt=(Get-Date).ToString('s')
        UpdatedAt=(Get-Date).ToString('s')
    }
    $issues = @(Get-Issues); $issues += $issue; Save-Issues $issues
    Send-Redirect $Context "/issue/view?id=$id"
}

function Issue-View-Page($Context) {
    $u = Require-Login $Context; if(!$u){return}
    $q = Parse-FormUrlEncoded ($Context.Request.Url.Query.TrimStart('?'))
    $issue = Get-Issues | Where-Object Id -eq $q.id | Select-Object -First 1
    if(-not $issue -or -not (Has-SectionAccess $u $issue.Section)){ Send-Response $Context (Layout $u 'Issue Not Found' '<div class="card">Issue not found or no access.</div>'); return }
    $imgs = @($issue.Images) | ForEach-Object { "<a href='$_' target='_blank'><img src='$_'></a>" }
    $edit = if((Is-Admin $u) -or $u.CanEditIssue -eq $true){"<form method='post' action='/issue/status' class='card'><input type='hidden' name='id' value='$($issue.Id)'><label>Update Status</label><select name='status'><option>Open</option><option>Ongoing</option><option>Hold</option><option>Closed</option></select><label>Solution / Remark Update</label><textarea name='solution'>$(HtmlEncode $issue.Solution)</textarea><button class='btn ok'>Update</button></form>"}else{''}
    $body = @"
<div class='card'>
<h3>$(HtmlEncode $issue.IssueTitle) <span class='tag status-$($issue.Status)'>$(HtmlEncode $issue.Status)</span></h3>
<table>
<tr><th>Issue ID</th><td>$(HtmlEncode $issue.Id)</td><th>Section</th><td>$(HtmlEncode $issue.Section)</td></tr>
<tr><th>Category</th><td>$(HtmlEncode $issue.Category)</td><th>Project</th><td>$(HtmlEncode $issue.Project)</td></tr>
<tr><th>Station Type</th><td>$(HtmlEncode $issue.Station)</td><th>Error Code</th><td>$(HtmlEncode $issue.ErrorCode)</td></tr>
<tr><th>DRI</th><td>$(HtmlEncode $issue.DRI)</td><th>Issue Date</th><td>$(HtmlEncode $issue.IssueDate)</td></tr>
<tr><th>Created By</th><td>$(HtmlEncode $issue.CreatedByName)</td><th>Created At</th><td>$(HtmlEncode $issue.CreatedAt)</td></tr>
</table>
<h4>Description</h4><p>$(HtmlEncode $issue.Description)</p>
<h4>Need To Check</h4><p>$(HtmlEncode $issue.NeedToCheck)</p>
<h4>Root Cause</h4><p>$(HtmlEncode $issue.RootCause)</p>
<h4>$($issue.SolutionType) Solution</h4><p>$(HtmlEncode $issue.Solution)</p>
<h4>Images</h4><div class='imgs'>$($imgs -join '')</div>
</div>
$edit
<a class='btn secondary' href='/issues'>Back</a>
"@
    Send-Response $Context (Layout $u 'Issue Details' $body 'issues')
}

function Update-IssueStatus($Context) {
    $u = Require-Login $Context; if(!$u){return}
    if(-not ((Is-Admin $u) -or $u.CanEditIssue -eq $true)){ Send-Redirect $Context '/issues'; return }
    $f = Parse-FormUrlEncoded (Get-BodyText $Context.Request)
    $issues = @(Get-Issues)
    foreach($x in $issues){ if($x.Id -eq $f.id -and (Has-SectionAccess $u $x.Section)){ $x.Status=$f.status; $x.Solution=$f.solution; $x.UpdatedAt=(Get-Date).ToString('s') } }
    Save-Issues $issues
    Send-Redirect $Context "/issue/view?id=$($f.id)"
}

function Masters-Page($Context, [string]$Msg='') {
    $u = Require-Login $Context; if(!$u){return}
    if(-not ((Is-Admin $u) -or $u.CanManageMasters -eq $true)){ Send-Redirect $Context '/'; return }
    $m = Get-Masters
    $cards = foreach($type in @('Sections','Categories','Projects','Stations')){
        $items = @($m.$type) | % { "<span class='tag'>$(HtmlEncode $_)</span>" }
        "<div class='card'><h3>$type</h3><p>$($items -join ' ')</p><form method='post' action='/admin/masters/add'><input type='hidden' name='type' value='$type'><label>Add New $type</label><input name='value' required><button class='btn'>Add</button></form></div>"
    }
    $body = "$(if($Msg){"<div class='alert'>$Msg</div>"})$($cards -join '')"
    Send-Response $Context (Layout $u 'Admin Masters' $body 'masters')
}
function Add-Master($Context) {
    $u = Require-Login $Context; if(!$u){return}
    if(-not ((Is-Admin $u) -or $u.CanManageMasters -eq $true)){ Send-Redirect $Context '/'; return }
    $f=Parse-FormUrlEncoded (Get-BodyText $Context.Request); $m=Get-Masters
    $type=$f.type; $val=($f.value).Trim()
    if($val -and @('Sections','Categories','Projects','Stations') -contains $type){
        $arr=@($m.$type); if($arr -notcontains $val){ $arr += $val; $m.$type=$arr; Save-Masters $m }
    }
    Send-Redirect $Context '/admin/masters'
}

function Users-Page($Context) {
    $u=Require-Login $Context; if(!$u){return}
    if(-not ((Is-Admin $u) -or $u.CanManageUsers -eq $true)){ Send-Redirect $Context '/'; return }
    $m=Get-Masters; $users=Get-Users
    $secChecks = ($m.Sections | % { "<label style='display:inline-block;margin-right:12px'><input type='checkbox' name='sections' value='$(HtmlEncode $_)'> $(HtmlEncode $_)</label>" }) -join ''
    $rows = foreach($x in $users){ "<tr><td>$(HtmlEncode $x.UserId)</td><td>$(HtmlEncode $x.DisplayName)</td><td>$(HtmlEncode $x.Role)</td><td>$(@($x.VisibleSections) -join ', ')</td><td>$(HtmlEncode $x.DomainLogin)</td><td>$(HtmlEncode $x.CanAddIssue)</td><td>$(HtmlEncode $x.IsActive)</td></tr>" }
    $body=@"
<div class='card'><h3>Create / Update User</h3>
<form method='post' action='/admin/users/save' class='formgrid'>
<div><label>User ID</label><input name='userid' required placeholder='Employee ID'></div>
<div><label>Display Name</label><input name='displayname' required></div>
<div><label>Role</label><select name='role'><option>User</option><option>Lead</option><option>Manager</option><option>Admin</option></select></div>
<div><label>Local Password</label><input name='password' placeholder='Required if domain login off'></div>
<div class='full'><label>Visible Sections</label>$secChecks</div>
<div><label><input type='checkbox' name='domainlogin' value='true'> Use idpbgtn.com domain password</label></div>
<div><label><input type='checkbox' name='canadd' value='true'> Can Add Issue</label></div>
<div><label><input type='checkbox' name='canedit' value='true'> Can Edit Issue Status / Solution</label></div>
<div><label><input type='checkbox' name='canmasters' value='true'> Can Manage Masters</label></div>
<div><label><input type='checkbox' name='canusers' value='true'> Can Manage Users</label></div>
<div><label><input type='checkbox' name='active' value='true' checked> Active</label></div>
<div class='full'><button class='btn ok'>Save User</button></div>
</form></div>
<div class='card'><h3>User List</h3><table><tr><th>User ID</th><th>Name</th><th>Role</th><th>Sections</th><th>Domain</th><th>Add Issue</th><th>Active</th></tr>$($rows -join "`n")</table></div>
"@
    Send-Response $Context (Layout $u 'Admin Users' $body 'users')
}
function Save-User($Context){
    $u=Require-Login $Context; if(!$u){return}
    if(-not ((Is-Admin $u) -or $u.CanManageUsers -eq $true)){ Send-Redirect $Context '/'; return }
    $f=Parse-FormUrlEncoded (Get-BodyText $Context.Request); $users=@(Get-Users)
    $secs=@(); if($f.sections){ if($f.sections -is [array]){$secs=@($f.sections)}else{$secs=@($f.sections)} }
    $existing=$users|?{$_.UserId -ieq $f.userid}|select -first 1
    if($existing){
        $existing.DisplayName=$f.displayname; $existing.Role=$f.role; if($f.password){$existing.Password=$f.password}
        $existing.DomainLogin=($f.domainlogin -eq 'true'); $existing.IsActive=($f.active -eq 'true'); $existing.VisibleSections=@($secs)
        $existing.CanAddIssue=($f.canadd -eq 'true'); $existing.CanEditIssue=($f.canedit -eq 'true'); $existing.CanManageMasters=($f.canmasters -eq 'true'); $existing.CanManageUsers=($f.canusers -eq 'true')
    } else {
        $users += [ordered]@{UserId=$f.userid;DisplayName=$f.displayname;Role=$f.role;Password=$f.password;DomainLogin=($f.domainlogin -eq 'true');IsActive=($f.active -eq 'true');VisibleSections=@($secs);CanAddIssue=($f.canadd -eq 'true');CanEditIssue=($f.canedit -eq 'true');CanManageMasters=($f.canmasters -eq 'true');CanManageUsers=($f.canusers -eq 'true');CreatedAt=(Get-Date).ToString('s')}
    }
    Save-Users $users; Send-Redirect $Context '/admin/users'
}

function Serve-Upload($Context){
    $local = $Context.Request.Url.AbsolutePath.Replace('/uploads/','') -replace '/', '\'
    $path = Join-Path $Config.UploadPath $local
    if(Test-Path $path){
        $bytes=[IO.File]::ReadAllBytes($path); $Context.Response.ContentType='image/jpeg'; $Context.Response.ContentLength64=$bytes.Length; $Context.Response.OutputStream.Write($bytes,0,$bytes.Length); $Context.Response.OutputStream.Close()
    } else { Send-Response $Context 'Not found' 'text/plain' 404 }
}

# ========================= ROUTER =========================
function Handle-Request($Context) {
    try {
        $path = $Context.Request.Url.AbsolutePath
        $method = $Context.Request.HttpMethod
        Write-AppLog "$method $path"
        if($path -like '/uploads/*'){ Serve-Upload $Context; return }
        switch -Regex ("$method $path") {
            '^GET /login$' { Send-Response $Context (Login-Page); return }
            '^POST /login$' { $f=Parse-FormUrlEncoded (Get-BodyText $Context.Request); $u=Authenticate-User $f.userid $f.password; if($u){New-Session $Context.Response $u; Send-Redirect $Context '/'}else{Send-Response $Context (Login-Page 'Invalid user ID or password.')} ; return }
            '^GET /logout$' { $tok=Get-Cookie $Context.Request 'ITD_SESSION'; if($tok){$Script:Sessions.Remove($tok)}; Send-Redirect $Context '/login'; return }
            '^GET /$' { Dashboard-Page $Context; return }
            '^GET /issues$' { Issues-Page $Context; return }
            '^GET /issue/add$' { Issue-Form-Page $Context; return }
            '^POST /issue/add$' { Save-NewIssue $Context; return }
            '^GET /issue/view$' { Issue-View-Page $Context; return }
            '^POST /issue/status$' { Update-IssueStatus $Context; return }
            '^GET /admin/masters$' { Masters-Page $Context; return }
            '^POST /admin/masters/add$' { Add-Master $Context; return }
            '^GET /admin/users$' { Users-Page $Context; return }
            '^POST /admin/users/save$' { Save-User $Context; return }
            default { Send-Response $Context 'Page not found' 'text/plain' 404; return }
        }
    } catch {
        Write-AppLog "Unhandled error: $($_.Exception.Message)`n$($_.ScriptStackTrace)" 'ERROR'
        try { Send-Response $Context "Server error: $([System.Net.WebUtility]::HtmlEncode($_.Exception.Message))" 'text/html' 500 } catch {}
    }
}

# ========================= START SERVER =========================
Initialize-Data
$prefix = "http://$($Config.IP):$($Config.Port)/"
Write-Host "============================================================" -ForegroundColor Cyan
Write-Host " $($Config.AppName)" -ForegroundColor Cyan
Write-Host " URL : $prefix" -ForegroundColor Green
Write-Host " Data: $($Config.RootPath)" -ForegroundColor Green
Write-Host " Login: admin / admin@123" -ForegroundColor Yellow
Write-Host " Press Ctrl+C to stop" -ForegroundColor Yellow
Write-Host "============================================================" -ForegroundColor Cyan

try {
    $listener = New-Object System.Net.HttpListener
    $listener.Prefixes.Add($prefix)
    $listener.Start()
    Write-AppLog "Server started at $prefix"
    while($listener.IsListening){
        $ctx = $listener.GetContext()
        Handle-Request $ctx
    }
} catch {
    Write-AppLog "Server start failed: $($_.Exception.Message)" 'ERROR'
    Write-Host "ERROR: $($_.Exception.Message)" -ForegroundColor Red
    Write-Host "Try running PowerShell as Administrator or reserve URL:" -ForegroundColor Yellow
    Write-Host "netsh http add urlacl url=$prefix user=Everyone" -ForegroundColor Yellow
} finally {
    if($listener){ $listener.Stop(); $listener.Close() }
}
