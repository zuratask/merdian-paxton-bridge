# Merdian Paxton Bridge - CSV to Paxton NET2 (UI prototype: read-only SQL connect; no automated writes)
#
# NET2 SQL (same idea as MeridianSqlConfig / sqlcmd on the server box):
#   sqlcmd -S ".\NET2" -d Net2 -E -Q "SELECT ..."
#   => Server=.\NET2; Database=Net2; Integrated Security=True;  (this app uses SqlClient with that style)
#   -S / Server = instance (.\NET2 = local machine, NET2 instance). Use your real host if SQL is remote.
#   -d / Database = Paxton DB (often Net2). -E = Windows auth (same as SqlConnection Integrated Security=True).
#   User checks are SELECTs on dbo.Users / dbo.Cards etc., not the Net2 Access Control API.
#
# Run UI: powershell -ExecutionPolicy Bypass -File ".\N2SYNC-UI.ps1"
# Cursor terminal (connect + read CSV folder + print report, no GUI):
#   powershell -ExecutionPolicy Bypass -File ".\N2SYNC-UI.ps1" -CursorReport
# Optional N2SYNC.local.json next to script: SqlServer, SqlDatabase, WatchFolder, FilePattern (see N2SYNC.local.json.example)

param(
    [switch]$CursorReport,
    [string]$SqlServer = ".\NET2",
    [string]$SqlDatabase = "Net2",
    [string]$WatchFolder = "D:\Inbound\NET2Sync\",
    [string]$FilePattern = "*.csv"
)

$effSqlServer = $SqlServer
$effSqlDatabase = $SqlDatabase
$effWatchFolder = $WatchFolder
$effFilePattern = $FilePattern
$n2ScriptRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$cfgPath = Join-Path $n2ScriptRoot "N2SYNC.local.json"
if ($CursorReport -and (Test-Path -LiteralPath $cfgPath)) {
    try {
        $j = Get-Content -LiteralPath $cfgPath -Raw -Encoding UTF8 | ConvertFrom-Json
        if ($null -ne $j.SqlServer -and -not [string]::IsNullOrWhiteSpace([string]$j.SqlServer)) { $effSqlServer = [string]$j.SqlServer.Trim() }
        if ($null -ne $j.SqlDatabase -and -not [string]::IsNullOrWhiteSpace([string]$j.SqlDatabase)) { $effSqlDatabase = [string]$j.SqlDatabase.Trim() }
        if ($null -ne $j.WatchFolder -and -not [string]::IsNullOrWhiteSpace([string]$j.WatchFolder)) { $effWatchFolder = [string]$j.WatchFolder.Trim() }
        if ($null -ne $j.FilePattern -and -not [string]::IsNullOrWhiteSpace([string]$j.FilePattern)) { $effFilePattern = [string]$j.FilePattern.Trim() }
    } catch {
        Write-Warning ("N2SYNC.local.json ignored: " + $_.Exception.Message)
    }
}

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Add-Type -AssemblyName System.Data

[void][System.Windows.Forms.Application]::EnableVisualStyles()
[void][System.Windows.Forms.Application]::SetCompatibleTextRenderingDefault($false)

function New-N2Font {
    param(
        [string]$FamilyName = "Segoe UI",
        [float]$Size,
        [System.Drawing.FontStyle]$Style = [System.Drawing.FontStyle]::Regular
    )
    # 4-arg constructor avoids PowerShell New-Object failing to resolve (string,float,FontStyle) on some hosts.
    return New-Object System.Drawing.Font($FamilyName, $Size, $Style, [System.Drawing.GraphicsUnit]::Point)
}

$script:Ui = @{
    FormBg       = [System.Drawing.Color]::FromArgb(241, 245, 249)
    Surface      = [System.Drawing.Color]::White
    SurfaceAlt   = [System.Drawing.Color]::FromArgb(248, 250, 252)
    Border       = [System.Drawing.Color]::FromArgb(226, 232, 240)
    Text         = [System.Drawing.Color]::FromArgb(15, 23, 42)
    Muted        = [System.Drawing.Color]::FromArgb(100, 116, 139)
    Accent       = [System.Drawing.Color]::FromArgb(29, 78, 216)
    AccentHover  = [System.Drawing.Color]::FromArgb(37, 99, 235)
    AccentSoft   = [System.Drawing.Color]::FromArgb(239, 246, 255)
    Danger       = [System.Drawing.Color]::FromArgb(185, 28, 28)
    DangerSoft   = [System.Drawing.Color]::FromArgb(254, 242, 242)
    GridHeader   = [System.Drawing.Color]::FromArgb(30, 41, 59)
    GridSelect   = [System.Drawing.Color]::FromArgb(219, 234, 254)
    GridLine     = [System.Drawing.Color]::FromArgb(226, 232, 240)
}

$script:Connected = $false
$script:Net2Connection = $null
$script:Net2UserLookupCache = $null
# Bump when lookup rules change so an open session does not keep a stale column list.
$script:Net2SqlLookupSchemaRev = 8

function Close-Net2SqlConnection {
    if ($null -ne $script:Net2Connection) {
        try {
            if ($script:Net2Connection.State -eq [System.Data.ConnectionState]::Open) {
                $script:Net2Connection.Close()
            }
        } catch { }
        try { $script:Net2Connection.Dispose() } catch { }
        $script:Net2Connection = $null
    }
    $script:Connected = $false
    $script:Net2UserLookupCache = $null
}

function Escape-SqlConnFragment {
    param([string]$Value)
    if ($null -eq $Value) { return "" }
    # Semicolon or leading/trailing spaces break naive strings; strip CR/LF.
    return ($Value.Trim() -replace "[\r\n]+", "" -replace ";", "")
}

function Get-Net2ConnectionString {
    param(
        [string]$Server,
        [string]$Database,
        [string]$SqlUser,
        [string]$SqlPassword
    )
    $s = Escape-SqlConnFragment $Server
    $d = Escape-SqlConnFragment $Database
    if ([string]::IsNullOrWhiteSpace($s) -or $s -eq "(not configured)") {
        throw "Set NET2 SQL server on the Configuration tab (e.g. .\NET2 for local instance, or SERVER\INSTANCE)."
    }
    if ([string]::IsNullOrWhiteSpace($d)) {
        throw "Set NET2 database name (commonly Net2)."
    }
    # Build with Server=/Database= (not Data Source=) to avoid rare SqlClient "keyword not supported: datasource" issues.
    $parts = New-Object System.Collections.Generic.List[string]
    $parts.Add("Server=$s")
    $parts.Add("Database=$d")
    $parts.Add("Connect Timeout=15")
    $parts.Add("TrustServerCertificate=True")
    $uRaw = if ($null -eq $SqlUser) { "" } else { ($SqlUser.Trim() -replace "[\r\n]+", "") }
    if ([string]::IsNullOrWhiteSpace($uRaw)) {
        $parts.Add("Integrated Security=True")
    } else {
        $eu = $uRaw -replace ";", ""
        $pw = if ($null -eq $SqlPassword) { "" } else { $SqlPassword }
        $ep = (($pw -replace "[\r\n]+", "") -replace "'", "''")
        $parts.Add("User ID=$eu")
        $parts.Add("Password=$ep")
    }
    return ($parts -join ";")
}

function Apply-N2DataGridView {
    param([System.Windows.Forms.DataGridView]$Grid)
    $Grid.BorderStyle = [System.Windows.Forms.BorderStyle]::None
    $Grid.CellBorderStyle = [System.Windows.Forms.DataGridViewCellBorderStyle]::SingleHorizontal
    $Grid.GridColor = $script:Ui.GridLine
    $Grid.BackgroundColor = $script:Ui.Surface
    $Grid.EnableHeadersVisualStyles = $false
    $Grid.ColumnHeadersBorderStyle = [System.Windows.Forms.DataGridViewHeaderBorderStyle]::None
    $Grid.RowHeadersVisible = $false
    $Grid.ColumnHeadersHeight = 36
    $Grid.RowTemplate.Height = 28
    $Grid.DefaultCellStyle.BackColor = $script:Ui.Surface
    $Grid.DefaultCellStyle.ForeColor = $script:Ui.Text
    $Grid.DefaultCellStyle.Font = New-N2Font -Size 9
    $Grid.DefaultCellStyle.SelectionBackColor = $script:Ui.GridSelect
    $Grid.DefaultCellStyle.SelectionForeColor = $script:Ui.Text
    $Grid.AlternatingRowsDefaultCellStyle.BackColor = $script:Ui.SurfaceAlt
    $Grid.AlternatingRowsDefaultCellStyle.ForeColor = $script:Ui.Text
    $Grid.ColumnHeadersDefaultCellStyle.BackColor = $script:Ui.GridHeader
    $Grid.ColumnHeadersDefaultCellStyle.ForeColor = [System.Drawing.Color]::White
    $Grid.ColumnHeadersDefaultCellStyle.Font = New-N2Font -Size 8.5 -Style Bold
    $Grid.ColumnHeadersDefaultCellStyle.SelectionBackColor = $script:Ui.GridHeader
    # Default is Disable, which blocks Ctrl+C; tab-separated text for Notepad / Excel.
    $Grid.ClipboardCopyMode = [System.Windows.Forms.DataGridViewClipboardCopyMode]::EnableWithoutHeaderText
}

function Initialize-N2Button {
    param(
        [System.Windows.Forms.Button]$Button,
        [ValidateSet("Primary", "Secondary", "Ghost")][string]$Kind
    )
    $Button.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $Button.FlatAppearance.BorderSize = 0
    $Button.Cursor = [System.Windows.Forms.Cursors]::Hand
    $Button.Font = New-N2Font -Size 9 -Style Bold
    $Button.Height = 32
    switch ($Kind) {
        "Primary" {
            $Button.ForeColor = [System.Drawing.Color]::White
            $Button.BackColor = $script:Ui.Accent
            $Button.FlatAppearance.MouseOverBackColor = $script:Ui.AccentHover
            $Button.FlatAppearance.MouseDownBackColor = [System.Drawing.Color]::FromArgb(30, 64, 175)
        }
        "Secondary" {
            $Button.ForeColor = $script:Ui.Text
            $Button.BackColor = $script:Ui.Surface
            $Button.FlatAppearance.BorderSize = 1
            $Button.FlatAppearance.BorderColor = $script:Ui.Border
            $Button.FlatAppearance.MouseOverBackColor = $script:Ui.SurfaceAlt
        }
        "Ghost" {
            $Button.ForeColor = $script:Ui.Muted
            $Button.BackColor = $script:Ui.Surface
            $Button.FlatAppearance.BorderSize = 1
            $Button.FlatAppearance.BorderColor = $script:Ui.Border
            $Button.FlatAppearance.MouseOverBackColor = $script:Ui.SurfaceAlt
        }
    }
}

function Get-InitialHistory {
    @(
        [pscustomobject]@{ When = "2026-04-17 08:03"; Event = "Run completed"; Detail = "renewals_20260417_0800.csv - 24 rows (dry run)" }
        [pscustomobject]@{ When = "2026-04-17 02:02"; Event = "Run completed"; Detail = "renewals_20260417_0200.csv - 18 rows (dry run)" }
        [pscustomobject]@{ When = "2026-04-16 20:05"; Event = "Run completed"; Detail = "renewals_20260416_2000.csv - 31 rows, 4 lookup failures" }
        [pscustomobject]@{ When = "2026-04-16 20:05"; Event = "Service started"; Detail = "UI session (mock)" }
    )
}

$script:History = [System.Collections.Generic.List[object]]::new()
(Get-InitialHistory) | ForEach-Object { $script:History.Add($_) }

function Get-NowStamp {
    (Get-Date).ToString("yyyy-MM-dd HH:mm")
}

function Set-ConnectButton {
    param($Button)
    if ($script:Connected) {
        $Button.Text = "Disconnect"
        $Button.ForeColor = $script:Ui.Danger
        $Button.BackColor = $script:Ui.DangerSoft
        $Button.FlatAppearance.BorderSize = 1
        $Button.FlatAppearance.BorderColor = [System.Drawing.Color]::FromArgb(252, 165, 165)
        $Button.FlatAppearance.MouseOverBackColor = [System.Drawing.Color]::FromArgb(254, 226, 226)
    } else {
        $Button.Text = "Connect"
        $Button.ForeColor = [System.Drawing.Color]::White
        $Button.BackColor = $script:Ui.Accent
        $Button.FlatAppearance.BorderSize = 0
        $Button.FlatAppearance.BorderColor = $script:Ui.Accent
        $Button.FlatAppearance.MouseOverBackColor = $script:Ui.AccentHover
        $Button.FlatAppearance.MouseDownBackColor = [System.Drawing.Color]::FromArgb(30, 64, 175)
    }
}

function Set-GridRows {
    param($Grid, $Objects)
    $Grid.Rows.Clear()
    foreach ($o in $Objects) {
        $null = $Grid.Rows.Add(@($o.PSObject.Properties.Value))
    }
}

function Refresh-ServiceState {
    param($lblDry, $btnConn)
    if ($script:Connected -and $null -ne $script:Net2Connection) {
        try {
            $ds = $script:Net2Connection.DataSource
            $db = $script:Net2Connection.Database
            $net2 = "Connected ($ds / $db)"
        } catch {
            $net2 = "Connected"
        }
    } else {
        $net2 = "Disconnected"
    }
    $lblDry.Text = "Safe mode: no automated writes to NET2  |  SQL: $net2"
    if ($null -ne $btnConn) { Set-ConnectButton -Button $btnConn }
}

function Refresh-HistoryGrid {
    param($Grid, [int]$MaxRows = 0)
    if ($null -eq $Grid) { return }
    if ($MaxRows -gt 0) {
        Set-GridRows -Grid $Grid -Objects @($script:History | Select-Object -First $MaxRows)
    } else {
        Set-GridRows -Grid $Grid -Objects $script:History
    }
}

function Refresh-AllLogGrids {
    param($GridService, $GridHistory)
    Refresh-HistoryGrid -Grid $GridService -MaxRows 200
    Refresh-HistoryGrid -Grid $GridHistory
}

function Add-HistoryEntry {
    param([string]$Event, [string]$Detail)
    $script:History.Insert(0, [pscustomobject]@{ When = (Get-NowStamp); Event = $Event; Detail = $Detail })
}

function Resolve-N2WatchFolderPath {
    param([string]$Raw)
    $t = if ($null -eq $Raw) { "" } else { $Raw.Trim() }
    if ([string]::IsNullOrWhiteSpace($t)) { return $null }
    try {
        return [System.IO.Path]::GetFullPath($t)
    } catch {
        return $null
    }
}

function Get-N2CsvIdentifierPropertyName {
    param([string[]]$PropertyNames)
    if ($null -eq $PropertyNames -or $PropertyNames.Count -eq 0) { return $null }
    $candidates = @(
        "GuiNumber", "GUINumber", "Gui No", "GuiNo", "GUI", "UserID", "User Id", "UserId",
        "CardHolderId", "CardHolderID", "CardNumber", "Card Number", "Token",
        "PayrollNumber", "Payroll No", "PayrollNo", "EmployeeNumber", "Employee No", "EmployeeNo",
        "PersonnelNumber", "Personnel No", "StaffCode", "Staff Code", "EmpCode", "EmployeeCode",
        "Reference", "RefNo", "Badge", "BadgeNumber", "SitePersonnelNo", "HRID", "PersonnelCode",
        "Id", "ID"
    )
    $norm = @{}
    foreach ($p in $PropertyNames) {
        if ([string]::IsNullOrWhiteSpace($p)) { continue }
        $norm[$p.Trim()] = $p.Trim()
    }
    foreach ($c in $candidates) {
        foreach ($k in $norm.Keys) {
            if ([string]::Equals($k, $c, [StringComparison]::OrdinalIgnoreCase)) {
                return $k
            }
        }
    }
    return ($PropertyNames | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -First 1)
}

function Format-N2SqlBracketId {
    param([string]$Name)
    if ([string]::IsNullOrWhiteSpace($Name)) { return $null }
    return ('[' + (($Name.Trim()) -replace '\]', ']]') + ']')
}

function Unformat-N2SqlBracketId {
    param([string]$Quoted)
    if ([string]::IsNullOrWhiteSpace($Quoted)) { return $Quoted }
    $s = $Quoted.Trim()
    if ($s.Length -ge 2 -and $s.StartsWith("[") -and $s.EndsWith("]")) {
        return ($s.Substring(1, $s.Length - 2) -replace '\]\]', ']')
    }
    return $s
}

function Get-Net2MatchColumnsOrderedBySampleHit {
    param(
        [System.Data.SqlClient.SqlConnection]$Connection,
        [string]$FromQuoted,
        [string[]]$MatchColsQuoted,
        [string]$Sample,
        [int]$MaxColumnsToProbe = 28
    )
    $hits = New-Object System.Collections.Generic.List[string]
    $rest = New-Object System.Collections.Generic.List[string]
    $n = 0
    foreach ($cq in $MatchColsQuoted) {
        if ($null -eq $cq -or [string]::IsNullOrWhiteSpace($cq)) { continue }
        $n++
        if ($n -gt $MaxColumnsToProbe) {
            $null = $rest.Add($cq)
            continue
        }
        try {
            $cmd = $Connection.CreateCommand()
            $cmd.CommandTimeout = 5
            $null = $cmd.Parameters.Add("@probe", [System.Data.SqlDbType]::NVarChar, 128).Value = $Sample
            $cmd.CommandText = "SELECT TOP 1 1 AS n FROM $FromQuoted WHERE CAST($cq AS nvarchar(128)) = @probe"
            $scalar = $cmd.ExecuteScalar()
            if ($null -ne $scalar) {
                $null = $hits.Add($cq)
            } else {
                $null = $rest.Add($cq)
            }
        } catch {
            $null = $rest.Add($cq)
        }
    }
    $hitNames = foreach ($h in $hits) { Unformat-N2SqlBracketId $h }
    $sum = if ($hits.Count -gt 0) {
        "First NET2 probe for sample '" + $Sample + "' hit column(s): " + ($hitNames -join ", ")
    } else {
        "First NET2 probe for sample '" + $Sample + "' did not match a single column alone (still checking all columns in batch)"
    }
    return [pscustomobject]@{
        MatchCols  = @($hits.ToArray() + $rest.ToArray())
        HitSummary = $sum
    }
}

function Test-N2FromQuotedIsUsersLike {
    param([string]$FromQuoted)
    if ([string]::IsNullOrWhiteSpace($FromQuoted)) { return $false }
    $seg = ($FromQuoted -split "\.")[-1].Trim()
    $tbl = Unformat-N2SqlBracketId $seg
    if ([string]::IsNullOrWhiteSpace($tbl)) { return $false }
    return ($tbl -match "^(Users|ExpandedUsers|User)$")
}

function Add-Net2GuiUserDetailSection {
    param(
        [System.Data.SqlClient.SqlConnection]$Conn,
        [System.Collections.Generic.IEnumerable[string]]$MatchedGuiValues,
        [System.Text.StringBuilder]$Rb,
        [string]$UsersKeyColQuoted = "[UserID]",
        [string]$CardsUserFkColQuoted = "[UserID]",
        [string]$UsersDetailFromQuoted = "[dbo].[Users]",
        [string]$CardsDetailFromQuoted = "[dbo].[Cards]"
    )
    if ($null -eq $MatchedGuiValues) { return }
    $lst = @($MatchedGuiValues)
    if ($lst.Count -eq 0) { return }
    [void]$Rb.AppendLine("")
    [void]$Rb.AppendLine("--- Matched users (detail: read-only) ---")
    foreach ($gv in $lst) {
        if ([string]::IsNullOrWhiteSpace($gv)) { continue }
        $cmd = $Conn.CreateCommand()
        $cmd.CommandTimeout = 25
        $null = $cmd.Parameters.Add("@v", [System.Data.SqlDbType]::NVarChar, 128).Value = $gv.Trim()
        $cmd.CommandText = @"
SELECT TOP 1
  u.FirstName,
  u.Surname,
  d.Name,
  u.ActivateDate,
  u.ExpiryDate,
  al.Name,
  (SELECT COUNT(*) FROM $CardsDetailFromQuoted c WHERE c.$CardsUserFkColQuoted = u.$UsersKeyColQuoted)
FROM $UsersDetailFromQuoted u
LEFT JOIN dbo.Departments d ON d.DepartmentID = u.DepartmentID
LEFT JOIN dbo.[Access levels] al ON al.AccessLevelID = u.AccessLevelID
WHERE CAST(u.$UsersKeyColQuoted AS nvarchar(128)) = @v
   OR u.Field1_100 = @v OR u.Field2_100 = @v OR u.Field3_50 = @v OR u.Field4_50 = @v OR u.Field5_50 = @v
   OR u.Field6_50 = @v OR u.Field7_50 = @v OR u.Field8_50 = @v OR u.Field9_50 = @v OR u.Field10_50 = @v
   OR u.Field11_50 = @v OR u.Field12_50 = @v OR u.Field14_50 = @v
"@
        try {
            $r = $cmd.ExecuteReader()
            try {
                if (-not $r.Read()) {
                    [void]$Rb.AppendLine("GUI: " + $gv + " | (no Users row for detail lookup)")
                    continue
                }
                $fn = if ($r.IsDBNull(0)) { "" } else { [string]$r.GetValue(0) }
                $sn = if ($r.IsDBNull(1)) { "" } else { [string]$r.GetValue(1) }
                $dept = if ($r.IsDBNull(2)) { "" } else { [string]$r.GetValue(2) }
                $vf = ""
                if (-not $r.IsDBNull(3)) {
                    $vf = $r.GetDateTime(3).ToString("yyyy-MM-dd")
                }
                $ed = ""
                if (-not $r.IsDBNull(4)) {
                    $ed = $r.GetDateTime(4).ToString("yyyy-MM-dd")
                }
                $al = if ($r.IsDBNull(5)) { "" } else { [string]$r.GetValue(5) }
                $tok = 0
                if (-not $r.IsDBNull(6)) { $tok = [int]$r.GetValue(6) }
                [void]$Rb.AppendLine(
                    "GUI: " + $gv +
                    " | First name: " + $fn.Trim() +
                    " | Last name: " + $sn.Trim() +
                    " | Department: " + $dept.Trim() +
                    " | Valid from: " + $vf +
                    " | End date: " + $ed +
                    " | Access level: " + $al.Trim() +
                    " | Tokens (cards): " + $tok
                )
            } finally {
                $r.Close()
            }
        } catch {
            [void]$Rb.AppendLine("GUI: " + $gv + " | Detail error: " + $_.Exception.Message)
        }
    }
}

function Test-N2SqlColumnNameIsUserIdKey {
    param([string]$ColName)
    if ([string]::IsNullOrWhiteSpace($ColName)) { return $false }
    $norm = ([regex]::Replace($ColName.Trim(), "[\s_]+", "")).ToUpperInvariant()
    return ($norm -eq "USERID")
}

function Get-Net2SqlUserIdColumnPredicate {
    return "UPPER(REPLACE(REPLACE(LTRIM(RTRIM(c.name)), N' ', N''), N'_', N'')) = N'USERID'"
}

function Get-Net2MatchableColumnWhitelist {
    return @(
        "UserID", "UserId", "User ID",
        "PayrollNumber", "PayrollNo", "Payroll No", "Payroll",
        "EmployeeNumber", "EmployeeNo", "Employee No", "EmployeeID", "Employee Id",
        "PersonnelNumber", "PersonnelNo", "Personnel No", "PersonnelID", "Personnel Id",
        "StaffCode", "Staff Code", "EmpCode", "EmployeeCode", "Employee Code", "PersonnelCode",
        "Reference", "RefNo", "ReferenceNo", "ExternalReference", "ExtReference",
        "CardNumber", "CardNo", "Card No", "Token", "TokenNo", "TokenNumber",
        "SerialNumber", "BadgeNumber", "Badge", "IDCard", "IdCard", "SitePersonnelNo", "HRID", "HRRef",
        "Code", "UserCode", "CompanyNumber", "AccountName", "LoginName",
        "Field1", "Field2", "Field3", "Field4", "Field5", "TextField1", "TextField2",
        "Custom1", "Custom2", "DeisterNumber", "MifareID"
    )
}

function Test-N2SqlColumnIsLookupCandidate {
    param(
        [string]$ColName,
        [string]$TypeName,
        [int]$MaxLength,
        [int]$ColumnId
    )
    if ([string]::IsNullOrWhiteSpace($ColName)) { return $false }
    $n = $ColName.Trim()
    $t = if ([string]::IsNullOrWhiteSpace($TypeName)) { "" } else { $TypeName.Trim().ToLowerInvariant() }
    $uid = Test-N2SqlColumnNameIsUserIdKey -ColName $n
    if ($uid) {
        return ($t -in @("int", "bigint", "smallint", "tinyint", "varchar", "nvarchar", "char", "nchar", "decimal", "numeric"))
    }
    if ($t -notin @("varchar", "nvarchar", "char", "nchar")) { return $false }
    $nl = $n.ToLowerInvariant()
    $excludeExact = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    foreach ($x in @(
            "FirstName", "MiddleName", "Surname", "LastName", "Title", "Initials", "KnownAs", "Salutation",
            "Email", "E-mail", "Address1", "Address2", "Address3", "Address4", "Address", "Town", "City",
            "County", "Country", "PostCode", "PostalCode", "ZipCode", "Telephone", "Phone", "Mobile", "Fax",
            "Department", "AccessLevel", "SiteName", "Building", "Floor", "Zone", "Notes", "Description",
            "Photo", "Picture", "Image", "Password", "PIN", "Pin", "UserImage", "UserGuid", "RowGuid",
            "Created", "Modified", "LastModified", "ExpiryDate", "ActiveDate", "DateOfBirth", "DOB",
            "Gender", "Nationality", "Language", "TimeZone", "CardType", "Status", "Enabled", "Deleted",
            "Forename", "Extension", "Web", "URL", "URI", "Signature", "Template", "Blob", "Binary", "Xml"
        )) {
        $null = $excludeExact.Add($x)
    }
    if ($excludeExact.Contains($n)) { return $false }
    foreach ($pat in @("address", "email", "telephone", "phone", "mobile", "fax", "photo", "picture", "image", "password", "signature", "template", "notes", "description", "department", "accesslevel", "postcode", "postal", "country", "firstname", "surname", "lastname", "forename", "middlename", "expiry", "created", "modified", "guid")) {
        if ($nl.Contains($pat)) { return $false }
    }
    if ($MaxLength -eq -1) {
        return ($nl -match "(number|code|ref|payroll|employee|badge|token|card|field|custom|site|personnel|staff|emp|serial|login|account|external|gateway|import|export|key|id$)")
    }
    $minBytes = if ($t.StartsWith("n")) { 18 } else { 9 }
    if ($MaxLength -lt $minBytes) { return $false }
    if ($MaxLength -gt 800) { return $false }
    return $true
}

function Get-Net2QualifiedUsersTable {
    param([System.Data.SqlClient.SqlConnection]$Conn)
    $cmd = $Conn.CreateCommand()
    $cmd.CommandTimeout = 18
    $readPair = {
        param($SqlText)
        $cmd.CommandText = $SqlText
        try {
            $r = $cmd.ExecuteReader()
            try {
                if (-not $r.Read()) { return $null }
                $s0 = if ($r.IsDBNull(0)) { "" } else { ([string]$r.GetValue(0)).Trim() }
                $o0 = if ($r.IsDBNull(1)) { "" } else { ([string]$r.GetValue(1)).Trim() }
                if ($s0.Length -eq 0 -or $o0.Length -eq 0) { return $null }
                return [pscustomobject]@{ SchemaName = $s0; ObjectName = $o0 }
            } finally {
                $r.Close()
            }
        } catch {
            return $null
        }
    }
    $sqlNamed = @"
SELECT TOP 1 s.name AS sch, o.name AS obj
FROM sys.objects o
INNER JOIN sys.schemas s ON s.schema_id = o.schema_id
WHERE o.type = N'U' AND o.is_ms_shipped = 0
  AND o.name COLLATE Latin1_General_CI_AI IN (
    N'Users', N'User', N'tblUsers', N'tblUser', N't_Users',
    N'Operators', N'Operator', N'tblOperator', N'tblOperators',
    N'Personnel', N'Employees', N'NET2_Users', N'PaxtonUsers', N'Cardholders', N'Cardholder'
  )
ORDER BY
  CASE WHEN s.name = N'dbo' AND o.name COLLATE Latin1_General_CI_AI = N'Users' THEN 0
       WHEN o.name COLLATE Latin1_General_CI_AI = N'Users' THEN 1
       WHEN o.name COLLATE Latin1_General_CI_AI LIKE N'%user%' THEN 2
       WHEN s.name = N'dbo' THEN 3 ELSE 4 END,
  s.name, o.name
"@
    $hit = & $readPair $sqlNamed
    if ($null -ne $hit) { return $hit }
    $sqlHeuristic = @"
SELECT TOP 1 s.name AS sch, o.name AS obj
FROM sys.objects o
INNER JOIN sys.schemas s ON s.schema_id = o.schema_id
WHERE o.type = N'U' AND o.is_ms_shipped = 0
  AND EXISTS (
    SELECT 1 FROM sys.columns c
    WHERE c.object_id = o.object_id
      AND c.name COLLATE Latin1_General_CI_AI IN (N'FirstName', N'Forename', N'GivenName')
  )
  AND EXISTS (
    SELECT 1 FROM sys.columns c
    WHERE c.object_id = o.object_id
      AND c.name COLLATE Latin1_General_CI_AI IN (N'Surname', N'LastName', N'Last_Name', N'FamilyName')
  )
ORDER BY
  CASE WHEN o.name COLLATE Latin1_General_CI_AI LIKE N'%user%' THEN 0
       WHEN o.name COLLATE Latin1_General_CI_AI LIKE N'%person%' THEN 1
       WHEN o.name COLLATE Latin1_General_CI_AI LIKE N'%operator%' THEN 2
       WHEN o.name COLLATE Latin1_General_CI_AI LIKE N'%employee%' THEN 3
       WHEN o.name COLLATE Latin1_General_CI_AI LIKE N'%member%' THEN 4
       ELSE 9 END,
  s.name, o.name
"@
    return (& $readPair $sqlHeuristic)
}

function Get-Net2QualifiedCardsTable {
    param([System.Data.SqlClient.SqlConnection]$Conn)
    $cmd = $Conn.CreateCommand()
    $cmd.CommandTimeout = 12
    $cmd.CommandText = @"
SELECT TOP 1 s.name AS sch, o.name AS obj
FROM sys.objects o
INNER JOIN sys.schemas s ON s.schema_id = o.schema_id
WHERE o.type = N'U' AND o.is_ms_shipped = 0
  AND o.name COLLATE Latin1_General_CI_AI IN (N'Cards', N'Card', N'tblCards', N'tblCard')
ORDER BY
  CASE WHEN s.name = N'dbo' AND o.name COLLATE Latin1_General_CI_AI = N'Cards' THEN 0
       WHEN o.name COLLATE Latin1_General_CI_AI = N'Cards' THEN 1
       WHEN s.name = N'dbo' THEN 2 ELSE 3 END,
  s.name, o.name
"@
    try {
        $r = $cmd.ExecuteReader()
        try {
            if (-not $r.Read()) { return $null }
            $s0 = if ($r.IsDBNull(0)) { "" } else { ([string]$r.GetValue(0)).Trim() }
            $o0 = if ($r.IsDBNull(1)) { "" } else { ([string]$r.GetValue(1)).Trim() }
            if ($s0.Length -eq 0 -or $o0.Length -eq 0) { return $null }
            return [pscustomobject]@{ SchemaName = $s0; ObjectName = $o0 }
        } finally {
            $r.Close()
        }
    } catch {
        return $null
    }
}

function Get-Net2UsersTableKeyColumnName {
    param(
        [System.Data.SqlClient.SqlConnection]$Conn,
        [string]$SchemaName,
        [string]$ObjectName
    )
    if ([string]::IsNullOrWhiteSpace($SchemaName) -or [string]::IsNullOrWhiteSpace($ObjectName)) { return $null }
    $sch = $SchemaName.Trim()
    $obj = $ObjectName.Trim()
    $pred = Get-Net2SqlUserIdColumnPredicate
    $cmd = $Conn.CreateCommand()
    $cmd.CommandTimeout = 12
    $null = $cmd.Parameters.Add("@usch", [System.Data.SqlDbType]::NVarChar, 128).Value = $sch
    $null = $cmd.Parameters.Add("@uobj", [System.Data.SqlDbType]::NVarChar, 128).Value = $obj
    $cmd.CommandText = @"
SELECT TOP 1 c.name
FROM sys.columns c
WHERE c.object_id = OBJECT_ID(QUOTENAME(@usch) + N'.' + QUOTENAME(@uobj)) AND $pred
ORDER BY c.column_id
"@
    try {
        $x = $cmd.ExecuteScalar()
        if ($null -ne $x) {
            $t = ([string]$x).Trim()
            if ($t.Length -gt 0) { return $t }
        }
    } catch { }
    $cmd.CommandText = @"
SELECT TOP 1 c.name
FROM sys.indexes i
INNER JOIN sys.index_columns ic ON i.object_id = ic.object_id AND i.index_id = ic.index_id
INNER JOIN sys.columns c ON c.object_id = ic.object_id AND c.column_id = ic.column_id
WHERE i.object_id = OBJECT_ID(QUOTENAME(@usch) + N'.' + QUOTENAME(@uobj)) AND i.is_primary_key = 1
ORDER BY ic.key_ordinal
"@
    try {
        $x = $cmd.ExecuteScalar()
        if ($null -ne $x) {
            $t = ([string]$x).Trim()
            if ($t.Length -gt 0) { return $t }
        }
    } catch { }
    $cmd.CommandText = @"
SELECT TOP 1 c.name
FROM sys.columns c
WHERE c.object_id = OBJECT_ID(QUOTENAME(@usch) + N'.' + QUOTENAME(@uobj)) AND c.is_identity = 1
ORDER BY c.column_id
"@
    try {
        $x = $cmd.ExecuteScalar()
        if ($null -ne $x) {
            $t = ([string]$x).Trim()
            if ($t.Length -gt 0) { return $t }
        }
    } catch { }
    $cmd.CommandText = @"
SELECT TOP 1 c.name
FROM sys.columns c
WHERE c.object_id = OBJECT_ID(QUOTENAME(@usch) + N'.' + QUOTENAME(@uobj))
  AND c.name COLLATE Latin1_General_CI_AI IN (N'Id', N'ID')
ORDER BY c.column_id
"@
    try {
        $x = $cmd.ExecuteScalar()
        if ($null -ne $x) {
            $t = ([string]$x).Trim()
            if ($t.Length -gt 0) { return $t }
        }
    } catch { }
    $cmd.CommandText = @"
SELECT TOP 1 c.name
FROM sys.columns c
INNER JOIN sys.types ty ON ty.user_type_id = c.user_type_id
WHERE c.object_id = OBJECT_ID(QUOTENAME(@usch) + N'.' + QUOTENAME(@uobj))
  AND c.column_id = 1 AND c.is_computed = 0
  AND ty.name IN (N'int', N'bigint', N'smallint', N'tinyint', N'uniqueidentifier')
"@
    try {
        $x = $cmd.ExecuteScalar()
        if ($null -ne $x) {
            $t = ([string]$x).Trim()
            if ($t.Length -gt 0) { return $t }
        }
    } catch { }
    return $null
}

function Get-Net2DboCardsUserFkColumnName {
    param([System.Data.SqlClient.SqlConnection]$Conn)
    $ct = Get-Net2QualifiedCardsTable -Conn $Conn
    if ($null -eq $ct) { return $null }
    $pred = Get-Net2SqlUserIdColumnPredicate
    $cmd = $Conn.CreateCommand()
    $cmd.CommandTimeout = 10
    $null = $cmd.Parameters.Add("@csch", [System.Data.SqlDbType]::NVarChar, 128).Value = $ct.SchemaName
    $null = $cmd.Parameters.Add("@cobj", [System.Data.SqlDbType]::NVarChar, 128).Value = $ct.ObjectName
    $cmd.CommandText = @"
SELECT TOP 1 c.name
FROM sys.columns c
WHERE c.object_id = OBJECT_ID(QUOTENAME(@csch) + N'.' + QUOTENAME(@cobj)) AND $pred
ORDER BY c.column_id
"@
    try {
        $x = $cmd.ExecuteScalar()
        if ($null -eq $x) { return $null }
        $t = ([string]$x).Trim()
        if ($t.Length -eq 0) { return $null }
        return $t
    } catch {
        return $null
    }
}

function Get-Net2DboUsersKeyColumnName {
    param([System.Data.SqlClient.SqlConnection]$Conn)
    $ut = Get-Net2QualifiedUsersTable -Conn $Conn
    if ($null -eq $ut) { return $null }
    return (Get-Net2UsersTableKeyColumnName -Conn $Conn -SchemaName $ut.SchemaName -ObjectName $ut.ObjectName)
}

function New-Net2LookupResolutionFromQualifiedObject {
    param(
        [System.Data.SqlClient.SqlConnection]$Conn,
        [string]$SchemaName,
        [string]$ObjectName,
        [string]$ForcedUserKeyColumnName = $null
    )
    if ([string]::IsNullOrWhiteSpace($SchemaName) -or [string]::IsNullOrWhiteSpace($ObjectName)) { return $null }
    $sch = $SchemaName.Trim()
    $obj = $ObjectName.Trim()
    $cmd = $Conn.CreateCommand()
    $cmd.CommandTimeout = 15
    $white = Get-Net2MatchableColumnWhitelist
    $uidPred = Get-Net2SqlUserIdColumnPredicate
    $colName = ""
    if (-not [string]::IsNullOrWhiteSpace($ForcedUserKeyColumnName)) {
        $fcp = $ForcedUserKeyColumnName.Trim()
        $cmd.Parameters.Clear()
        $null = $cmd.Parameters.Add("@sch", [System.Data.SqlDbType]::NVarChar, 128).Value = $sch
        $null = $cmd.Parameters.Add("@obj", [System.Data.SqlDbType]::NVarChar, 128).Value = $obj
        $null = $cmd.Parameters.Add("@fkc", [System.Data.SqlDbType]::NVarChar, 128).Value = $fcp
        $cmd.CommandText = @"
SELECT TOP 1 c.name
FROM sys.columns c
WHERE c.object_id = OBJECT_ID(QUOTENAME(@sch) + N'.' + QUOTENAME(@obj))
  AND c.name COLLATE Latin1_General_CI_AI = @fkc
"@
        try {
            $cr = $cmd.ExecuteScalar()
            if ($null -ne $cr) { $colName = ([string]$cr).Trim() }
        } catch {
            return $null
        }
        if ($colName.Length -eq 0) { return $null }
    } else {
        $cmd.Parameters.Clear()
        $null = $cmd.Parameters.Add("@sch", [System.Data.SqlDbType]::NVarChar, 128).Value = $sch
        $null = $cmd.Parameters.Add("@obj", [System.Data.SqlDbType]::NVarChar, 128).Value = $obj
        $cmd.CommandText = @"
SELECT TOP 1 c.name
FROM sys.objects o
INNER JOIN sys.schemas s ON s.schema_id = o.schema_id
INNER JOIN sys.columns c ON c.object_id = o.object_id
WHERE s.name = @sch AND o.name = @obj AND o.type IN (N'U', N'V')
  AND $uidPred
ORDER BY c.column_id
"@
        $colRaw = $null
        try {
            $colRaw = $cmd.ExecuteScalar()
        } catch {
            return $null
        }
        $colName = if ($null -eq $colRaw) { "" } else { ([string]$colRaw).Trim() }
        if ($colName.Length -eq 0) { return $null }
    }
    $fromQ = (Format-N2SqlBracketId $sch) + "." + (Format-N2SqlBracketId $obj)
    $cmd.Parameters.Clear()
    $null = $cmd.Parameters.Add("@sch2", [System.Data.SqlDbType]::NVarChar, 128).Value = $sch
    $null = $cmd.Parameters.Add("@obj2", [System.Data.SqlDbType]::NVarChar, 128).Value = $obj
    $cmd.CommandText = @"
SELECT c.name, ty.name AS typname, c.max_length, c.column_id
FROM sys.columns c
INNER JOIN sys.types ty ON ty.user_type_id = c.user_type_id
WHERE c.object_id = OBJECT_ID(QUOTENAME(@sch2) + N'.' + QUOTENAME(@obj2))
  AND c.is_computed = 0
ORDER BY c.column_id
"@
    $candRows = New-Object System.Collections.Generic.List[object]
    try {
        $r = $cmd.ExecuteReader()
        try {
            while ($r.Read()) {
                if ($r.IsDBNull(0)) { continue }
                $nm = ([string]$r.GetValue(0)).Trim()
                $tn = if ($r.IsDBNull(1)) { "" } else { ([string]$r.GetValue(1)).Trim() }
                $ml2 = 0
                if (-not $r.IsDBNull(2)) {
                    $rawMl = $r.GetValue(2)
                    try {
                        $ml2 = [System.Convert]::ToInt32($rawMl)
                    } catch {
                        $ml2 = 0
                    }
                }
                $cid2 = 0
                if (-not $r.IsDBNull(3)) {
                    $cid2 = [int]$r.GetValue(3)
                }
                $null = $candRows.Add([pscustomobject]@{
                        Name       = $nm
                        TypeName   = $tn
                        MaxLength  = $ml2
                        ColumnId   = $cid2
                    })
            }
        } finally {
            $r.Close()
        }
    } catch {
        return $null
    }
    $foundNames = New-Object System.Collections.Generic.List[string]
    foreach ($row in $candRows) {
        if (Test-N2SqlColumnIsLookupCandidate -ColName $row.Name -TypeName $row.TypeName -MaxLength $row.MaxLength -ColumnId $row.ColumnId) {
            $null = $foundNames.Add($row.Name)
        }
    }
    if ($foundNames.Count -eq 0) {
        $null = $foundNames.Add($colName)
    }
    $sortable = New-Object System.Collections.Generic.List[object]
    foreach ($name in ($foundNames | Select-Object -Unique)) {
        $ix = 9999
        for ($wi = 0; $wi -lt $white.Count; $wi++) {
            if ([string]::Equals($name, $white[$wi], [StringComparison]::OrdinalIgnoreCase)) {
                $ix = $wi
                break
            }
        }
        $cid3 = 99999
        foreach ($row in $candRows) {
            if ([string]::Equals($row.Name, $name, [StringComparison]::OrdinalIgnoreCase)) {
                $cid3 = $row.ColumnId
                break
            }
        }
        $null = $sortable.Add([pscustomobject]@{ Name = $name; SortKey = ("{0:0000}-{1:0000}" -f $ix, $cid3) })
    }
    $maxCols = 45
    $sorted = @(
        $sortable |
        Sort-Object { $_.SortKey } |
        Select-Object -ExpandProperty Name |
        Select-Object -First $maxCols
    )
    $matchColsQ = New-Object System.Collections.Generic.List[string]
    foreach ($sn in $sorted) {
        $bq = Format-N2SqlBracketId $sn
        if ($null -ne $bq -and -not $matchColsQ.Contains($bq)) { $null = $matchColsQ.Add($bq) }
    }
    if ($matchColsQ.Count -eq 0) {
        $null = $matchColsQ.Add((Format-N2SqlBracketId $colName))
    }
    $tail = if (@($sorted).Count -gt 6) { ",..." } else { "" }
    $shortLabel = $sch + "." + $obj + "[" + (($sorted | Select-Object -First 6) -join ",") + $tail + "]"
    $usersDetailKeyQ = Format-N2SqlBracketId $colName
    if (-not (Test-N2FromQuotedIsUsersLike -FromQuoted $fromQ)) {
        $uK = Get-Net2DboUsersKeyColumnName -Conn $Conn
        if ($null -ne $uK -and $uK.Length -gt 0) {
            $usersDetailKeyQ = Format-N2SqlBracketId $uK
        }
    }
    $cardsFkNm = Get-Net2DboCardsUserFkColumnName -Conn $Conn
    if ($null -eq $cardsFkNm -or $cardsFkNm.Length -eq 0) { $cardsFkNm = "UserID" }
    $cardsFkQ = Format-N2SqlBracketId $cardsFkNm
    $utPhys = Get-Net2QualifiedUsersTable -Conn $Conn
    $usersDetailFromQ = if ($null -ne $utPhys) {
        (Format-N2SqlBracketId $utPhys.SchemaName) + "." + (Format-N2SqlBracketId $utPhys.ObjectName)
    } else {
        "[dbo].[Users]"
    }
    $ctPhys = Get-Net2QualifiedCardsTable -Conn $Conn
    $cardsDetailFromQ = if ($null -ne $ctPhys) {
        (Format-N2SqlBracketId $ctPhys.SchemaName) + "." + (Format-N2SqlBracketId $ctPhys.ObjectName)
    } else {
        "[dbo].[Cards]"
    }
    return [pscustomobject]@{
        Ok                      = $true
        FromQuoted              = $fromQ
        MatchColsQuoted         = @($matchColsQ.ToArray())
        SummaryLabel            = $shortLabel
        Error                   = $null
        UsersDetailKeyQuoted    = $usersDetailKeyQ
        CardsUserFkQuoted       = $cardsFkQ
        UsersDetailFromQuoted   = $usersDetailFromQ
        CardsDetailFromQuoted   = $cardsDetailFromQ
    }
}

function New-Net2LookupResolutionFromUsersPrimaryKey {
    param([System.Data.SqlClient.SqlConnection]$Conn)
    $ut = Get-Net2QualifiedUsersTable -Conn $Conn
    if ($null -eq $ut) { return $null }
    $sch = $ut.SchemaName
    $obj = $ut.ObjectName
    $cmd = $Conn.CreateCommand()
    $cmd.CommandTimeout = 18
    $null = $cmd.Parameters.Add("@sch", [System.Data.SqlDbType]::NVarChar, 128).Value = $sch
    $null = $cmd.Parameters.Add("@obj", [System.Data.SqlDbType]::NVarChar, 128).Value = $obj
    $pred = Get-Net2SqlUserIdColumnPredicate
    $nm = ""
    $cmd.CommandText = @"
SELECT TOP 1 c.name
FROM sys.indexes i
INNER JOIN sys.index_columns ic ON i.object_id = ic.object_id AND i.index_id = ic.index_id
INNER JOIN sys.columns c ON c.object_id = ic.object_id AND c.column_id = ic.column_id
WHERE i.object_id = OBJECT_ID(QUOTENAME(@sch) + N'.' + QUOTENAME(@obj)) AND i.is_primary_key = 1
ORDER BY ic.key_ordinal
"@
    try {
        $x = $cmd.ExecuteScalar()
        if ($null -ne $x) { $nm = ([string]$x).Trim() }
    } catch { }
    if ($nm.Length -eq 0) {
        $cmd.CommandText = @"
SELECT TOP 1 c.name
FROM sys.columns c
WHERE c.object_id = OBJECT_ID(QUOTENAME(@sch) + N'.' + QUOTENAME(@obj)) AND c.is_identity = 1
ORDER BY c.column_id
"@
        try {
            $x = $cmd.ExecuteScalar()
            if ($null -ne $x) { $nm = ([string]$x).Trim() }
        } catch { }
    }
    if ($nm.Length -eq 0) {
        $cmd.CommandText = @"
SELECT TOP 1 c.name
FROM sys.columns c
WHERE c.object_id = OBJECT_ID(QUOTENAME(@sch) + N'.' + QUOTENAME(@obj))
  AND c.name COLLATE Latin1_General_CI_AI IN (N'Id', N'ID')
ORDER BY c.column_id
"@
        try {
            $x = $cmd.ExecuteScalar()
            if ($null -ne $x) { $nm = ([string]$x).Trim() }
        } catch { }
    }
    if ($nm.Length -eq 0) {
        $cmd.CommandText = @"
SELECT TOP 1 c.name
FROM sys.columns c
WHERE c.object_id = OBJECT_ID(QUOTENAME(@sch) + N'.' + QUOTENAME(@obj)) AND $pred
ORDER BY c.column_id
"@
        try {
            $x = $cmd.ExecuteScalar()
            if ($null -ne $x) { $nm = ([string]$x).Trim() }
        } catch { }
    }
    if ($nm.Length -eq 0) {
        $cmd.CommandText = @"
SELECT TOP 1 c.name
FROM sys.columns c
INNER JOIN sys.types ty ON ty.user_type_id = c.user_type_id
WHERE c.object_id = OBJECT_ID(QUOTENAME(@sch) + N'.' + QUOTENAME(@obj))
  AND c.column_id = 1 AND c.is_computed = 0
  AND ty.name IN (N'int', N'bigint', N'smallint', N'tinyint', N'uniqueidentifier')
"@
        try {
            $x = $cmd.ExecuteScalar()
            if ($null -ne $x) { $nm = ([string]$x).Trim() }
        } catch { }
    }
    if ($nm.Length -eq 0) { return $null }
    return (New-Net2LookupResolutionFromQualifiedObject -Conn $Conn -SchemaName $sch -ObjectName $obj -ForcedUserKeyColumnName $nm)
}

function Get-Net2UserIdBearingObjectDiscovery {
    param([System.Data.SqlClient.SqlConnection]$Conn)
    $uidPred = Get-Net2SqlUserIdColumnPredicate
    $cmd = $Conn.CreateCommand()
    $cmd.CommandTimeout = 20
    $cmd.CommandText = @"
SELECT TOP 1 s.name AS sch, o.name AS obj
FROM sys.columns c
INNER JOIN sys.objects o ON o.object_id = c.object_id AND o.type IN (N'U', N'V') AND o.is_ms_shipped = 0
INNER JOIN sys.schemas s ON s.schema_id = o.schema_id
WHERE $uidPred
ORDER BY
  CASE
    WHEN s.name = N'dbo' AND o.name = N'ExpandedUsers' THEN 0
    WHEN s.name = N'dbo' AND o.name = N'Users' THEN 1
    WHEN s.name = N'dbo' AND o.name = N'User' THEN 2
    WHEN o.name = N'ExpandedUsers' THEN 3
    WHEN o.name = N'Users' THEN 4
    WHEN o.name = N'User' THEN 5
    WHEN o.type = N'V' AND (o.name LIKE N'%Expanded%' OR o.name LIKE N'%User%') THEN 8
    WHEN o.name IN (N'Cardholders', N'Cardholder') THEN 12
    WHEN o.name = N'Cards' THEN 15
    ELSE 20
  END,
  s.name,
  o.name
"@
    try {
        $r = $cmd.ExecuteReader()
        try {
            if (-not $r.Read()) { return $null }
            $s0 = ([string]$r.GetValue(0)).Trim()
            $o0 = ([string]$r.GetValue(1)).Trim()
            if ($s0.Length -eq 0 -or $o0.Length -eq 0) { return $null }
            return [pscustomobject]@{ SchemaName = $s0; ObjectName = $o0 }
        } finally {
            $r.Close()
        }
    } catch {
        return $null
    }
}

function Resolve-Net2UserIdLookupSource {
    param([System.Data.SqlClient.SqlConnection]$Conn)
    $objCandidates = @(
        @{ s = "dbo"; n = "ExpandedUsers" },
        @{ s = "dbo"; n = "Users" },
        @{ s = "dbo"; n = "Cards" },
        @{ s = "dbo"; n = "vwExpandedUsers" },
        @{ s = "dbo"; n = "vw_ExpandedUsers" },
        @{ s = "dbo"; n = "User" },
        @{ s = "dbo"; n = "tblUsers" },
        @{ s = "dbo"; n = "tblUser" },
        @{ s = "dbo"; n = "Cardholders" },
        @{ s = "dbo"; n = "Cardholder" }
    )
    foreach ($o in $objCandidates) {
        $built = New-Net2LookupResolutionFromQualifiedObject -Conn $Conn -SchemaName $o.s -ObjectName $o.n
        if ($null -ne $built) { return $built }
    }
    $disc = Get-Net2UserIdBearingObjectDiscovery -Conn $Conn
    if ($null -ne $disc) {
        $built2 = New-Net2LookupResolutionFromQualifiedObject -Conn $Conn -SchemaName $disc.SchemaName -ObjectName $disc.ObjectName
        if ($null -ne $built2) { return $built2 }
    }
    $builtPk = New-Net2LookupResolutionFromUsersPrimaryKey -Conn $Conn
    if ($null -ne $builtPk) { return $builtPk }
    $ds = ""
    $dbn = ""
    try { if ($null -ne $Conn.DataSource) { $ds = [string]$Conn.DataSource } } catch { }
    try { if ($null -ne $Conn.Database) { $dbn = [string]$Conn.Database } } catch { }
    return [pscustomobject]@{
        Ok               = $false
        FromQuoted       = $null
        MatchColsQuoted  = $null
        SummaryLabel     = $null
        Error            = @"
No NET2-style user lookup on this connection (expanded table names + name heuristic were tried).

Server: $ds
Database: $dbn

Check the UI uses the same database as SSMS (often Net2, never master). In SSMS: SELECT DB_NAME();

If tables exist but names differ, run:
SELECT TOP 40 s.name AS schema_name, o.name AS object_name FROM sys.objects o JOIN sys.schemas s ON s.schema_id = o.schema_id WHERE o.type IN ('U','V') AND o.is_ms_shipped = 0 ORDER BY o.name;
"@
    }
}

function Get-Net2ExpandedUserIdsMatching {
    param(
        [System.Data.SqlClient.SqlConnection]$Connection,
        [string[]]$Ids,
        [int]$ChunkSize = 80
    )
    $out = [pscustomobject]@{
        Found         = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
        HadError      = $false
        ErrorMessage  = $null
        SourceLabel   = $null
        ProbeNote     = $null
        ReportText    = $null
    }
    if ($null -eq $Connection -or $Connection.State -ne [System.Data.ConnectionState]::Open) {
        $out.HadError = $true
        $out.ErrorMessage = "SQL connection is not open."
        $out.ReportText = "=== NET2 read-only check ===`r`nNot connected. Use Connect, then Manual sync."
        return $out
    }
    $cacheOk = (
        $null -ne $script:Net2UserLookupCache -and
        $script:Net2UserLookupCache.DataSource -eq $Connection.DataSource -and
        $script:Net2UserLookupCache.Database -eq $Connection.Database -and
        $null -ne $script:Net2UserLookupCache.MatchColsQuoted -and
        $script:Net2UserLookupCache.MatchColsQuoted.Count -gt 0 -and
        $script:Net2UserLookupCache.SchemaRevision -eq $script:Net2SqlLookupSchemaRev
    )
    if (-not $cacheOk) {
        $script:Net2UserLookupCache = $null
        $res = Resolve-Net2UserIdLookupSource -Conn $Connection
        if (-not $res.Ok) {
            $out.HadError = $true
            $out.ErrorMessage = $res.Error
            $out.ReportText = "=== NET2 read-only check ===`r`n" + $res.Error
            return $out
        }
        $script:Net2UserLookupCache = [pscustomobject]@{
            DataSource           = $Connection.DataSource
            Database             = $Connection.Database
            FromQuoted           = $res.FromQuoted
            MatchColsQuoted      = @($res.MatchColsQuoted)
            SummaryLabel         = $res.SummaryLabel
            SchemaRevision       = $script:Net2SqlLookupSchemaRev
            UsersDetailKeyQuoted    = $(if ($null -ne $res.UsersDetailKeyQuoted -and $res.UsersDetailKeyQuoted.Length -gt 0) { $res.UsersDetailKeyQuoted } else { "[UserID]" })
            CardsUserFkQuoted       = $(if ($null -ne $res.CardsUserFkQuoted -and $res.CardsUserFkQuoted.Length -gt 0) { $res.CardsUserFkQuoted } else { "[UserID]" })
            UsersDetailFromQuoted   = $(if ($null -ne $res.UsersDetailFromQuoted -and $res.UsersDetailFromQuoted.Length -gt 0) { $res.UsersDetailFromQuoted } else { "[dbo].[Users]" })
            CardsDetailFromQuoted   = $(if ($null -ne $res.CardsDetailFromQuoted -and $res.CardsDetailFromQuoted.Length -gt 0) { $res.CardsDetailFromQuoted } else { "[dbo].[Cards]" })
        }
    }
    $fromQ = $script:Net2UserLookupCache.FromQuoted
    $matchCols = @($script:Net2UserLookupCache.MatchColsQuoted)
    $out.SourceLabel = $script:Net2UserLookupCache.SummaryLabel
    $list = New-Object System.Collections.Generic.List[string]
    foreach ($id in $Ids) {
        if ($null -eq $id) { continue }
        $t = $id.Trim()
        if ($t.Length -eq 0) { continue }
        if ($t.Length -gt 128) { $t = $t.Substring(0, 128) }
        $null = $list.Add($t)
    }
    if ($list.Count -eq 0) {
        $out.ReportText = "=== NET2 read-only check ===`r`nNo non-blank CSV identifiers were supplied."
        return $out
    }
    $probeSample = $null
    foreach ($id in $Ids) {
        if ($null -eq $id) { continue }
        $tp = $id.Trim()
        if ($tp.Length -lt 4) { continue }
        if ($tp -match "[A-Za-z]") {
            $probeSample = $tp
            break
        }
    }
    if ($null -eq $probeSample) {
        $probeSample = $list[0]
    }
    if ($null -ne $probeSample -and $probeSample.Length -gt 0 -and $matchCols.Count -gt 0) {
        $pr = Get-Net2MatchColumnsOrderedBySampleHit -Connection $Connection -FromQuoted $fromQ -MatchColsQuoted $matchCols -Sample $probeSample
        $matchCols = @($pr.MatchCols)
        $out.ProbeNote = $pr.HitSummary
    }
    for ($off = 0; $off -lt $list.Count; $off += $ChunkSize) {
        $take = [Math]::Min($ChunkSize, $list.Count - $off)
        $cmd = $Connection.CreateCommand()
        $cmd.CommandTimeout = 30
        $names = New-Object System.Collections.Generic.List[string]
        for ($i = 0; $i -lt $take; $i++) {
            $pn = "@p$i"
            $names.Add($pn)
            $v = $list[$off + $i]
            $null = $cmd.Parameters.Add($pn, [System.Data.SqlDbType]::NVarChar, 128).Value = $v
        }
        $inList = $names -join ", "
        $sqlParts = New-Object System.Collections.Generic.List[string]
        foreach ($cq in $matchCols) {
            $null = $sqlParts.Add(
                "SELECT CAST($cq AS nvarchar(128)) AS k FROM $fromQ WHERE CAST($cq AS nvarchar(128)) IN ($inList)"
            )
        }
        $cmd.CommandText = ($sqlParts -join " UNION ALL ")
        try {
            $r = $cmd.ExecuteReader()
            try {
                while ($r.Read()) {
                    if (-not $r.IsDBNull(0)) {
                        $null = $out.Found.Add([string]$r.GetValue(0))
                    }
                }
            } finally {
                $r.Close()
            }
        } catch {
            $out.HadError = $true
            $out.ErrorMessage = $_.Exception.Message
            $script:Net2UserLookupCache = $null
            break
        }
    }
    $colLine = ($matchCols | ForEach-Object { Unformat-N2SqlBracketId $_ }) -join ", "
    $uniqIds = New-Object System.Collections.Generic.List[string]
    $seenId = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    foreach ($id in $Ids) {
        if ($null -eq $id) { continue }
        $t = $id.Trim()
        if ($t.Length -eq 0) { continue }
        if ($seenId.Add($t)) { $null = $uniqIds.Add($t) }
    }
    $foundList = New-Object System.Collections.Generic.List[string]
    $missList = New-Object System.Collections.Generic.List[string]
    foreach ($id in $uniqIds) {
        if ($out.Found.Contains($id)) {
            $null = $foundList.Add($id)
        } else {
            $null = $missList.Add($id)
        }
    }
    $rb = New-Object System.Text.StringBuilder
    [void]$rb.AppendLine("=== NET2 read-only check (no writes) ===")
    [void]$rb.AppendLine("Object: " + $fromQ)
    [void]$rb.AppendLine("Columns scanned (" + $matchCols.Count + "): " + $colLine)
    [void]$rb.AppendLine("")
    if (-not [string]::IsNullOrWhiteSpace($out.ProbeNote)) {
        [void]$rb.AppendLine($out.ProbeNote)
        [void]$rb.AppendLine("")
    }
    if ($out.HadError) {
        [void]$rb.AppendLine("Error: " + $out.ErrorMessage)
    } else {
        [void]$rb.AppendLine("--- Matched in NET2 (" + $foundList.Count + ") ---")
        foreach ($x in $foundList) {
            [void]$rb.AppendLine("  " + $x)
        }
        [void]$rb.AppendLine("")
        [void]$rb.AppendLine("--- Not found in NET2 (" + $missList.Count + ") ---")
        foreach ($x in $missList) {
            [void]$rb.AppendLine("  " + $x)
        }
        if ($foundList.Count -gt 0 -and (Test-N2FromQuotedIsUsersLike -FromQuoted $fromQ)) {
            $uk = $script:Net2UserLookupCache.UsersDetailKeyQuoted
            $ck = $script:Net2UserLookupCache.CardsUserFkQuoted
            $uf = $script:Net2UserLookupCache.UsersDetailFromQuoted
            $cf = $script:Net2UserLookupCache.CardsDetailFromQuoted
            if ([string]::IsNullOrWhiteSpace($uk)) { $uk = "[UserID]" }
            if ([string]::IsNullOrWhiteSpace($ck)) { $ck = "[UserID]" }
            if ([string]::IsNullOrWhiteSpace($uf)) { $uf = "[dbo].[Users]" }
            if ([string]::IsNullOrWhiteSpace($cf)) { $cf = "[dbo].[Cards]" }
            Add-Net2GuiUserDetailSection -Conn $Connection -MatchedGuiValues $foundList -Rb $rb -UsersKeyColQuoted $uk -CardsUserFkColQuoted $ck -UsersDetailFromQuoted $uf -CardsDetailFromQuoted $cf
        }
    }
    $out.ReportText = $rb.ToString()
    return $out
}

function Invoke-N2CursorNet2CsvReport {
    param(
        [string]$FolderRaw,
        [string]$Pattern,
        [System.Data.SqlClient.SqlConnection]$Connection
    )
    $result = [pscustomobject]@{
        ReportText = ""
        HadError     = $false
    }
    $folderPath = Resolve-N2WatchFolderPath -Raw $FolderRaw
    if ($null -eq $folderPath -or -not (Test-Path -LiteralPath $folderPath -PathType Container)) {
        $result.ReportText = "=== NET2 read-only check (Cursor / console) ===`r`nInvalid watch folder: " + $FolderRaw
        $result.HadError = $true
        return $result
    }
    $pat = $Pattern.Trim()
    if ([string]::IsNullOrWhiteSpace($pat)) { $pat = "*.csv" }
    try {
        $files = @(Get-ChildItem -LiteralPath $folderPath -File -Filter $pat -ErrorAction Stop)
    } catch {
        $result.ReportText = "=== NET2 read-only check (Cursor / console) ===`r`n" + $_.Exception.Message
        $result.HadError = $true
        return $result
    }
    if ($files.Count -eq 0) {
        $result.ReportText = "=== NET2 read-only check (Cursor / console) ===`r`nNo files matching '$pat' in:`r`n$folderPath"
        $result.HadError = $true
        return $result
    }
    $maxIds = 5000
    $allIds = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    foreach ($f in $files) {
        try {
            $rows = @(Import-Csv -LiteralPath $f.FullName -Encoding UTF8 -ErrorAction Stop)
        } catch {
            continue
        }
        if ($rows.Count -eq 0) { continue }
        $props = @($rows[0].PSObject.Properties.Name)
        $idCol = Get-N2CsvIdentifierPropertyName -PropertyNames $props
        if ([string]::IsNullOrWhiteSpace($idCol)) { continue }
        foreach ($row in $rows) {
            $v = $row.$idCol
            if ($null -eq $v) { continue }
            $s = ([string]$v).Trim()
            if ($s.Length -eq 0) { continue }
            if ($s.Length -gt 128) { $s = $s.Substring(0, 128) }
            if ($allIds.Count -ge $maxIds) { break }
            $null = $allIds.Add($s)
        }
        if ($allIds.Count -ge $maxIds) { break }
    }
    if ($allIds.Count -eq 0) {
        $result.ReportText = "=== NET2 read-only check (Cursor / console) ===`r`nNo identifiers collected from CSV(s)."
        $result.HadError = $true
        return $result
    }
    $match = Get-Net2ExpandedUserIdsMatching -Connection $Connection -Ids @($allIds)
    $result.ReportText = if ($null -ne $match.ReportText) { $match.ReportText } else { "" }
    $result.HadError = $match.HadError
    return $result
}

if ($CursorReport) {
    Close-Net2SqlConnection
    $script:Net2UserLookupCache = $null
    try {
        Write-Host ("Connecting to NET2 SQL (read-only): " + $effSqlServer + " / " + $effSqlDatabase)
        Write-Host ("CSV folder: " + $effWatchFolder + "  pattern: " + $effFilePattern)
        $cs = Get-Net2ConnectionString -Server $effSqlServer -Database $effSqlDatabase -SqlUser "" -SqlPassword ""
        $conn = New-Object System.Data.SqlClient.SqlConnection($cs)
        $conn.Open()
        $script:Net2Connection = $conn
        $script:Connected = $true
        Write-Host "Connected. Running NET2 vs CSV id check (SELECT only, no writes)..."
        $rpt = Invoke-N2CursorNet2CsvReport -FolderRaw $effWatchFolder -Pattern $effFilePattern -Connection $conn
        Write-Host ""
        Write-Host $rpt.ReportText
        if ($rpt.HadError) { exit 1 }
        exit 0
    } catch {
        Write-Host ""
        Write-Host "=== NET2 connection or read failed ==="
        Write-Host $_.Exception.Message
        if ($null -ne $_.Exception.InnerException) { Write-Host $_.Exception.InnerException.Message }
        Write-Host ""
        Write-Host "Check: server name (SSMS), Windows login on SQL, firewall, database name. Optional N2SYNC.local.json next to script."
        exit 1
    } finally {
        Close-Net2SqlConnection
    }
}

# --- Form ---
$form = New-Object System.Windows.Forms.Form
$form.Text = "Merdian Paxton Bridge"
$form.Size = New-Object System.Drawing.Size(1160, 760)
$form.MinimumSize = New-Object System.Drawing.Size(920, 580)
$form.StartPosition = "CenterScreen"
$form.BackColor = $script:Ui.FormBg
$form.Font = New-N2Font -Size 9
$lblTitle = New-Object System.Windows.Forms.Label
$lblTitle.AutoSize = $true
$lblTitle.Location = New-Object System.Drawing.Point(16, 14)
$lblTitle.Font = New-N2Font -Size 18 -Style Bold
$lblTitle.ForeColor = $script:Ui.Text
$lblTitle.Text = "Merdian Paxton Bridge"

$lblSub = New-Object System.Windows.Forms.Label
$lblSub.AutoSize = $true
$lblSub.Location = New-Object System.Drawing.Point(16, 46)
$lblSub.ForeColor = $script:Ui.Muted
$lblSub.Font = New-N2Font -Size 9.25
$lblSub.Text = "CSV drops, Paxton NET2 checks, renewals - bridge UI (mock data only)."

$pnlAccent = New-Object System.Windows.Forms.Panel
$pnlAccent.Location = New-Object System.Drawing.Point(16, 70)
$pnlAccent.Size = New-Object System.Drawing.Size(72, 3)
$pnlAccent.BackColor = $script:Ui.Accent

$tabControl = New-Object System.Windows.Forms.TabControl
$tabControl.Location = New-Object System.Drawing.Point(16, 82)
$tabControl.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$tabControl.Size = New-Object System.Drawing.Size(($form.ClientSize.Width - 32), ($form.ClientSize.Height - 94))
# TabControl.Padding is Point on some WinForms stacks; skip tab chrome padding to stay compatible.
$tabControl.Font = New-N2Font -Size 9 -Style Bold
$tabControl.SizeMode = [System.Windows.Forms.TabSizeMode]::Fixed
$tabControl.ItemSize = New-Object System.Drawing.Size(112, 32)
$tabControl.Appearance = [System.Windows.Forms.TabAppearance]::Normal
$tabControl.BackColor = $script:Ui.FormBg
$tabControl.Margin = New-Object System.Windows.Forms.Padding(0)

$tabService = New-Object System.Windows.Forms.TabPage
$tabService.Text = "Service"
$tabService.BackColor = $script:Ui.Surface
$tabService.UseVisualStyleBackColor = $false

$tabConfig = New-Object System.Windows.Forms.TabPage
$tabConfig.Text = "Configuration"
$tabConfig.BackColor = $script:Ui.FormBg
$tabConfig.UseVisualStyleBackColor = $false

$tabHistory = New-Object System.Windows.Forms.TabPage
$tabHistory.Text = "History"
$tabHistory.BackColor = $script:Ui.Surface
$tabHistory.UseVisualStyleBackColor = $false

$tabSupport = New-Object System.Windows.Forms.TabPage
$tabSupport.Text = "Support"
$tabSupport.BackColor = $script:Ui.Surface
$tabSupport.UseVisualStyleBackColor = $false

$tabControl.TabPages.Add($tabService)
$tabControl.TabPages.Add($tabConfig)
$tabControl.TabPages.Add($tabHistory)
$tabControl.TabPages.Add($tabSupport)

# --- Service tab ---
$lblDry = New-Object System.Windows.Forms.Label
$lblDry.AutoSize = $true
$lblDry.Location = New-Object System.Drawing.Point(16, 18)
$lblDry.Font = New-N2Font -Size 9.25 -Style Bold
$lblDry.ForeColor = $script:Ui.Text
$lblDry.AutoEllipsis = $true

$btnConn = New-Object System.Windows.Forms.Button
$btnConn.Text = "Connect"
$btnConn.Width = 108
$btnConn.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
Initialize-N2Button -Button $btnConn -Kind Primary

$btnManualSync = New-Object System.Windows.Forms.Button
$btnManualSync.Text = "Manual sync"
$btnManualSync.Width = 126
$btnManualSync.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
Initialize-N2Button -Button $btnManualSync -Kind Secondary

$btnClearLogs = New-Object System.Windows.Forms.Button
$btnClearLogs.Text = "Clear logs"
$btnClearLogs.Width = 104
$btnClearLogs.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
Initialize-N2Button -Button $btnClearLogs -Kind Ghost

function Sync-TopRightButtons {
    param($Parent, $ConnBtn, $ManualSyncBtn, $ClearBtn)
    $m = 16
    $g = 10
    $w = $Parent.ClientSize.Width
    $y = 14
    $ClearBtn.Location = New-Object System.Drawing.Point(($w - $m - $ClearBtn.Width), $y)
    $ManualSyncBtn.Location = New-Object System.Drawing.Point(($w - $m - $ClearBtn.Width - $g - $ManualSyncBtn.Width), $y)
    $ConnBtn.Location = New-Object System.Drawing.Point(($w - $m - $ClearBtn.Width - $g - $ManualSyncBtn.Width - $g - $ConnBtn.Width), $y)
}

$pnlServiceStrip = New-Object System.Windows.Forms.Panel
$pnlServiceStrip.Location = New-Object System.Drawing.Point(16, 16)
$pnlServiceStrip.Height = 52
$pnlServiceStrip.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$pnlServiceStrip.BackColor = $script:Ui.AccentSoft
$pnlServiceStrip.Controls.Add($lblDry)
$pnlServiceStrip.Controls.Add($btnConn)
$pnlServiceStrip.Controls.Add($btnManualSync)
$pnlServiceStrip.Controls.Add($btnClearLogs)

$lblServiceLogs = New-Object System.Windows.Forms.Label
$lblServiceLogs.AutoSize = $true
$lblServiceLogs.Location = New-Object System.Drawing.Point(16, 78)
$lblServiceLogs.ForeColor = $script:Ui.Muted
$lblServiceLogs.Font = New-N2Font -Size 8.75 -Style Bold
$lblServiceLogs.Text = "Recent activity (newest first, up to 200 lines). Select row(s), then Ctrl+C to copy."

$dgServiceLog = New-Object System.Windows.Forms.DataGridView
$dgServiceLog.ReadOnly = $true
$dgServiceLog.AllowUserToAddRows = $false
$dgServiceLog.AllowUserToDeleteRows = $false
$dgServiceLog.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
$dgServiceLog.MultiSelect = $true
$dgServiceLog.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::Fill
$dgServiceLog.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$dgServiceLog.Location = New-Object System.Drawing.Point(16, 100)
$dgServiceLog.Size = New-Object System.Drawing.Size(800, 260)

$lblSearchPreview = New-Object System.Windows.Forms.Label
$lblSearchPreview.AutoSize = $false
$lblSearchPreview.Location = New-Object System.Drawing.Point(16, 368)
$lblSearchPreview.Size = New-Object System.Drawing.Size(800, 20)
$lblSearchPreview.Anchor = [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$lblSearchPreview.ForeColor = $script:Ui.Muted
$lblSearchPreview.Font = New-N2Font -Size 8.75 -Style Bold
$lblSearchPreview.Text = "What we found in NET2 (read-only display - no database changes)"

$tbSearchPreview = New-Object System.Windows.Forms.TextBox
$tbSearchPreview.Multiline = $true
$tbSearchPreview.ReadOnly = $true
$tbSearchPreview.ScrollBars = [System.Windows.Forms.ScrollBars]::Vertical
$tbSearchPreview.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
$tbSearchPreview.BackColor = $script:Ui.SurfaceAlt
$tbSearchPreview.ForeColor = $script:Ui.Text
$tbSearchPreview.Font = New-N2Font -Size 9
$tbSearchPreview.Location = New-Object System.Drawing.Point(16, 392)
$tbSearchPreview.Size = New-Object System.Drawing.Size(800, 120)
$tbSearchPreview.Anchor = [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$tbSearchPreview.TabStop = $false
$tbSearchPreview.Text = "Run Manual sync while connected to NET2 to see probe details, NET2 columns checked, and which CSV ids matched or did not match."
foreach ($c in @("When", "Event", "Detail")) {
    $null = $dgServiceLog.Columns.Add($c, $c)
}
$dgServiceLog.Columns["When"].FillWeight = 22
$dgServiceLog.Columns["Event"].FillWeight = 20
$dgServiceLog.Columns["Detail"].FillWeight = 58
Apply-N2DataGridView -Grid $dgServiceLog

$tabService.Controls.AddRange(@(
    $pnlServiceStrip, $lblServiceLogs, $dgServiceLog, $lblSearchPreview, $tbSearchPreview
))

function Sync-ServiceTabLayout {
    $m = 16
    $previewH = [Math]::Min(240, [Math]::Max(96, [int]($tabService.ClientSize.Height * 0.28)))
    $tbSearchPreview.Height = $previewH
    $tbSearchPreview.Width = [Math]::Max(200, $tabService.ClientSize.Width - ($m * 2))
    $tbSearchPreview.Left = $m
    $tbSearchPreview.Top = $tabService.ClientSize.Height - $m - $tbSearchPreview.Height
    $lblSearchPreview.Width = $tbSearchPreview.Width
    $lblSearchPreview.Left = $m
    $lblSearchPreview.Top = $tbSearchPreview.Top - 22
    $dgServiceLog.Width = $tbSearchPreview.Width
    $dgServiceLog.Height = [Math]::Max(100, $lblSearchPreview.Top - $dgServiceLog.Top - 6)
}

# --- Configuration tab ---
function New-CfgSectionLabel {
    param([int]$Y, [string]$Text, [int]$InnerWidth)
    $h = New-Object System.Windows.Forms.Label
    $h.AutoSize = $false
    $h.Location = New-Object System.Drawing.Point(20, $Y)
    $h.Size = New-Object System.Drawing.Size([Math]::Max(200, $InnerWidth - 40), 24)
    $h.Text = $Text
    $h.Font = New-N2Font -Size 10 -Style Bold
    $h.ForeColor = $script:Ui.Accent
    return $h
}

function New-CfgRow {
    param(
        [int]$Y,
        [string]$LabelText,
        [string]$BoxText,
        [int]$BoxWidth,
        [switch]$Editable,
        [switch]$Password,
        [switch]$WideTextBox
    )
    $lab = New-Object System.Windows.Forms.Label
    $lab.AutoSize = $false
    $lab.Size = New-Object System.Drawing.Size(184, 24)
    $lab.Location = New-Object System.Drawing.Point(20, $Y)
    $lab.Text = $LabelText
    $lab.TextAlign = [System.Drawing.ContentAlignment]::MiddleRight
    $lab.ForeColor = $script:Ui.Text
    $lab.Font = New-N2Font -Size 9
    $tb = New-Object System.Windows.Forms.TextBox
    $tb.Location = New-Object System.Drawing.Point(212, $Y)
    $tb.Size = New-Object System.Drawing.Size($BoxWidth, 24)
    $tb.Text = $BoxText
    $tb.ReadOnly = -not $Editable
    $tb.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    $tb.BackColor = if ($Editable) { $script:Ui.Surface } else { $script:Ui.SurfaceAlt }
    $tb.ForeColor = $script:Ui.Text
    $tb.Font = New-N2Font -Size 9
    if ($WideTextBox) {
        $tb.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
    }
    if ($Password) { $tb.UseSystemPasswordChar = $true }
    return @($lab, $tb)
}

$pnlCfgCard = New-Object System.Windows.Forms.Panel
$pnlCfgCard.Location = New-Object System.Drawing.Point(14, 12)
$pnlCfgCard.Size = New-Object System.Drawing.Size(920, 500)
$pnlCfgCard.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$pnlCfgCard.BackColor = $script:Ui.Surface
$pnlCfgCard.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle

$yCfg = 18
$lblCfgHdr = New-Object System.Windows.Forms.Label
$lblCfgHdr.AutoSize = $false
$lblCfgHdr.Location = New-Object System.Drawing.Point(20, $yCfg)
$lblCfgHdr.Size = New-Object System.Drawing.Size(860, 36)
$lblCfgHdr.Font = New-N2Font -Size 11 -Style Bold
$lblCfgHdr.ForeColor = $script:Ui.Text
$lblCfgHdr.Text = "Bridge settings"
$yCfg += 42

$lblSecIn = New-CfgSectionLabel -Y $yCfg -Text "Inbound CSV" -InnerWidth 880
$yCfg += 30
$pairA = New-CfgRow -Y $yCfg -LabelText "Watch folder" -BoxText "D:\Inbound\NET2Sync\" -BoxWidth 400 -Editable -WideTextBox
$pairA[1].Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left

$btnBrowseWatch = New-Object System.Windows.Forms.Button
$btnBrowseWatch.Text = "Browse..."
$btnBrowseWatch.Width = 100
$btnBrowseWatch.Height = 26
$btnBrowseWatch.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Right
$btnBrowseWatch.Font = New-N2Font -Size 9
$btnBrowseWatch.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
$btnBrowseWatch.FlatAppearance.BorderColor = $script:Ui.Border
$btnBrowseWatch.BackColor = $script:Ui.SurfaceAlt
$btnBrowseWatch.ForeColor = $script:Ui.Text
$btnBrowseWatch.Cursor = [System.Windows.Forms.Cursors]::Hand
$btnBrowseWatch.FlatAppearance.MouseOverBackColor = $script:Ui.Surface

$btnBrowseWatch.add_Click({
    $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
    $dlg.Description = "Select the folder to watch for incoming CSV files."
    $cur = $pairA[1].Text.Trim()
    if ($cur -and (Test-Path -LiteralPath $cur -PathType Container)) {
        $dlg.SelectedPath = $cur
    }
    if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $p = $dlg.SelectedPath
        if ($p -and -not ($p.EndsWith("\"))) { $p += "\" }
        $pairA[1].Text = $p
    }
})

$yCfg += 36
$pairB = New-CfgRow -Y $yCfg -LabelText "File pattern" -BoxText "*.csv" -BoxWidth 220 -Editable
$yCfg += 36
$pairC = New-CfgRow -Y $yCfg -LabelText "Drop interval (minutes)" -BoxText "60" -BoxWidth 100 -Editable
$yCfg += 44
$lblSecSql = New-CfgSectionLabel -Y $yCfg -Text "Paxton NET2 SQL" -InnerWidth 880
$yCfg += 30
$pairD = New-CfgRow -Y $yCfg -LabelText "SQL server" -BoxText ".\NET2" -BoxWidth 400 -Editable -WideTextBox
$yCfg += 36
$pairE = New-CfgRow -Y $yCfg -LabelText "Database name" -BoxText "Net2" -BoxWidth 220 -Editable
$yCfg += 40
$lblCfgNote = New-Object System.Windows.Forms.Label
$lblCfgNote.AutoSize = $false
$lblCfgNote.Location = New-Object System.Drawing.Point(20, $yCfg)
$lblCfgNote.Size = New-Object System.Drawing.Size(860, 96)
$lblCfgNote.ForeColor = $script:Ui.Text
$lblCfgNote.BackColor = $script:Ui.AccentSoft
$lblCfgNote.Font = New-N2Font -Size 9
$lblCfgNote.Text = "Server is the same value you would pass to sqlcmd -S (for example .\NET2 on the machine that hosts the instance). Same as SSMS. Connect uses Windows Authentication only (sqlcmd -E / Integrated Security), not SQL sa.`r`n`r`nManual sync (read-only): checks whether each CSV GUI / id exists in NET2; for matches on dbo.Users it then loads first name, last name, department, valid from / end date, access level, and token (card) count into the Service report panel.`r`n`r`nSettings are kept in memory only until you close the app. Trust Server Certificate is enabled for typical lab or self-signed SQL setups."

$pnlCfgCard.Controls.AddRange(@(
    $lblCfgHdr, $lblSecIn,
    $pairA[0], $pairA[1], $btnBrowseWatch, $pairB[0], $pairB[1], $pairC[0], $pairC[1],
    $lblSecSql, $pairD[0], $pairD[1], $pairE[0], $pairE[1],
    $lblCfgNote
))

$tabConfig.Controls.Add($pnlCfgCard)

$tabConfig.add_Resize({
    $pad = 14
    $pnlCfgCard.Width = [Math]::Max(360, $tabConfig.ClientSize.Width - ($pad * 2))
    $pnlCfgCard.Height = [Math]::Max(420, $tabConfig.ClientSize.Height - 24)
    $innerW = $pnlCfgCard.ClientSize.Width
    $btnGap = 10
    $btnW = $btnBrowseWatch.Width
    $btnX = $innerW - 20 - $btnW
    $btnBrowseWatch.Location = New-Object System.Drawing.Point($btnX, $pairA[1].Location.Y)
    $pairA[1].Width = [Math]::Max(140, $btnX - 212 - $btnGap)
    $rw = $innerW - 232
    if ($rw -lt 160) { $rw = 160 }
    $pairD[1].Width = $rw
    $lblCfgHdr.Width = $innerW - 40
    $lblSecIn.Width = $innerW - 40
    $lblSecSql.Width = $innerW - 40
    $lblCfgNote.Width = $innerW - 40
})

# --- History tab ---
$lblHistHint = New-Object System.Windows.Forms.Label
$lblHistHint.AutoSize = $true
$lblHistHint.Location = New-Object System.Drawing.Point(16, 14)
$lblHistHint.ForeColor = $script:Ui.Muted
$lblHistHint.Font = New-N2Font -Size 8.75
$lblHistHint.Text = "Complete archive (mock). The Service tab shows only the latest 200 lines. Select row(s), then Ctrl+C to copy."

$dgHistory = New-Object System.Windows.Forms.DataGridView
$dgHistory.ReadOnly = $true
$dgHistory.AllowUserToAddRows = $false
$dgHistory.AllowUserToDeleteRows = $false
$dgHistory.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
$dgHistory.MultiSelect = $true
$dgHistory.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::Fill
$dgHistory.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$dgHistory.Location = New-Object System.Drawing.Point(16, 40)
$dgHistory.Size = New-Object System.Drawing.Size(800, 400)
foreach ($c in @("When", "Event", "Detail")) {
    $null = $dgHistory.Columns.Add($c, $c)
}
$dgHistory.Columns["When"].FillWeight = 22
$dgHistory.Columns["Event"].FillWeight = 20
$dgHistory.Columns["Detail"].FillWeight = 58
Apply-N2DataGridView -Grid $dgHistory

$tabHistory.Controls.AddRange(@($lblHistHint, $dgHistory))

$tabHistory.add_Resize({
    $dgHistory.Width = [Math]::Max(200, $tabHistory.ClientSize.Width - 32)
    $dgHistory.Height = [Math]::Max(120, $tabHistory.ClientSize.Height - $dgHistory.Top - 20)
})

# --- Support tab ---
$lblSupTitle = New-Object System.Windows.Forms.Label
$lblSupTitle.AutoSize = $true
$lblSupTitle.Font = New-N2Font -Size 12 -Style Bold
$lblSupTitle.ForeColor = $script:Ui.Text
$lblSupTitle.Location = New-Object System.Drawing.Point(16, 18)
$lblSupTitle.Text = "About Merdian Paxton Bridge"

$lblSupBody = New-Object System.Windows.Forms.Label
$lblSupBody.AutoSize = $false
$lblSupBody.Location = New-Object System.Drawing.Point(16, 52)
$lblSupBody.Size = New-Object System.Drawing.Size(720, 220)
$lblSupBody.ForeColor = $script:Ui.Text
$lblSupBody.Font = New-N2Font -Size 9.25
$lblSupBody.Text = "Merdian Paxton Bridge is a front-end prototype: log entries and NET2 connectivity are simulated.`r`n`r`nBefore production you will need a secure watch folder, a defined CSV layout, NET2 SQL access that matches your IT policy, tested renewals in a sandbox, and backups.`r`n`r`nFor deployment help, contact your administrator or integration partner."

$tabSupport.Controls.AddRange(@($lblSupTitle, $lblSupBody))

$tabSupport.add_Resize({
    $lblSupBody.Width = [Math]::Max(280, $tabSupport.ClientSize.Width - 32)
})

$form.add_Resize({
    $tabControl.Width = [Math]::Max(400, $form.ClientSize.Width - 32)
    $tabControl.Height = [Math]::Max(220, $form.ClientSize.Height - 94)
    $pnlServiceStrip.Width = [Math]::Max(480, $tabService.ClientSize.Width - 32)
    $lblDry.MaximumSize = New-Object System.Drawing.Size([Math]::Max(200, $pnlServiceStrip.ClientSize.Width - 400), 40)
    Sync-TopRightButtons -Parent $pnlServiceStrip -ConnBtn $btnConn -ManualSyncBtn $btnManualSync -ClearBtn $btnClearLogs
    Sync-ServiceTabLayout
})

$btnConn.add_Click({
    if ($script:Connected) {
        Close-Net2SqlConnection
        Add-HistoryEntry -Event "NET2 SQL" -Detail "Disconnected"
        Refresh-ServiceState -lblDry $lblDry -btnConn $btnConn
        Refresh-AllLogGrids -GridService $dgServiceLog -GridHistory $dgHistory
        return
    }
    try {
        $cs = Get-Net2ConnectionString -Server $pairD[1].Text -Database $pairE[1].Text -SqlUser "" -SqlPassword ""
        $conn = New-Object System.Data.SqlClient.SqlConnection($cs)
        $conn.Open()
        $script:Net2Connection = $conn
        $script:Connected = $true
        $ver = ""
        try {
            $cmd = $conn.CreateCommand()
            $cmd.CommandText = "SELECT CAST(SERVERPROPERTY('ProductVersion') AS nvarchar(128))"
            $ver = [string]$cmd.ExecuteScalar()
        } catch { $ver = "?" }
        Add-HistoryEntry -Event "NET2 SQL" -Detail ("Connected: " + $conn.DataSource + " / " + $conn.Database + " (SQL " + $ver + ")")
    } catch {
        Close-Net2SqlConnection
        $msg = $_.Exception.Message
        if ($_.Exception.InnerException) { $msg += " " + $_.Exception.InnerException.Message }
        [void][System.Windows.Forms.MessageBox]::Show(
            $msg,
            "Cannot connect to NET2 SQL",
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning
        )
        $flat = ($msg -replace "[\r\n]+", " ").Trim()
        if ($flat.Length -gt 240) { $flat = $flat.Substring(0, 240) }
        Add-HistoryEntry -Event "NET2 SQL" -Detail ("Connection failed: " + $flat)
    }
    Refresh-ServiceState -lblDry $lblDry -btnConn $btnConn
    Refresh-AllLogGrids -GridService $dgServiceLog -GridHistory $dgHistory
})

$btnManualSync.add_Click({
    $folderPath = Resolve-N2WatchFolderPath -Raw $pairA[1].Text
    if ($null -eq $folderPath -or -not (Test-Path -LiteralPath $folderPath -PathType Container)) {
        $tbSearchPreview.Text = "Watch folder missing or not found.`r`nConfigured path: " + $pairA[1].Text
        Add-HistoryEntry -Event "Manual sync" -Detail "Watch folder missing or not found: $($pairA[1].Text)"
        Refresh-ServiceState -lblDry $lblDry -btnConn $btnConn
        Refresh-AllLogGrids -GridService $dgServiceLog -GridHistory $dgHistory
        return
    }
    $pat = $pairB[1].Text.Trim()
    if ([string]::IsNullOrWhiteSpace($pat)) { $pat = "*.csv" }
    $files = @()
    try {
        $files = @(Get-ChildItem -LiteralPath $folderPath -File -Filter $pat -ErrorAction Stop)
    } catch {
        $tbSearchPreview.Text = "Could not list CSV folder.`r`n" + $_.Exception.Message
        Add-HistoryEntry -Event "Manual sync" -Detail ("Cannot list CSV folder: " + $_.Exception.Message)
        Refresh-ServiceState -lblDry $lblDry -btnConn $btnConn
        Refresh-AllLogGrids -GridService $dgServiceLog -GridHistory $dgHistory
        return
    }
    if ($files.Count -eq 0) {
        $tbSearchPreview.Text = "No CSV files matched pattern '$pat' in:`r`n$folderPath"
        Refresh-ServiceState -lblDry $lblDry -btnConn $btnConn
        Refresh-AllLogGrids -GridService $dgServiceLog -GridHistory $dgHistory
        return
    }
    $sqlReady = $false
    if (
        $script:Connected -and
        $null -ne $script:Net2Connection -and
        $script:Net2Connection.State -eq [System.Data.ConnectionState]::Open
    ) {
        $sqlReady = $true
    }
    if (-not $sqlReady) {
        Add-HistoryEntry -Event "Manual sync" -Detail "Found $($files.Count) file(s) under watch folder. SQL user check skipped - use Connect first (read-only SELECT only when connected)."
    }
    $maxIds = 5000
    $allIds = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
    foreach ($f in $files) {
        $rows = @()
        try {
            $rows = @(Import-Csv -LiteralPath $f.FullName -Encoding UTF8 -ErrorAction Stop)
        } catch {
            Add-HistoryEntry -Event "Manual sync" -Detail ("CSV read failed: " + $f.Name + " - " + $_.Exception.Message)
            continue
        }
        if ($rows.Count -eq 0) {
            Add-HistoryEntry -Event "Manual sync" -Detail ("Empty or no rows: " + $f.Name)
            continue
        }
        $props = @($rows[0].PSObject.Properties.Name)
        $idCol = Get-N2CsvIdentifierPropertyName -PropertyNames $props
        if ([string]::IsNullOrWhiteSpace($idCol)) {
            Add-HistoryEntry -Event "Manual sync" -Detail ("No columns in CSV: " + $f.Name)
            continue
        }
        $fileIds = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::OrdinalIgnoreCase)
        foreach ($row in $rows) {
            $v = $row.$idCol
            if ($null -eq $v) { continue }
            $s = ([string]$v).Trim()
            if ($s.Length -eq 0) { continue }
            if ($s.Length -gt 128) { $s = $s.Substring(0, 128) }
            $null = $fileIds.Add($s)
        }
        $merged = 0
        foreach ($s in $fileIds) {
            if ($allIds.Count -ge $maxIds) { break }
            if ($allIds.Add($s)) { $merged++ }
        }
        $capHit = ($allIds.Count -ge $maxIds)
        $trunc = if ($capHit) { " (identifier cap $maxIds reached)" } else { "" }
        if ($sqlReady) {
            Add-HistoryEntry -Event "Manual sync" -Detail ("File " + $f.Name + ": " + $rows.Count + " row(s), column '" + $idCol + "', " + $fileIds.Count + " distinct id(s) this file, " + $merged + " new for check" + $trunc)
        } else {
            Add-HistoryEntry -Event "Manual sync" -Detail ("File " + $f.Name + ": " + $rows.Count + " row(s), id column '" + $idCol + "', " + $fileIds.Count + " distinct value(s) (SQL check not run)" + $trunc)
        }
        if ($capHit) { break }
    }
    if ($sqlReady -and $allIds.Count -gt 0) {
        $idArr = @($allIds)
        $match = Get-Net2ExpandedUserIdsMatching -Connection $script:Net2Connection -Ids $idArr
        if ($null -ne $match.ReportText -and $match.ReportText.Length -gt 0) {
            $tbSearchPreview.Text = $match.ReportText
        }
        if ($match.HadError) {
            $em = $match.ErrorMessage
            if ($null -ne $em) { $em = ($em -replace "[\r\n]+", " ").Trim() }
            if ([string]::IsNullOrWhiteSpace($em)) { $em = "(no message)" }
            elseif ($em.Length -gt 220) { $em = $em.Substring(0, 220) }
            Add-HistoryEntry -Event "Manual sync" -Detail ("SQL read-only check failed: " + $em)
        } else {
            $foundN = $match.Found.Count
            $missN = $allIds.Count - $foundN
            $src = if ($null -ne $match.SourceLabel -and $match.SourceLabel.Length -gt 0) { $match.SourceLabel } else { "user lookup" }
            $detailOk = "NET2 " + $src + " (read-only): " + $foundN + " of " + $allIds.Count + " distinct CSV id(s) matched; " + $missN + " not found."
            if ($null -ne $match.ProbeNote -and $match.ProbeNote.Length -gt 0) {
                $detailOk += " | " + $match.ProbeNote
            }
            Add-HistoryEntry -Event "Manual sync" -Detail $detailOk
        }
    } elseif ($sqlReady -and $allIds.Count -eq 0) {
        $tbSearchPreview.Text = "CSV files were read, but no non-blank identifiers were collected for a NET2 check."
        Add-HistoryEntry -Event "Manual sync" -Detail "No non-blank identifiers collected from CSV(s); SQL check not run."
    } elseif (-not $sqlReady -and $allIds.Count -gt 0) {
        $tbSearchPreview.Text = "CSV identifiers were collected (" + $allIds.Count + " distinct), but NET2 SQL was not connected.`r`nUse Connect, then Manual sync again to see matches here (read-only)."
    } elseif (-not $sqlReady -and $allIds.Count -eq 0) {
        $tbSearchPreview.Text = "Found " + $files.Count + " CSV file(s) under the watch folder.`r`nNo identifiers were collected from rows, or connect to NET2 first to compare ids here (read-only)."
    }
    Refresh-ServiceState -lblDry $lblDry -btnConn $btnConn
    Refresh-AllLogGrids -GridService $dgServiceLog -GridHistory $dgHistory
})

$btnClearLogs.add_Click({
    $script:History.Clear()
    $tbSearchPreview.Text = "Run Manual sync while connected to NET2 to see probe details, NET2 columns checked, and which CSV ids matched or did not match."
    Refresh-ServiceState -lblDry $lblDry -btnConn $btnConn
    Refresh-AllLogGrids -GridService $dgServiceLog -GridHistory $dgHistory
})

$form.add_FormClosing({
    Close-Net2SqlConnection
})

$form.Controls.AddRange(@($lblTitle, $lblSub, $pnlAccent, $tabControl))

$pnlServiceStrip.Width = [Math]::Max(480, $tabService.ClientSize.Width - 32)
$lblDry.MaximumSize = New-Object System.Drawing.Size([Math]::Max(200, $pnlServiceStrip.ClientSize.Width - 400), 40)
Sync-TopRightButtons -Parent $pnlServiceStrip -ConnBtn $btnConn -ManualSyncBtn $btnManualSync -ClearBtn $btnClearLogs
Sync-ServiceTabLayout
Refresh-ServiceState -lblDry $lblDry -btnConn $btnConn
Refresh-AllLogGrids -GridService $dgServiceLog -GridHistory $dgHistory
[void]$form.ShowDialog()
