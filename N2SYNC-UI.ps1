# Merdian Paxton Bridge - CSV to Paxton NET2 (UI prototype: mock data only, no database)
# Run: powershell -ExecutionPolicy Bypass -File ".\N2SYNC-UI.ps1"

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

[void][System.Windows.Forms.Application]::EnableVisualStyles()
[void][System.Windows.Forms.Application]::SetCompatibleTextRenderingDefault($false)

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
    $Grid.DefaultCellStyle.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $Grid.DefaultCellStyle.SelectionBackColor = $script:Ui.GridSelect
    $Grid.DefaultCellStyle.SelectionForeColor = $script:Ui.Text
    $Grid.AlternatingRowsDefaultCellStyle.BackColor = $script:Ui.SurfaceAlt
    $Grid.AlternatingRowsDefaultCellStyle.ForeColor = $script:Ui.Text
    $Grid.ColumnHeadersDefaultCellStyle.BackColor = $script:Ui.GridHeader
    $Grid.ColumnHeadersDefaultCellStyle.ForeColor = [System.Drawing.Color]::White
    $Grid.ColumnHeadersDefaultCellStyle.Font = New-Object System.Drawing.Font("Segoe UI", 8.5, [System.Drawing.FontStyle]::Bold)
    $Grid.ColumnHeadersDefaultCellStyle.SelectionBackColor = $script:Ui.GridHeader
}

function Initialize-N2Button {
    param(
        [System.Windows.Forms.Button]$Button,
        [ValidateSet("Primary", "Secondary", "Ghost")][string]$Kind
    )
    $Button.FlatStyle = [System.Windows.Forms.FlatStyle]::Flat
    $Button.FlatAppearance.BorderSize = 0
    $Button.Cursor = [System.Windows.Forms.Cursors]::Hand
    $Button.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::SemiBold)
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
    $net2 = if ($script:Connected) { "Connected" } else { "Disconnected" }
    $lblDry.Text = "Safe mode: no live writes to NET2 from this build  |  SQL: $net2"
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

# --- Form ---
$form = New-Object System.Windows.Forms.Form
$form.Text = "Merdian Paxton Bridge"
$form.Size = New-Object System.Drawing.Size(1160, 760)
$form.MinimumSize = New-Object System.Drawing.Size(920, 580)
$form.StartPosition = "CenterScreen"
$form.BackColor = $script:Ui.FormBg
$form.Font = New-Object System.Drawing.Font("Segoe UI", 9)
$lblTitle = New-Object System.Windows.Forms.Label
$lblTitle.AutoSize = $true
$lblTitle.Location = New-Object System.Drawing.Point(16, 14)
$lblTitle.Font = New-Object System.Drawing.Font("Segoe UI", 18, [System.Drawing.FontStyle]::Bold)
$lblTitle.ForeColor = $script:Ui.Text
$lblTitle.Text = "Merdian Paxton Bridge"

$lblSub = New-Object System.Windows.Forms.Label
$lblSub.AutoSize = $true
$lblSub.Location = New-Object System.Drawing.Point(16, 46)
$lblSub.ForeColor = $script:Ui.Muted
$lblSub.Font = New-Object System.Drawing.Font("Segoe UI", 9.25)
$lblSub.Text = "CSV drops, Paxton NET2 checks, renewals - bridge UI (mock data only)."

$pnlAccent = New-Object System.Windows.Forms.Panel
$pnlAccent.Location = New-Object System.Drawing.Point(16, 70)
$pnlAccent.Size = New-Object System.Drawing.Size(72, 3)
$pnlAccent.BackColor = $script:Ui.Accent

$tabControl = New-Object System.Windows.Forms.TabControl
$tabControl.Location = New-Object System.Drawing.Point(16, 82)
$tabControl.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$tabControl.Size = New-Object System.Drawing.Size(($form.ClientSize.Width - 32), ($form.ClientSize.Height - 94))
$tabControl.Padding = New-Object System.Windows.Forms.Padding(6, 8, 6, 6)
$tabControl.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::SemiBold)
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
$tabConfig.BackColor = $script:Ui.Surface
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
$lblDry.Font = New-Object System.Drawing.Font("Segoe UI", 9.25, [System.Drawing.FontStyle]::SemiBold)
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
$lblServiceLogs.Font = New-Object System.Drawing.Font("Segoe UI", 8.75, [System.Drawing.FontStyle]::SemiBold)
$lblServiceLogs.Text = "Recent activity (newest first, up to 200 lines)"

$dgServiceLog = New-Object System.Windows.Forms.DataGridView
$dgServiceLog.ReadOnly = $true
$dgServiceLog.AllowUserToAddRows = $false
$dgServiceLog.AllowUserToDeleteRows = $false
$dgServiceLog.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
$dgServiceLog.MultiSelect = $false
$dgServiceLog.AutoSizeColumnsMode = [System.Windows.Forms.DataGridViewAutoSizeColumnsMode]::Fill
$dgServiceLog.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Bottom -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
$dgServiceLog.Location = New-Object System.Drawing.Point(16, 100)
$dgServiceLog.Size = New-Object System.Drawing.Size(800, 400)
foreach ($c in @("When", "Event", "Detail")) {
    $null = $dgServiceLog.Columns.Add($c, $c)
}
$dgServiceLog.Columns["When"].FillWeight = 22
$dgServiceLog.Columns["Event"].FillWeight = 20
$dgServiceLog.Columns["Detail"].FillWeight = 58
Apply-N2DataGridView -Grid $dgServiceLog

$tabService.Controls.AddRange(@(
    $pnlServiceStrip, $lblServiceLogs, $dgServiceLog
))

# --- Configuration tab ---
function New-CfgRow {
    param([int]$Y, [string]$LabelText, [string]$BoxText, [int]$BoxWidth)
    $lab = New-Object System.Windows.Forms.Label
    $lab.AutoSize = $true
    $lab.Location = New-Object System.Drawing.Point(16, $Y)
    $lab.Text = $LabelText
    $lab.ForeColor = $script:Ui.Muted
    $lab.Font = New-Object System.Drawing.Font("Segoe UI", 9)
    $tb = New-Object System.Windows.Forms.TextBox
    $tb.Location = New-Object System.Drawing.Point(208, ($Y - 2))
    $tb.Width = $BoxWidth
    $tb.Text = $BoxText
    $tb.ReadOnly = $true
    $tb.BorderStyle = [System.Windows.Forms.BorderStyle]::FixedSingle
    $tb.BackColor = $script:Ui.SurfaceAlt
    $tb.ForeColor = $script:Ui.Text
    $tb.Font = New-Object System.Drawing.Font("Consolas", 9)
    $tb.Anchor = [System.Windows.Forms.AnchorStyles]::Top -bor [System.Windows.Forms.AnchorStyles]::Left -bor [System.Windows.Forms.AnchorStyles]::Right
    return @($lab, $tb)
}

$yCfg = 20
$pairA = New-CfgRow -Y $yCfg -LabelText "Watch folder" -BoxText "D:\Inbound\NET2Sync\" -BoxWidth 520
$yCfg += 36
$pairB = New-CfgRow -Y $yCfg -LabelText "File pattern" -BoxText "*.csv" -BoxWidth 200
$yCfg += 36
$pairC = New-CfgRow -Y $yCfg -LabelText "Expected drop interval (minutes)" -BoxText "60" -BoxWidth 80
$yCfg += 36
$pairD = New-CfgRow -Y $yCfg -LabelText "NET2 SQL server" -BoxText "(not configured)" -BoxWidth 520
$yCfg += 36
$pairE = New-CfgRow -Y $yCfg -LabelText "NET2 database name" -BoxText "Net2" -BoxWidth 200
$yCfg += 36
$pairF = New-CfgRow -Y $yCfg -LabelText "GUI number column in CSV" -BoxText "GuiNumber" -BoxWidth 200

$lblCfgNote = New-Object System.Windows.Forms.Label
$lblCfgNote.AutoSize = $false
$lblCfgNote.Location = New-Object System.Drawing.Point(16, ($yCfg + 32))
$lblCfgNote.Size = New-Object System.Drawing.Size(700, 52)
$lblCfgNote.ForeColor = $script:Ui.Muted
$lblCfgNote.Font = New-Object System.Drawing.Font("Segoe UI", 9, [System.Drawing.FontStyle]::Italic)
$lblCfgNote.Text = "These values are placeholders. Saving to disk and real SQL connections will be added in a later step."

$tabConfig.Controls.AddRange(@(
    $pairA[0], $pairA[1], $pairB[0], $pairB[1], $pairC[0], $pairC[1],
    $pairD[0], $pairD[1], $pairE[0], $pairE[1], $pairF[0], $pairF[1], $lblCfgNote
))

$tabConfig.add_Resize({
    $rw = $tabConfig.ClientSize.Width - 236
    if ($rw -lt 200) { $rw = 200 }
    $pairA[1].Width = $rw
    $pairD[1].Width = $rw
})

# --- History tab ---
$lblHistHint = New-Object System.Windows.Forms.Label
$lblHistHint.AutoSize = $true
$lblHistHint.Location = New-Object System.Drawing.Point(16, 14)
$lblHistHint.ForeColor = $script:Ui.Muted
$lblHistHint.Font = New-Object System.Drawing.Font("Segoe UI", 8.75)
$lblHistHint.Text = "Complete archive (mock). The Service tab shows only the latest 200 lines."

$dgHistory = New-Object System.Windows.Forms.DataGridView
$dgHistory.ReadOnly = $true
$dgHistory.AllowUserToAddRows = $false
$dgHistory.AllowUserToDeleteRows = $false
$dgHistory.SelectionMode = [System.Windows.Forms.DataGridViewSelectionMode]::FullRowSelect
$dgHistory.MultiSelect = $false
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
$lblSupTitle.Font = New-Object System.Drawing.Font("Segoe UI", 12, [System.Drawing.FontStyle]::Bold)
$lblSupTitle.ForeColor = $script:Ui.Text
$lblSupTitle.Location = New-Object System.Drawing.Point(16, 18)
$lblSupTitle.Text = "About Merdian Paxton Bridge"

$lblSupBody = New-Object System.Windows.Forms.Label
$lblSupBody.AutoSize = $false
$lblSupBody.Location = New-Object System.Drawing.Point(16, 52)
$lblSupBody.Size = New-Object System.Drawing.Size(720, 220)
$lblSupBody.ForeColor = $script:Ui.Text
$lblSupBody.Font = New-Object System.Drawing.Font("Segoe UI", 9.25)
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
    $dgServiceLog.Width = [Math]::Max(200, $tabService.ClientSize.Width - 32)
    $dgServiceLog.Height = [Math]::Max(140, $tabService.ClientSize.Height - $dgServiceLog.Top - 20)
})

$btnConn.add_Click({
    $script:Connected = -not $script:Connected
    if ($script:Connected) {
        Add-HistoryEntry -Event "NET2 SQL" -Detail "Connected (mock - no real server)"
    } else {
        Add-HistoryEntry -Event "NET2 SQL" -Detail "Disconnected"
    }
    Refresh-ServiceState -lblDry $lblDry -btnConn $btnConn
    Refresh-AllLogGrids -GridService $dgServiceLog -GridHistory $dgHistory
})

$btnManualSync.add_Click({
    $detail = if ($script:Connected) {
        "Watch folder poll (mock) - CSV and NET2 rules not wired yet"
    } else {
        "UI-only tick (mock) - connect NET2 first when SQL step exists"
    }
    Add-HistoryEntry -Event "Manual sync" -Detail $detail
    Refresh-ServiceState -lblDry $lblDry -btnConn $btnConn
    Refresh-AllLogGrids -GridService $dgServiceLog -GridHistory $dgHistory
})

$btnClearLogs.add_Click({
    $script:History.Clear()
    Refresh-ServiceState -lblDry $lblDry -btnConn $btnConn
    Refresh-AllLogGrids -GridService $dgServiceLog -GridHistory $dgHistory
})

$form.Controls.AddRange(@($lblTitle, $lblSub, $pnlAccent, $tabControl))

$pnlServiceStrip.Width = [Math]::Max(480, $tabService.ClientSize.Width - 32)
$lblDry.MaximumSize = New-Object System.Drawing.Size([Math]::Max(200, $pnlServiceStrip.ClientSize.Width - 400), 40)
Sync-TopRightButtons -Parent $pnlServiceStrip -ConnBtn $btnConn -ManualSyncBtn $btnManualSync -ClearBtn $btnClearLogs
Refresh-ServiceState -lblDry $lblDry -btnConn $btnConn
Refresh-AllLogGrids -GridService $dgServiceLog -GridHistory $dgHistory
[void]$form.ShowDialog()
