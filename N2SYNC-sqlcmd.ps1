# Run one-off read-only SQL against Paxton Net2 using sqlcmd (same style as: sqlcmd -S ".\NET2" -d Net2 -E).
# Requires sqlcmd on PATH (SQL Server Command Line Utilities).
# Example:
#   powershell -ExecutionPolicy Bypass -File ".\N2SYNC-sqlcmd.ps1" -Query "SELECT TOP 5 name FROM sys.tables WHERE name IN ('Users','Cards') ORDER BY name"

param(
    [string]$Server = ".\NET2",
    [string]$Database = "Net2",
    [string]$Query = ""
)

$sqlcmdPath = @(
    (Get-Command sqlcmd -ErrorAction SilentlyContinue | Select-Object -ExpandProperty Source),
    "${env:ProgramFiles}\Microsoft SQL Server\Client SDK\ODBC\170\Tools\Binn\sqlcmd.exe",
    "${env:ProgramFiles(x86)}\Microsoft SQL Server\Client SDK\ODBC\170\Tools\Binn\sqlcmd.exe"
) | Where-Object { $_ -and (Test-Path -LiteralPath $_) } | Select-Object -First 1

if (-not $sqlcmdPath) {
    Write-Host "sqlcmd.exe not found. Install SQL Server Command Line Utilities, or add sqlcmd to PATH."
    Write-Host "Equivalent to: sqlcmd -S `"$Server`" -d $Database -E -Q `"<your SELECT>`""
    exit 1
}

if ([string]::IsNullOrWhiteSpace($Query)) {
    $Query = @"
SELECT TOP 40 c.name, ty.name AS typ, c.max_length
FROM sys.columns c
INNER JOIN sys.types ty ON ty.user_type_id = c.user_type_id
WHERE c.object_id = OBJECT_ID(N'dbo.Users')
ORDER BY c.column_id;
"@
}

Write-Host "Running: $sqlcmdPath -S `"$Server`" -d $Database -E -Q `"...`""
Write-Host "-----"
$args = @("-S", $Server, "-d", $Database, "-E", "-b", "-Q", $Query)
$p = Start-Process -FilePath $sqlcmdPath -ArgumentList $args -NoNewWindow -Wait -PassThru
exit $p.ExitCode
