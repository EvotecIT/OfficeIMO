Import-Module PSWriteOffice -ErrorAction Stop
$documents = Join-Path $PSScriptRoot '..\Documents'
$null = New-Item -Path $documents -ItemType Directory -Force
$lock = Join-Path $documents 'Data-Cargo.lock.toml'
$packages = Join-Path $documents 'Data-Packages.csv'
$json = Join-Path $documents 'Data-Packages.json'
$settings = Join-Path $documents 'Data-Settings.toml'

# A Cargo lock file: one [[package]] table per dependency.
@'
version = 4

[[package]]
name = "serde"
version = "1.0.210"
checksum = "c8e3592472072e6e22e0a54d5904d9febf8508f65fb8552499a1abc7d1078c3a"

[[package]]
name = "toml"
version = "0.8.19"
'@ | Set-Content -LiteralPath $lock -NoNewline

# Read it as objects, as ConvertFrom-Json does for JSON.
(Import-OfficeToml -Path $lock).package | Format-Table -Property name, version

# Convert files between data formats; the output extension selects the target. Select picks the rows.
Convert-OfficeDocument -Path $lock -OutputPath $packages -Options @{ Select = 'package' } | Format-Table -Property Target, ItemsWritten, HasLoss
Convert-OfficeDocument -Path $packages -OutputPath $json | Out-Null
Get-Content -LiteralPath $json -Raw

# Write settings as TOML: nested objects become [table] sections; nulls are omitted with a warning.
[ordered]@{
    title = 'Nightly report'
    retries = 3
    owner = [ordered]@{ team = 'Ops'; updated = [DateTimeOffset]'2026-11-11T09:30:00Z' }
} | Export-OfficeToml -Path $settings
Get-Content -LiteralPath $settings -Raw
