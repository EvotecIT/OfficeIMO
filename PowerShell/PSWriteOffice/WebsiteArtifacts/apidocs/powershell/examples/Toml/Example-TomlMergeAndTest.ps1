Import-Module PSWriteOffice -ErrorAction Stop
$documents = Join-Path $PSScriptRoot '..\Documents'
$null = New-Item -Path $documents -ItemType Directory -Force
$shared = Join-Path $documents 'Toml-pyproject.shared.toml'
$local = Join-Path $documents 'Toml-pyproject.local.toml'
$merged = Join-Path $documents 'Toml-pyproject.merged.toml'

# Settings the team shares, with comments a hand-maintained file would have.
@'
# Shared settings for the reporting tool
[project]
name = "reporting"
dependencies = ["httpx>=0.27", "rich"]

[tool.ruff]
line-length = 88 # team default

[[tool.reporting.plugin]]
name = "lint"
version = "1.0"
'@ | Set-Content -LiteralPath $shared -NoNewline

# Local overrides: a longer line length, an extra lint rule and an updated plugin.
@'
[tool.ruff]
line-length = 100

[tool.ruff.lint]
select = ["E", "F"]

[[tool.reporting.plugin]]
name = "lint"
version = "2.0"

[[tool.reporting.plugin]]
name = "docs"
version = "1.0"
'@ | Set-Content -LiteralPath $local -NoNewline

# Check both files before merging; -Detailed says where a broken file fails.
Test-OfficeToml -Path $shared, $local -Detailed | Select-Object -Property Path, IsValid, Line, Column, Code

# Merge: tables merge key by key, [[plugin]] tables match by name, and the shared file keeps its comments.
$report = Merge-OfficeDocument -Path $shared, $local -OutputPath $merged -Options @{ Arrays = @{ MergeByKey = 'name' } } -PassThruReport
$report | Select-Object -Property Added, Replaced
Get-Content -LiteralPath $merged -Raw
