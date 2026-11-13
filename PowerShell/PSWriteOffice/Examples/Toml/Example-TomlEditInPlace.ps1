Import-Module PSWriteOffice -ErrorAction Stop
$documents = Join-Path $PSScriptRoot '..\Documents'
$null = New-Item -Path $documents -ItemType Directory -Force
$project = Join-Path $documents 'Toml-pyproject.toml'

# A Python project file with the comments and alignment a hand-maintained file would have.
@'
# Build settings for the reporting tool
[project]
name    = "reporting"
version = "1.1.0"   # bumped by the release script
dependencies = [
    "httpx>=0.26",
    "rich",
]

[tool.ruff]
line-length = 88

[tool.ruff.lint]
ignore = ["E501"]
'@ | Set-Content -LiteralPath $project -NoNewline

# Edit in place: only the edited values change when the block saves.
Open-OfficeToml -Path $project {
    Set-OfficeTomlNode -Select 'project.version' -Value '1.2.0'
    Set-OfficeTomlNode -Select 'project.dependencies[0]' -Value 'httpx>=0.27'
    Set-OfficeTomlNode -Select 'tool.ruff.line-length' -Value 100
    Remove-OfficeTomlNode -Select 'tool.ruff.lint.ignore'
}
Get-Content -LiteralPath $project -Raw

# Read values with their positions in the file.
$document = Open-OfficeToml -Path $project
Get-OfficeTomlNode -Document $document -Select 'project.version', 'tool.ruff.line-length' |
    Select-Object -Property Path, Value, Line, Column
Close-OfficeToml -Document $document
