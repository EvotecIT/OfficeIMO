<#
.SYNOPSIS
Builds the compact removed-command hints from the approved API map.
.DESCRIPTION
Writes an ordered JSON object keyed by old command name. Existing commands keep working;
hints are only displayed when normal command lookup fails. Map files remain the source.
.PARAMETER MapDirectory
Directory containing the PowerShell command maps.
.PARAMETER OutputPath
File to write. Its parent directory is created when needed.
.EXAMPLE
./New-CommandHintResource.ps1 -MapDirectory ../../../Build/ApiMap/powershell -OutputPath ../Sources/PSWriteOffice/obj/OfficeCommandHints.json
.EXAMPLE
./New-CommandHintResource.ps1 -MapDirectory /src/OfficeIMO/Build/ApiMap/powershell -OutputPath /tmp/PSWriteOffice/OfficeCommandHints.json
.OUTPUTS
None. Writes the JSON resource consumed by module import.
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)][string] $MapDirectory,
    [Parameter(Mandatory)][string] $OutputPath
)
$ErrorActionPreference = 'Stop'
$hints = [ordered]@{}
$files = @(Get-ChildItem -LiteralPath $MapDirectory -Filter '*.json' | Sort-Object Name)
if (!$files.Count) { throw "No command map files in $MapDirectory." }
foreach ($file in $files) {
    $data = Get-Content -LiteralPath $file.FullName -Raw | ConvertFrom-Json
    foreach ($entry in $data.map) {
        if ($entry.action -eq 'keep' -or !$entry.hint) { continue }
        if ($hints.Contains($entry.old)) { throw "Duplicate command hint: $($entry.old)" }
        $hints[$entry.old] = [string]$entry.hint
    }
}
$fullPath = [IO.Path]::GetFullPath($OutputPath)
$null = [IO.Directory]::CreateDirectory([IO.Path]::GetDirectoryName($fullPath))
[IO.File]::WriteAllText($fullPath, ($hints | ConvertTo-Json -Compress -Depth 4), [Text.UTF8Encoding]::new($false))
