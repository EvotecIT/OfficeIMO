#Requires -Version 7.0
<#
.SYNOPSIS
Writes a nullable public API baseline from the compiler's analyzer SARIF output.
.DESCRIPTION
Run the public API analyzer against empty baseline files with ErrorLog enabled, then
pass that SARIF file here. Only RS0016 declarations are used. Existing source and
other diagnostics are unchanged. Build again to verify the generated baseline.
.EXAMPLE
./Build/ApiMap/New-PublicApiBaseline.ps1 -SarifPath ./api.sarif -OutputPath ./PublicAPI.Shipped.txt
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)][string] $SarifPath,
    [Parameter(Mandatory)][string] $OutputPath
)
Set-StrictMode -Version 3.0
$ErrorActionPreference = 'Stop'
$inputLog = Get-Content -LiteralPath $SarifPath -Raw | ConvertFrom-Json
$declarations = [System.Collections.Generic.HashSet[string]]::new([StringComparer]::Ordinal)
$prefix = "Symbol '"
$suffix = "' is not part of the declared public API"
foreach ($run in $inputLog.runs) {
    foreach ($result in $run.results) {
        if ($result.ruleId -ne 'RS0016') { continue }
        $message = if ($result.message -is [string]) { [string] $result.message } else { [string] $result.message.text }
        if (-not $message.StartsWith($prefix, [StringComparison]::Ordinal) -or
            -not $message.EndsWith($suffix, [StringComparison]::Ordinal)) {
            throw "Unrecognized RS0016 declaration: $message"
        }
        $declaration = $message.Substring($prefix.Length, $message.Length - $prefix.Length - $suffix.Length)
        if ($declaration.Contains("`r") -or $declaration.Contains("`n")) {
            throw 'A public API declaration contains an unexpected line break.'
        }
        [void] $declarations.Add($declaration)
    }
}
if ($declarations.Count -eq 0) { throw 'The analyzer log contains no RS0016 declarations.' }
$sorted = [string[]] @($declarations)
[Array]::Sort($sorted, [StringComparer]::Ordinal)
$destination = [IO.Path]::GetFullPath($OutputPath)
[void] [IO.Directory]::CreateDirectory([IO.Path]::GetDirectoryName($destination))
[IO.File]::WriteAllText($destination, "#nullable enable`n" + ($sorted -join "`n") + "`n", [Text.UTF8Encoding]::new($false))
[pscustomobject] @{ Path = $destination; Declarations = $sorted.Length }
