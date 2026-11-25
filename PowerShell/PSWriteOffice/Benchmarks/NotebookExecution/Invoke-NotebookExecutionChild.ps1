param(
    [ValidateSet('OfficeIMO', 'PowerShellNotebook')]
    [string] $Engine,
    [string] $ModuleManifest,
    [string] $InputPath,
    [string] $OutputPath,
    [string] $PowerShellPath
)

$ErrorActionPreference = 'Stop'
Import-Module -Name $ModuleManifest -WarningAction SilentlyContinue
if ($Engine -eq 'OfficeIMO') {
    $report = PSWriteOffice\Invoke-OfficeNotebook -LiteralPath $InputPath -OutputPath $OutputPath -PowerShellPath $PowerShellPath -Confirm:$false
    if (-not $report.Succeeded) {
        throw 'Invoke-OfficeNotebook reported an unsuccessful execution.'
    }
} else {
    PowerShellNotebook\Invoke-PowerShellNotebook -NoteBookFullName $InputPath -OutFile $OutputPath
}
