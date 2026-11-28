param(
    [ValidateSet('OfficeIMO', 'OfficeIMOProcess', 'PowerShellNotebook')]
    [string] $Engine,
    [string] $ModuleManifest,
    [string] $InputPath,
    [string] $OutputPath,
    [string] $PowerShellPath
)

$ErrorActionPreference = 'Stop'
Import-Module -Name $ModuleManifest -WarningAction SilentlyContinue
if ($Engine -like 'OfficeIMO*') {
    $isolation = @{}
    if ($Engine -eq 'OfficeIMOProcess') {
        $isolation.PowerShellPath = $PowerShellPath
    }
    $report = PSWriteOffice\Invoke-OfficeNotebook -LiteralPath $InputPath -OutputPath $OutputPath @isolation -Confirm:$false
    if (-not $report.Succeeded) {
        throw 'Invoke-OfficeNotebook reported an unsuccessful execution.'
    }
} else {
    PowerShellNotebook\Invoke-PowerShellNotebook -NoteBookFullName $InputPath -OutFile $OutputPath
}
