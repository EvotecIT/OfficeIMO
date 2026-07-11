<#
.SYNOPSIS
Opens generated ODT, ODS, and ODP artifacts with installed Microsoft Office desktop applications.

.DESCRIPTION
Runs an opt-in Windows interoperability smoke test through Word, Excel, and PowerPoint COM automation. Files are opened read-only with macros disabled and are not added to recent-file lists. This script reports automation failures; it does not replace an interactive repair-prompt check before release.

.PARAMETER ArtifactPath
Directory containing the generated OpenDocument package artifacts.

.EXAMPLE
./Build/Test-OpenDocumentMicrosoftOffice.ps1 -ArtifactPath ./Artifacts/OpenDocument
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)]
    [string] $ArtifactPath
)

$resolvedPath = (Resolve-Path -LiteralPath $ArtifactPath -ErrorAction Stop).Path
$files = @(Get-ChildItem -LiteralPath $resolvedPath -File | Where-Object Extension -In '.odt', '.ods', '.odp')
if ($files.Count -eq 0) {
    throw "No ODT, ODS, or ODP files were found in '$resolvedPath'."
}

function Close-ComObject {
    param([object] $InputObject)
    if ($null -ne $InputObject -and [System.Runtime.InteropServices.Marshal]::IsComObject($InputObject)) {
        [void] [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($InputObject)
    }
}

$word = $null
$excel = $null
$powerPoint = $null
try {
    $odtFiles = @($files | Where-Object Extension -EQ '.odt')
    if ($odtFiles.Count -gt 0) {
        $word = New-Object -ComObject Word.Application
        $word.Visible = $false
        $word.DisplayAlerts = 0
        $word.AutomationSecurity = 3
        foreach ($file in $odtFiles) {
            $document = $null
            try {
                $document = $word.Documents.Open($file.FullName, $false, $true, $false)
            } finally {
                if ($null -ne $document) { $document.Close(0) }
                Close-ComObject $document
            }
        }
    }

    $odsFiles = @($files | Where-Object Extension -EQ '.ods')
    if ($odsFiles.Count -gt 0) {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        $excel.AutomationSecurity = 3
        foreach ($file in $odsFiles) {
            $workbook = $null
            try {
                $workbook = $excel.Workbooks.Open($file.FullName, 0, $true)
            } finally {
                if ($null -ne $workbook) { $workbook.Close($false) }
                Close-ComObject $workbook
            }
        }
    }

    $odpFiles = @($files | Where-Object Extension -EQ '.odp')
    if ($odpFiles.Count -gt 0) {
        $powerPoint = New-Object -ComObject PowerPoint.Application
        $powerPoint.AutomationSecurity = 3
        foreach ($file in $odpFiles) {
            $presentation = $null
            try {
                $presentation = $powerPoint.Presentations.Open($file.FullName, $true, $true, $false)
            } finally {
                if ($null -ne $presentation) { $presentation.Close() }
                Close-ComObject $presentation
            }
        }
    }

    Write-Host "Microsoft Office opened $($files.Count) OpenDocument files without an automation error."
} finally {
    if ($null -ne $word) { $word.Quit(0) }
    if ($null -ne $excel) { $excel.Quit() }
    if ($null -ne $powerPoint) { $powerPoint.Quit() }
    Close-ComObject $word
    Close-ComObject $excel
    Close-ComObject $powerPoint
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
}
