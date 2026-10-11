<#
.SYNOPSIS
Builds the patch compliance report shown on the PSWriteOffice product page.

.DESCRIPTION
Creates a Word report from PowerShell objects with PSWriteOffice, then exports it to PDF.
The site build runs this file with the module version its command catalog was generated from,
so the downloadable DOCX and PDF come from the code shown beside them.

.EXAMPLE
pwsh -NoProfile -File .\PatchComplianceReport.ps1 -OutputDirectory .\out
#>
[CmdletBinding()]
param(
    [Parameter(Mandatory)][string] $OutputDirectory,
    [string] $ModuleVersion
)

$ErrorActionPreference = 'Stop'
if ($ModuleVersion) { Import-Module PSWriteOffice -RequiredVersion $ModuleVersion } else { Import-Module PSWriteOffice }
New-Item -ItemType Directory -Path $OutputDirectory -Force | Out-Null
$docxPath = Join-Path $OutputDirectory 'PatchComplianceReport.docx'
$pdfPath = Join-Path $OutputDirectory 'PatchComplianceReport.pdf'

#region Excerpt:pswriteoffice-patch-report
# Sample inventory; replace it with your own query. Any objects with properties work.
$sites = @(
    [pscustomobject]@{ Site = 'Warsaw';    Servers = 84; Patched = 80; Pending = 4; Status = 'On target' }
    [pscustomobject]@{ Site = 'Frankfurt'; Servers = 66; Patched = 58; Pending = 8; Status = 'Watch' }
    [pscustomobject]@{ Site = 'Dublin';    Servers = 52; Patched = 50; Pending = 2; Status = 'On target' }
    [pscustomobject]@{ Site = 'Chicago';   Servers = 71; Patched = 69; Pending = 2; Status = 'On target' }
    [pscustomobject]@{ Site = 'Singapore'; Servers = 38; Patched = 33; Pending = 5; Status = 'Watch' }
)
$servers = ($sites | Measure-Object Servers -Sum).Sum
$patched = ($sites | Measure-Object Patched -Sum).Sum
$rows = $sites | Select-Object Site, Servers, Patched, Pending,
    @{ Name = 'Compliance'; Expression = { "$([math]::Round(100 * $_.Patched / $_.Servers, 1))%" } }, Status

New-OfficeWord -Path $docxPath {
    WordHeader -Type Default { WordParagraph -Text 'Contoso IT operations - week 41' }
    WordFooter -Type Default { WordPageNumber -IncludeTotalPages }

    WordParagraph { WordText 'Patch compliance report' -Bold -FontSize 26 -Color '17365D' }
    WordParagraph {
        WordText "$([math]::Round(100 * $patched / $servers, 1))%" -Bold -FontSize 20 -Color '166534'
        WordText "  of $servers production servers are patched, $($servers - $patched) are pending." -FontSize 12
    }

    WordParagraph -Text 'Compliance by site' -Style Heading1
    WordTable -InputObject $rows -Style GridTable4Accent1 {
        WordTableCondition -FilterScript { $_.Status -eq 'Watch' } -BackgroundColor '#fff4d6'
    }
    WordChart -Type Bar -InputObject $rows -CategoryProperty Site -SeriesProperty Patched, Pending `
        -Title 'Patched and pending by site' -Legend -FitToPageWidth -WidthFraction 0.9

    WordParagraph -Text 'Next steps' -Style Heading1
    WordList -Style Numbered {
        WordListItem -Text 'Frankfurt: patch the eight pending servers in the Thursday window.'
        WordListItem -Text 'Singapore: confirm the maintenance slot with the regional owner.'
    }
}
Export-OfficeDocumentPdf -InputPath $docxPath -Path $pdfPath
#endregion

Get-Item -LiteralPath $docxPath, $pdfPath | Select-Object Name, Length
